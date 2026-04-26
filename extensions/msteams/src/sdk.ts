import * as fs from "node:fs";
// IHttpServerAdapter is re-exported via the public barrel (`export * from './http'`)
// but tsgo cannot resolve the chain. Use the dist subpath directly (type-only import).
import type { IHttpServerAdapter } from "@microsoft/teams.apps/dist/http/index.js";
import type { MSTeamsCredentials, MSTeamsFederatedCredentials } from "./token.js";
import { buildUserAgent } from "./user-agent.js";

/**
 * Resolved Teams SDK modules loaded lazily to avoid importing when the
 * provider is disabled.
 */
export type MSTeamsTeamsSdk = {
  App: typeof import("@microsoft/teams.apps").App;
  Client: typeof import("@microsoft/teams.api").Client;
  ExpressAdapter: typeof import("@microsoft/teams.apps").ExpressAdapter;
};

/**
 * A Teams SDK App instance used for token management and proactive messaging.
 */
export type MSTeamsApp = InstanceType<MSTeamsTeamsSdk["App"]>;

/**
 * Token provider compatible with the existing codebase, wrapping the Teams
 * SDK App's public tokenManager.
 */
export type MSTeamsTokenProvider = {
  getAccessToken: (scope: string) => Promise<string>;
};

type AzureAccessToken = {
  token?: string;
} | null;

type AzureTokenCredential = {
  getToken: (scope: string | string[]) => Promise<AzureAccessToken>;
};

type AzureIdentityModule = {
  ClientCertificateCredential: new (
    tenantId: string,
    clientId: string,
    options: { certificate: string },
  ) => AzureTokenCredential;
  ManagedIdentityCredential: new (clientId?: string) => AzureTokenCredential;
};

const AZURE_IDENTITY_MODULE = "@azure/identity";

let azureIdentityModulePromise: Promise<AzureIdentityModule> | null = null;

async function loadAzureIdentity(): Promise<AzureIdentityModule> {
  azureIdentityModulePromise ??= import(AZURE_IDENTITY_MODULE) as Promise<AzureIdentityModule>;
  return azureIdentityModulePromise;
}

let msTeamsSdkPromise: Promise<MSTeamsTeamsSdk> | null = null;

export async function loadMSTeamsSdk(): Promise<MSTeamsTeamsSdk> {
  msTeamsSdkPromise ??= Promise.all([
    import("@microsoft/teams.apps"),
    import("@microsoft/teams.api"),
  ]).then(([appsModule, apiModule]) => ({
    App: appsModule.App,
    Client: apiModule.Client,
    ExpressAdapter: appsModule.ExpressAdapter,
  }));
  return msTeamsSdkPromise;
}

/**
 * Create a no-op HTTP server adapter for non-server scenarios (probes,
 * proactive-only CLI sends) where no Express server exists.
 */
function createNoOpHttpServerAdapter(): IHttpServerAdapter {
  return {
    registerRoute() {},
  };
}

/**
 * Options for creating a Teams SDK App instance.
 */
export type CreateMSTeamsAppOptions = {
  /**
   * HTTP server adapter to use. When an Express app is available (monitor
   * mode), pass an ExpressAdapter so the SDK registers routes and handles
   * JWT validation. When omitted, a no-op adapter is used (probe/CLI mode).
   */
  httpServerAdapter?: IHttpServerAdapter;
  /**
   * Custom messaging endpoint path.
   * @default '/api/messages'
   */
  messagingEndpoint?: `/${string}`;
};

/**
 * Create a Teams SDK App instance from credentials. The App manages token
 * acquisition, JWT validation, and the HTTP server lifecycle.
 */
export async function createMSTeamsApp(
  creds: MSTeamsCredentials,
  sdk: MSTeamsTeamsSdk,
  options?: CreateMSTeamsAppOptions,
): Promise<MSTeamsApp> {
  const adapter = options?.httpServerAdapter ?? createNoOpHttpServerAdapter();
  const messagingEndpoint = options?.messagingEndpoint;

  if (creds.type === "federated") {
    return createFederatedApp(creds, sdk, adapter, messagingEndpoint);
  }
  return new sdk.App({
    clientId: creds.appId,
    clientSecret: creds.appPassword,
    tenantId: creds.tenantId,
    httpServerAdapter: adapter,
    ...(messagingEndpoint ? { messagingEndpoint } : {}),
  } as ConstructorParameters<MSTeamsTeamsSdk["App"]>[0]);
}

function createFederatedApp(
  creds: MSTeamsFederatedCredentials,
  sdk: MSTeamsTeamsSdk,
  adapter: IHttpServerAdapter,
  messagingEndpoint?: `/${string}`,
): MSTeamsApp {
  if (creds.useManagedIdentity) {
    return createManagedIdentityApp(creds, sdk, adapter, messagingEndpoint);
  }

  // Certificate-based auth
  if (!creds.certificatePath) {
    throw new Error("Federated credentials require either a certificate path or managed identity.");
  }

  let privateKey: string;
  try {
    privateKey = fs.readFileSync(creds.certificatePath, "utf-8");
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`Failed to read certificate file at '${creds.certificatePath}': ${msg}`, {
      cause: err,
    });
  }

  return createCertificateApp(creds, privateKey, sdk, adapter, messagingEndpoint);
}

function createCertificateApp(
  creds: MSTeamsFederatedCredentials,
  privateKey: string,
  sdk: MSTeamsTeamsSdk,
  adapter: IHttpServerAdapter,
  messagingEndpoint?: `/${string}`,
): MSTeamsApp {
  // Lazily create and cache the credential so the token cache is reused.
  let credentialPromise: Promise<AzureTokenCredential> | null = null;

  const getCredential = async () => {
    if (!credentialPromise) {
      credentialPromise = loadAzureIdentity().then(
        (az) =>
          new az.ClientCertificateCredential(creds.tenantId, creds.appId, {
            certificate: privateKey,
          }),
      );
    }
    return credentialPromise;
  };

  const tokenProvider = async (scope: string | string[]): Promise<string> => {
    const credential = await getCredential();
    const token = await credential.getToken(scope);

    if (!token?.token) {
      throw new Error("Failed to acquire token via certificate credential.");
    }

    return token.token;
  };

  return new sdk.App({
    clientId: creds.appId,
    tenantId: creds.tenantId,
    token: tokenProvider,
    httpServerAdapter: adapter,
    ...(messagingEndpoint ? { messagingEndpoint } : {}),
  } as unknown as ConstructorParameters<MSTeamsTeamsSdk["App"]>[0]);
}

function createManagedIdentityApp(
  creds: MSTeamsFederatedCredentials,
  sdk: MSTeamsTeamsSdk,
  adapter: IHttpServerAdapter,
  messagingEndpoint?: `/${string}`,
): MSTeamsApp {
  // Lazily create and cache the credential instance so the token cache is
  // reused across calls instead of hitting IMDS/AAD on every message.
  let credentialPromise: Promise<AzureTokenCredential> | null = null;

  const getCredential = async () => {
    if (!credentialPromise) {
      credentialPromise = loadAzureIdentity().then((az) =>
        creds.managedIdentityClientId
          ? new az.ManagedIdentityCredential(creds.managedIdentityClientId)
          : new az.ManagedIdentityCredential(),
      );
    }
    return credentialPromise;
  };

  const tokenProvider = async (scope: string | string[]): Promise<string> => {
    const credential = await getCredential();
    const token = await credential.getToken(scope);

    if (!token?.token) {
      throw new Error("Failed to acquire token via managed identity.");
    }

    return token.token;
  };

  return new sdk.App({
    clientId: creds.appId,
    tenantId: creds.tenantId,
    token: tokenProvider,
    httpServerAdapter: adapter,
    ...(messagingEndpoint ? { messagingEndpoint } : {}),
  } as unknown as ConstructorParameters<MSTeamsTeamsSdk["App"]>[0]);
}

/**
 * Build a token provider that uses the Teams SDK App's public tokenManager
 * for token acquisition. Replaces the previous implementation that cast the
 * App to `unknown` to access protected getBotToken/getAppGraphToken methods.
 */
export function createMSTeamsTokenProvider(app: MSTeamsApp): MSTeamsTokenProvider {
  return {
    async getAccessToken(scope: string): Promise<string> {
      if (scope.includes("graph.microsoft.com")) {
        const token = await app.tokenManager.getGraphToken();
        return token ? String(token) : "";
      }
      const token = await app.tokenManager.getBotToken();
      return token ? String(token) : "";
    },
  };
}

export async function loadMSTeamsSdkWithAuth(
  creds: MSTeamsCredentials,
  options?: CreateMSTeamsAppOptions,
) {
  const sdk = await loadMSTeamsSdk();
  const app = await createMSTeamsApp(creds, sdk, options);
  return { sdk, app };
}

/**
 * Minimal send context for proactive messaging. Consumers use this to
 * send/update/delete activities on a specific conversation.
 */
export type MSTeamsSendContext = {
  sendActivity: (textOrActivity: string | object) => Promise<unknown>;
  updateActivity: (activityUpdate: object) => Promise<{ id?: string } | void>;
  deleteActivity: (activityId: string) => Promise<void>;
};

/**
 * Create a send context for a specific conversation using the SDK's API Client.
 * Uses the per-conversation serviceUrl (from the stored conversation reference)
 * rather than the App's default serviceUrl, so proactive sends route to the
 * correct regional Bot Framework endpoint.
 */
export function createProactiveSendContext(params: {
  sdk: MSTeamsTeamsSdk;
  app: MSTeamsApp;
  serviceUrl: string;
  conversationId: string;
  conversationType?: string;
  bot?: { id?: string; name?: string };
  replyToActivityId?: string;
  tenantId?: string;
  recipientId?: string;
  recipientAadObjectId?: string;
}): MSTeamsSendContext {
  const apiClient = new params.sdk.Client(params.serviceUrl, {
    token: async () => {
      const token = await params.app.tokenManager.getBotToken();
      return token ? String(token) : undefined;
    },
    headers: { "User-Agent": buildUserAgent() },
  } as Record<string, unknown>);

  function normalizeActivity(textOrActivity: string | object): Record<string, unknown> {
    return typeof textOrActivity === "string"
      ? ({ type: "message", text: textOrActivity } as Record<string, unknown>)
      : (textOrActivity as Record<string, unknown>);
  }

  return {
    async sendActivity(textOrActivity: string | object): Promise<unknown> {
      const msg = normalizeActivity(textOrActivity);

      // Merge caller-provided channelData with the tenant metadata so Bot
      // Framework receives `channelData.tenant.id` for proactive routing.
      const existingChannelData =
        msg.channelData && typeof msg.channelData === "object"
          ? (msg.channelData as Record<string, unknown>)
          : undefined;
      const channelData = params.tenantId
        ? { ...existingChannelData, tenant: { id: params.tenantId } }
        : existingChannelData;

      return await apiClient.conversations.activities(params.conversationId).create({
        type: "message",
        ...msg,
        ...(channelData ? { channelData } : {}),
        from: params.bot?.id
          ? { id: params.bot.id, name: params.bot.name ?? "", role: "bot" }
          : undefined,
        conversation: {
          id: params.conversationId,
          conversationType: params.conversationType ?? "personal",
          ...(params.tenantId ? { tenantId: params.tenantId } : {}),
        },
        ...(params.recipientId || params.recipientAadObjectId
          ? {
              recipient: {
                ...(params.recipientId ? { id: params.recipientId } : {}),
                ...(params.recipientAadObjectId
                  ? { aadObjectId: params.recipientAadObjectId }
                  : {}),
              },
            }
          : {}),
        ...(params.replyToActivityId && !msg.replyToId
          ? { replyToId: params.replyToActivityId }
          : {}),
      } as Record<string, unknown>);
    },

    async updateActivity(activityUpdate: object): Promise<{ id?: string } | void> {
      const nextActivity = activityUpdate as { id?: string } & Record<string, unknown>;
      const activityId = nextActivity.id;
      if (!activityId) {
        throw new Error("updateActivity requires an activity id");
      }
      return await apiClient.conversations
        .activities(params.conversationId)
        .update(activityId, {
          type: "message",
          ...nextActivity,
        } as Record<string, unknown>);
    },

    async deleteActivity(activityId: string): Promise<void> {
      if (!activityId) {
        throw new Error("deleteActivity requires an activity id");
      }
      await apiClient.conversations.activities(params.conversationId).delete(activityId);
    },
  };
}
