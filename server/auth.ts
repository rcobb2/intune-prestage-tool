import { createRemoteJWKSet, jwtVerify } from 'jose';
import logger from './logger.ts';
import { CORS_HEADERS } from './utils.ts';

const { SKIP_ENTRA_AUTH, AZURE_CLIENT_ID, AZURE_AUTHORITY } = process.env;
const SKIP_AUTH = SKIP_ENTRA_AUTH === 'true';

// Trailing slash on AZURE_AUTHORITY would double up when building these URLs.
const authority = AZURE_AUTHORITY?.replace(/\/$/, '');

// Microsoft Graph's own well-known application ID / App ID URI — stable across every
// tenant. The token this app verifies is a delegated Graph access token (see
// GRAPH_SCOPES in client/azure-auth.ts), so its audience is Graph, not AZURE_CLIENT_ID.
const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';
const GRAPH_APP_ID_URI = 'https://graph.microsoft.com';

// This verifies AND forwards the same token — there is no separate server-held Graph
// credential (no client-credentials app registration). The signed-in admin's own
// delegated Graph access token both authenticates them to this tool and is the exact
// token utils.ts uses to call Graph; Graph itself enforces whatever Entra/Intune RBAC
// role that admin actually holds.
//
// Microsoft Graph access tokens are commonly issued in v1.0 claim format (issuer
// `https://sts.windows.net/{tenantId}/`, audience the GUID above) even when requested
// through the v2.0 authorize/token endpoint — Graph's own app registration controls
// this, not ours. Both v1- and v2-format issuers/audiences are accepted below; if
// verification unexpectedly fails, log payload.iss/payload.aud from a real token from
// your tenant and adjust.
const JWKS = !SKIP_AUTH && authority
  ? createRemoteJWKSet(new URL(`${authority}/discovery/v2.0/keys`))
  : null;

// Only usable when AZURE_AUTHORITY is a tenant-specific URL (not the multi-tenant
// `/common`) — needed to build the v1-format issuer, which embeds the tenant GUID.
const tenantId = authority?.split('/').pop();
const isCommonAuthority = tenantId === 'common' || tenantId === 'organizations' || tenantId === 'consumers';

if (!SKIP_AUTH && (!AZURE_CLIENT_ID || !authority)) {
  logger.error('AZURE_CLIENT_ID/AZURE_AUTHORITY are not set and SKIP_ENTRA_AUTH is not true — every API request will be rejected');
}
if (!SKIP_AUTH && isCommonAuthority) {
  logger.warn('AZURE_AUTHORITY is a multi-tenant alias (/common, /organizations, or /consumers) — v1-format issuer validation cannot embed a tenant GUID in this mode; use your specific tenant ID instead if token verification fails');
}

// Verifies the Graph-scoped access token the client attaches to every request. Trusts
// the token's signed claims for the caller's identity — never the client-supplied
// X-User-Name header, which anyone can set to anything.
async function authenticate(req: Request): Promise<{ actor: string; graphToken: string } | Response> {
  if (SKIP_AUTH) {
    // No real token exists in this mode — any route that calls Microsoft Graph will
    // fail downstream. Only local-only routes (device metadata, audit log, approval
    // bookkeeping) are meaningfully usable without real Entra sign-in.
    return { actor: req.headers.get('X-User-Name') ?? 'dev-user', graphToken: '' };
  }

  const match = (req.headers.get('Authorization') ?? '').match(/^Bearer (.+)$/);
  if (!match) {
    return new Response('Missing bearer token', { ...CORS_HEADERS, status: 401 });
  }

  if (!JWKS || !AZURE_CLIENT_ID || !authority) {
    return new Response('Server authentication is not configured', { ...CORS_HEADERS, status: 500 });
  }

  try {
    const { payload } = await jwtVerify(match[1], JWKS, {
      issuer: isCommonAuthority
        ? [`${authority}/v2.0`]
        : [`${authority}/v2.0`, `https://sts.windows.net/${tenantId}/`],
      audience: [GRAPH_APP_ID, GRAPH_APP_ID_URI],
    });

    // Confirm the token was actually requested through THIS tool's own app
    // registration (appid = v1 claim name, azp = v2 claim name for the same thing) —
    // otherwise any client holding these delegated scopes for this user (e.g. Graph
    // Explorer) could call our API with a Graph token never issued to us.
    const requestingApp = (payload.appid ?? payload.azp) as string | undefined;
    if (requestingApp !== AZURE_CLIENT_ID) {
      logger.warn({ requestingApp }, 'Rejected Graph token not issued to this app registration');
      return new Response('Invalid or expired token', { ...CORS_HEADERS, status: 401 });
    }

    const actor = (payload.preferred_username ?? payload.upn ?? payload.unique_name ?? payload.email ?? payload.name ?? payload.sub) as string;
    return { actor, graphToken: match[1] };
  } catch (err: any) {
    logger.warn({ err: err.message }, 'Rejected request with invalid Entra token');
    return new Response('Invalid or expired token', { ...CORS_HEADERS, status: 401 });
  }
}

// Wraps a route handler so it only runs after successful authentication, and attaches
// the verified caller identity (req.actor) and their raw Graph access token
// (req.graphToken) for the handler/utils.ts Graph calls/audit log to use.
export function withAuth(handler: (req: any) => Response | Promise<Response>) {
  return async (req: any): Promise<Response> => {
    const result = await authenticate(req);
    if (result instanceof Response) return result;
    req.actor = result.actor;
    req.graphToken = result.graphToken;
    return handler(req);
  };
}
