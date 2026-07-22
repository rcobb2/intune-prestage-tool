import { createRemoteJWKSet, jwtVerify } from 'jose';
import logger from './logger.ts';
import { CORS_HEADERS } from './utils.ts';

const { SKIP_ENTRA_AUTH, AZURE_CLIENT_ID, AZURE_AUTHORITY } = process.env;
const SKIP_AUTH = SKIP_ENTRA_AUTH === 'true';

// Trailing slash on AZURE_AUTHORITY would double up when building these URLs.
const authority = AZURE_AUTHORITY?.replace(/\/$/, '');

// This verifies the *user's* sign-in token (the person using this tool), via the same
// MSAL/Entra app registration pattern as the Jamf tool. It is a separate concern from
// graph-client.ts, which acquires a token so the *server* can call Microsoft Graph.
const JWKS = !SKIP_AUTH && authority
  ? createRemoteJWKSet(new URL(`${authority}/discovery/v2.0/keys`))
  : null;

if (!SKIP_AUTH && (!AZURE_CLIENT_ID || !authority)) {
  logger.error('AZURE_CLIENT_ID/AZURE_AUTHORITY are not set and SKIP_ENTRA_AUTH is not true — every API request will be rejected');
}

// Verifies the Entra ID token the client attaches to every request. Trusts the
// token's signed claims for the caller's identity — never the client-supplied
// X-User-Name header, which anyone can set to anything.
async function authenticate(req: Request): Promise<{ actor: string } | Response> {
  if (SKIP_AUTH) {
    return { actor: req.headers.get('X-User-Name') ?? 'dev-user' };
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
      issuer: `${authority}/v2.0`,
      audience: AZURE_CLIENT_ID,
    });
    const actor = (payload.preferred_username ?? payload.upn ?? payload.email ?? payload.name ?? payload.sub) as string;
    return { actor };
  } catch (err: any) {
    logger.warn({ err: err.message }, 'Rejected request with invalid Entra token');
    return new Response('Invalid or expired token', { ...CORS_HEADERS, status: 401 });
  }
}

// Wraps a route handler so it only runs after successful authentication, and
// attaches the verified caller identity to req.actor for the handler/audit log to use.
export function withAuth(handler: (req: any) => Response | Promise<Response>) {
  return async (req: any): Promise<Response> => {
    const result = await authenticate(req);
    if (result instanceof Response) return result;
    req.actor = result.actor;
    return handler(req);
  };
}
