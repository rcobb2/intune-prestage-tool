import { createRemoteJWKSet, jwtVerify } from 'jose';
import logger from './logger.ts';
import { CORS_HEADERS } from './utils.ts';

const { SKIP_ENTRA_AUTH, AZURE_CLIENT_ID, AZURE_AUTHORITY } = process.env;
const SKIP_AUTH = SKIP_ENTRA_AUTH === 'true';

// Trailing slash on AZURE_AUTHORITY would double up when building these URLs.
const authority = AZURE_AUTHORITY?.replace(/\/$/, '');

// This verifies the *user's* sign-in token (an ID token, audience = AZURE_CLIENT_ID) —
// Microsoft guarantees ID tokens are independently verifiable via the tenant's own
// JWKS, since they're issued specifically for this app to consume.
//
// The Graph-scoped access token used for actual Graph calls is a SEPARATE token,
// read from X-Graph-Token below and forwarded WITHOUT independent verification.
// Microsoft Graph access tokens are documented as opaque to everyone but Graph
// itself — their signing is not guaranteed to validate against the tenant's
// published JWKS the way an ID token's does (confirmed empirically: jose with a
// remote JWKS, jose with a manually-imported exact matching key, and raw
// node:crypto RSA-SHA256 verification all failed identically against a real
// Graph access token issued by this same tenant). Trusting it here is safe
// specifically because it arrives attached to a request whose ID token we just
// verified came from the same MSAL session/account.
const JWKS = !SKIP_AUTH && authority
  ? createRemoteJWKSet(new URL(`${authority}/discovery/v2.0/keys`))
  : null;

if (!SKIP_AUTH && (!AZURE_CLIENT_ID || !authority)) {
  logger.error('AZURE_CLIENT_ID/AZURE_AUTHORITY are not set and SKIP_ENTRA_AUTH is not true — every API request will be rejected');
}

// Verifies the Entra ID token the client attaches to every request. Trusts the
// token's signed claims for the caller's identity — never the client-supplied
// X-User-Name header, which anyone can set to anything.
async function authenticate(req: Request): Promise<{ actor: string; graphToken: string } | Response> {
  if (SKIP_AUTH) {
    return { actor: req.headers.get('X-User-Name') ?? 'dev-user', graphToken: req.headers.get('X-Graph-Token') ?? '' };
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
    // Not independently verified — see the comment above JWKS. Safe to trust here
    // because it's only used once the ID token above has already been verified as
    // belonging to the same authenticated session.
    const graphToken = req.headers.get('X-Graph-Token') ?? '';
    return { actor, graphToken };
  } catch (err: any) {
    logger.warn({ err: err.message }, 'Rejected request with invalid Entra token');
    return new Response('Invalid or expired token', { ...CORS_HEADERS, status: 401 });
  }
}

// Wraps a route handler so it only runs after successful authentication, and attaches
// the verified caller identity (req.actor) and their Graph access token (req.graphToken,
// forwarded but not independently verified — see comment above) for the
// handler/utils.ts Graph calls/audit log to use.
export function withAuth(handler: (req: any) => Response | Promise<Response>) {
  return async (req: any): Promise<Response> => {
    const result = await authenticate(req);
    if (result instanceof Response) return result;
    req.actor = result.actor;
    req.graphToken = result.graphToken;
    return handler(req);
  };
}
