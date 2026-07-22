/// <reference lib="dom" />
// SKIP_AUTH is inlined as a boolean literal by Bun at build time (env: 'inline' in worker.ts).
// When true, MSAL is never imported or executed — the dynamic import below is dead code.
const SKIP_AUTH = process.env.SKIP_ENTRA_AUTH === 'true';

import axios from 'axios';

// Set once the MSAL instance exists (see init()) so the request interceptor below
// can reach it without being tied to a specific Alpine component instance.
let msalInstance: any = null;

// Resolves once a real, usable session exists (SKIP_AUTH, an existing cached
// session found on load, or a fresh interactive sign-in) — i.e. once requests
// will actually carry a token instead of 401ing. Other components that fetch
// data on mount (not behind a user click) should await this first; otherwise
// they race the MSAL popup/redirect flow and fire before any token exists.
let resolveAuthReady: (() => void) | null = null;
export const authReady: Promise<void> = SKIP_AUTH
  ? Promise.resolve()
  : new Promise<void>((resolve) => { resolveAuthReady = resolve; });

// Attaches a fresh Entra ID token to every outbound API request, so the server can
// verify the caller's identity from a signed token instead of trusting a client
// header (which is what the old X-User-Name-only scheme amounted to).
if (!SKIP_AUTH) {
  axios.interceptors.request.use(async (config) => {
    if (!msalInstance) return config;
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return config;
    try {
      // 'User.Read' is a default, always-consented Graph scope — we don't use the
      // access token it grants, just the fresh idToken MSAL returns alongside it.
      const result = await msalInstance.acquireTokenSilent({ scopes: ['User.Read'], account: accounts[0] });
      if (result?.idToken) {
        config.headers.set('Authorization', `Bearer ${result.idToken}`);
      }
    } catch (err) {
      // Leave the request unauthenticated; the server will reject it with 401
      // rather than this silently sending a stale/missing token.
      console.error('Failed to acquire token for request:', err);
    }
    return config;
  });
}

export default () => {
  return {
    isAuthenticated: SKIP_AUTH,
    errorMessage: '',
    _msal: null as any,

    async init() {
      if (SKIP_AUTH) return;

      const { PublicClientApplication } = await import('@azure/msal-browser');
      this._msal = new PublicClientApplication({
        auth: {
          clientId: process.env.AZURE_CLIENT_ID as string,
          authority: process.env.AZURE_AUTHORITY,
          redirectUri: (() => {
            const port = process.env.CLIENT_PORT;
            const host = process.env.CLIENT_HOSTNAME;
            return (port === '443' || port === '') ? `https://${host}/` : `https://${host}:${port}/`;
          })(),
        },
        cache: {
          cacheLocation: 'localStorage',
          storeAuthStateInCookie: false,
        },
      });
      msalInstance = this._msal;
      try {
        await this._msal.initialize();
        const accounts = this._msal.getAllAccounts();
        if (accounts.length > 0) {
          this.isAuthenticated = accounts.some(
            (account: any) => account.idTokenClaims?.aud === process.env.AZURE_CLIENT_ID
          );
          if (!this.isAuthenticated) {
            this.errorMessage = 'No authenticated account with the expected tenant ID found.';
          } else {
            resolveAuthReady?.();
          }
        }
        await this._msal.handleRedirectPromise();
      } catch (error: any) {
        console.error('MSAL Initialization Error:', error);
        this.errorMessage = `Authentication error: ${error.message}`;
      }
    },

    async signIn() {
      if (SKIP_AUTH) return;
      this.errorMessage = '';
      try {
        // Request the same scope acquireTokenSilent uses above, so consent is
        // captured now rather than forcing an interactive prompt on the first API call.
        const result = await this._msal.loginPopup({ scopes: ['User.Read'] });
        this.isAuthenticated = true;
        resolveAuthReady?.();
        if (result?.account?.name) {
          axios.defaults.headers.common['X-User-Name'] = result.account.name;
        }
      } catch (error: any) {
        console.error('Login Error:', error);
        this.errorMessage = `Login failed: ${error.message}`;
      }
    },

    async signOut() {
      if (SKIP_AUTH) return;
      this.errorMessage = '';
      try {
        await this._msal.logoutPopup();
        this.isAuthenticated = false;
      } catch (error: any) {
        console.error('Logout Error:', error);
        this.errorMessage = `Logout failed: ${error.message}`;
      }
    },
  };
};
