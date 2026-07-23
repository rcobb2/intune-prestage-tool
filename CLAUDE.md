# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Microsoft Intune/Autopilot equivalent of the sibling `jamf-prestage-tool` project — same
runtime (Bun), same UI stack (Alpine.js + DaisyUI/Tailwind), same auth pattern (Entra ID),
same audit/two-person-approval workflow, but targets Microsoft Graph instead of the Jamf Pro
API. Read `README.md` before making changes — it documents exactly which Graph-backed
features are solid (Windows Autopilot) vs. lower-confidence or explicitly stubbed (Apple ADE),
and *why* Jamf's flat prestage-scope model doesn't map 1:1 onto Intune's group-membership-driven
enrollment profile assignment.

## Commands

```bash
bun install                 # install dependencies
docker compose up           # run both processes (client on :443, server on :8443)
bun test                    # run tests (jest via ts-jest; no test files exist yet)
```

There is no separate lint/typecheck script in `package.json` — use `tsc --noEmit` against
`tsconfig.json` if you need to typecheck. `docker-compose.yml` runs client and server as
separate containers with `network_mode: host` and bind-mounts (`develop.watch` syncs source
on change); `server/` runs `bun run server.ts` directly, `client/` runs `bun run --hot
worker.ts` which both bundles `client/main.ts` via `Bun.build` and serves the static output.

Set `SKIP_ENTRA_AUTH=true` in `.env` to bypass Microsoft login during local dev (mirrors the
Jamf tool's flag). TLS is opt-in: both `server/server.ts` and `client/worker.ts` only enable
`tls` if `server.cert`/`server.key` exist on disk — omit them when running behind a
TLS-terminating reverse proxy.

## Architecture

Two independent Bun processes, no shared runtime state:

- **`client/`** — an Alpine.js SPA. `client/worker.ts` both bundles `main.ts` (via
  `Bun.build`, with env vars baked in at build time through `define` — Bun 1.2 dropped the
  env-object API, so any new client-visible env var must be added explicitly to both the
  `define` block in `worker.ts` and read via `process.env.X` in `main.ts`/`azure-auth.ts`) and
  serves the built static output. `client/azure-auth.ts` owns the MSAL login flow and exposes
  `authReady`, a promise other Alpine components should `await` before firing authenticated
  requests on mount (avoids racing the popup/redirect flow).
- **`server/`** — a Bun-native REST API (`server/server.ts`, using Bun's built-in `routes`
  table, not a framework) that calls Microsoft Graph via `server/utils.ts`.

### Delegated auth — two tokens, one MSAL call

There is no separate server-to-Graph credential or second app registration, but there
ARE two distinct tokens per request, from a single `acquireTokenSilent({ scopes:
GRAPH_SCOPES })` call in `client/azure-auth.ts`:

- **ID token** (audience = this app's own `AZURE_CLIENT_ID`) — sent as the
  `Authorization` bearer, independently verified by `server/auth.ts`'s `withAuth` via
  the tenant's JWKS. This is what proves caller identity (`req.actor`).
- **Graph access token** (audience = Microsoft Graph) — sent via a separate
  `X-Graph-Token` header, forwarded to `req.graphToken` *without* independent
  verification. This is intentional, not an oversight: Microsoft documents Graph
  access tokens as opaque to everyone but Graph itself, and empirically (three
  independent verification methods, all failing identically against a real token
  from this tenant) they do not validate against the tenant's own published JWKS the
  way an ID token's signature does. Trusting it here is safe only because it arrives
  attached to a request whose ID token was already verified as the same session.

Every `server/utils.ts` function that calls Graph takes `req.graphToken` as an explicit
parameter; there is no cached, server-held Graph credential anywhere. Practically, this
means Graph enforces whatever Entra/Intune RBAC role the signed-in admin actually
holds — granting the delegated scopes on the app registration is necessary but not
sufficient, since each admin still needs a role like Intune Administrator. The
two-person approval flow relies on this: an approved wipe/retire executes using the
*approver's* Graph token, not the requester's, so Graph's own authorization is the real
second check, not just a button click in this UI. In `SKIP_ENTRA_AUTH` dev mode there
is no real token at all (`req.graphToken` is `''`), so only the local-only routes
(device metadata, audit log, approval bookkeeping) are meaningfully testable —
anything that calls Graph will fail.

### Platform split: Windows Autopilot vs. Apple ADE

Nearly every function in `server/utils.ts` branches on `Platform` ('windows' | 'apple'), and
the two paths have very different confidence levels — check the comment directly above a
function before extending it:

- **Windows** (`windowsAutopilotDeviceIdentities` / `windowsAutopilotDeploymentProfiles`):
  fully implemented — search, list/assign/remove enrollment profile, wipe, retire.
  `assignDeviceToProfile`/`removeDeviceFromProfile` have two modes switched by
  `AUTOPILOT_ASSIGNMENT_MODE`: `'direct'` (default) binds a deployment profile via
  `@odata.bind` then calls the identity's `assign` action, and unbinds via
  `deploymentProfile/$ref` DELETE; `'groupTag'` instead calls `updateDeviceProperties` to
  set/clear the Autopilot Group Tag (= the target profile's own `displayName`), for
  tenants (e.g. this one, Colgate) that assign profiles to Entra ID groups rather than
  individual devices — see the README's "Group-based Autopilot profile assignment"
  section for the one-time manual Entra/Intune group setup this mode depends on. This
  tool never creates/converts those groups itself.
- **Apple** (`depOnboardingSettings` / `importedDeviceIdentities`): read paths
  (list profiles, search) are implemented but lower-confidence (this corner of Graph beta is
  poorly documented). `assignDeviceToProfile`/`removeDeviceFromProfile` for `'apple'`
  deliberately `throw` — the correct per-device Graph mutation isn't confirmed, and this
  codebase does not guess at mutating Graph calls. `getEnrollmentProfileAssignment` returns a
  `'N/A'` sentinel for Apple for the same reason. Don't "fix" these without confirming the
  actual Graph endpoint against a real tenant first.

Most Graph calls hit `GRAPH_BASE` (v1.0); Autopilot deployment-profile and Apple ADE
operations hit `GRAPH_BETA` because those Graph surfaces are still beta-only.

### Local SQLite fills Graph's gaps (`server/db.ts`)

Three tables, all queried/written directly with `bun:sqlite` (no ORM):

- `audit_log` — every mutating action, written via `writeAudit()` in `server.ts` route handlers.
- `pending_approvals` — the two-person approval workflow for `wipe`/`retire`: one admin
  creates a request (`POST /api/approvals`), a *different* admin must approve it (`server.ts`
  rejects self-approval by comparing actor emails case-insensitively) before the underlying
  Graph mutation actually runs.
- `device_metadata` — the Jamf "Inventory Preload" equivalent, down to a single `username`
  field (covers both username and email — there's a startup migration in `db.ts` that
  drops pre-existing separate `email`/`building`/`room`/`asset_tag` columns if found, from
  when this table tracked more fields). Graph has no pre-enrollment metadata record, so
  this is tracked locally and joined onto live Graph device records at search time (see
  `buildRecordFromMetadata` in `utils.ts`).

### Route → utils.ts mapping

Routes in `server/server.ts` are thin: validate params/platform, call the corresponding
`server/utils.ts` export, write an audit entry, translate errors to HTTP status. When adding
a route, follow that pattern rather than putting Graph logic directly in `server.ts`. See the
README's route-mapping table for how each route corresponds to a Jamf-tool equivalent.

`CORS_HEADERS` (defined once in `utils.ts`, scoped to `https://${CLIENT_HOSTNAME}`) is spread
into every `Response` in `server.ts` — new routes must do the same rather than constructing
headers ad hoc.
