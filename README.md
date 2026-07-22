# Intune Prestage Tool

A Microsoft Intune / Autopilot equivalent of
[Jamf Prestage Tool](https://github.com/rcobb2/jamf-prestage-tool), originally scaffolded
as a sibling project there. It mirrors that tool's shape — same runtime (Bun), same UI
stack (Alpine.js + DaisyUI/Tailwind), same auth pattern (Entra ID), same audit/approval
workflow — but targets Microsoft Graph instead of the Jamf Pro API.

**This has real Microsoft Graph calls wired in, but validate against a test tenant
before trusting it in production.** The Windows Autopilot path (search, enrollment
profile list/assign/remove, wipe, retire) is implemented against Graph endpoints and
field names I'm confident are correct. The Apple ADE path is implemented less
confidently — Graph's `depOnboardingSettings`/`importedAppleDeviceIdentities` corner is
far less documented, and per-device profile assignment for Apple ADE is left as an
explicit "not implemented" throw rather than a guessed mutating call. See the comments
in `server/utils.ts` above each function for exactly what's solid vs. what to verify.

## Why this isn't a 1:1 port

Jamf's Prestage model (a flat list of serial numbers scoped to one enrollment profile,
replaced wholesale on every write) does not map cleanly onto Intune:

- **Enrollment profile assignment is group-membership-driven, not a per-device scope
  list.** A Windows Autopilot deployment profile is assigned to an Entra ID group, and a
  device joins that group either via a dynamic group rule matching its Autopilot
  **Group Tag**, or via static group membership. There is no Graph endpoint that lets you
  PUT a list of serial numbers onto a profile the way Jamf's prestage scope endpoint
  does. `assignDeviceToProfile()` in `server/utils.ts` documents both strategies — pick
  whichever matches how your tenant's dynamic/static groups are already set up.
- **Apple devices enrolled via Intune (Apple ADE)** use yet another profile-assignment
  model, scoped per Apple Business/School Manager token (`depOnboardingSettings`) rather
  than one flat list.
- **Jamf's "Inventory Preload"** (username, email, building, room, asset tag pre-filled
  before a device ever checks in) has no Graph equivalent — Intune has no generic
  arbitrary-metadata record for a not-yet-enrolled device. This tool tracks those same
  fields locally in a `device_metadata` SQLite table (`server/db.ts`) and joins them onto
  live Graph data at search time, the same way the Jamf tool joins Jamf's own Preload
  records onto inventory data.
- **Device wipe has no PIN-based Activation Lock bypass equivalent.** Jamf's erase
  endpoint takes a caller-supplied PIN (`DEVICE_ERASE_PIN`); Graph instead exposes a
  device's `activationLockBypassCode` to *read*, which your workflow needs to surface to
  the tech *before* wiping a macOS device enrolled via Intune.

## What's implemented vs. stubbed

| Layer | Status |
|---|---|
| Entra ID user auth (`server/auth.ts`, `client/azure-auth.ts`) | Fully implemented — one MSAL call yields both an ID token (verified, proves identity) and a Graph access token (forwarded unverified, since Graph tokens aren't independently verifiable by design); no separate service principal |
| Audit log & two-person approval workflow (`server/db.ts`, approval routes) | Fully implemented — approving a request executes the action using the *approver's* own delegated Graph token, not the original requester's |
| Device search (`searchDevices`) | Implemented — enrolled devices via `managedDevices`, falling back to Windows Autopilot / Apple ADE pre-enrollment identities. The `contains()` filter used for substring search needs your tenant to support Graph's advanced query capabilities on `managedDevices`; if not, it's caught and logged, and pre-enrollment fallback still runs |
| Enrollment profile list (`getEnrollmentProfiles`) | Implemented for both platforms — Windows via `windowsAutopilotDeploymentProfiles`, Apple via `depOnboardingSettings`/`enrollmentProfiles` (lower confidence on Apple, see comment) |
| Current profile assignment (`getEnrollmentProfileAssignment`) | Implemented for Windows; returns `'N/A'` for Apple (**stubbed** — TODO in `server/utils.ts`) |
| Assign / remove profile (`assignDeviceToProfile`, `removeDeviceFromProfile`) | Implemented for Windows, two modes via `AUTOPILOT_ASSIGNMENT_MODE` — `direct` (bind + `assign` action) or `groupTag` (sets Autopilot Group Tag, relies on a pre-existing dynamic Entra ID group to actually assign the profile; see Requirements below). **Stubbed** (throws) for Apple — the exact per-device Graph action isn't confirmed, and a mutating write shouldn't be guessed at |
| Wipe (`wipeDevice`) | Implemented — `managedDevices/{id}/wipe` with `keepEnrollmentData` defaulted `true` so Autopilot devices reprovision |
| Retire (`retireDevice`) | Implemented — Intune retire, then best-effort cleanup of the Windows Autopilot identity and the Entra ID device object |
| GLPI / ClearPass retirement cleanup steps | Fully implemented — reused as-is, these are vendor-agnostic |
| UI (`client/index.html`, `client/main.ts`) | Fully implemented against the live API shape |

Not implemented, by design rather than oversight: macOS Activation Lock bypass code
retrieval/display before wipe (Graph exposes `activationLockBypassCode` to read, but
surfacing it needs a UX change — show the code *before* the wipe is confirmed, not
after), and Apple ADE profile assignment (see above).

## Requirements
- A Microsoft Entra ID tenant with Intune licensing.
- One Entra app registration (`AZURE_CLIENT_ID`/`AZURE_AUTHORITY`) for user sign-in,
  granted the DELEGATED Graph API permissions listed in `.env.example` with admin
  consent. There is no second, server-held app registration or client secret — every
  Graph call runs as the signed-in admin, using the same delegated access token that
  authenticates them to this tool. Each admin using the tool still needs an
  Entra/Intune role (e.g. Intune Administrator) that actually grants those rights —
  the scope alone isn't enough for Graph to authorize the call. Note that
  `Device.ReadWrite.All` (used for Entra ID device object cleanup during retire) has no
  delegated version — `Directory.ReadWrite.All` is the delegated permission actually
  used for that step.
- Docker (or `bun` installed locally for development without containers).

## Installation
1. Clone this repository.
2. Copy `.env.example` to `.env` and fill in your values.
3. Copy & name your SSL certs as `certs/server.cert` & `certs/server.key`
   (or omit them and put this behind a reverse proxy that terminates TLS).
4. Test against a non-production tenant first, especially the Windows Autopilot
   assign/remove/retire flows — they mutate real enrollment state.

## Usage
```bash
docker compose up
```
Client on `https://<CLIENT_HOSTNAME>` (default port 443), API on `:8443`. If you're
running this alongside the Jamf tool on the same host, change the exposed ports in
`docker-compose.yml` and the corresponding `CLIENT_PORT`/`SERVER_API_PORT` env vars —
both tools default to the same 443/8443 pair.

### Dev mode without Entra auth
Set `SKIP_ENTRA_AUTH=true` in `.env` to bypass the Microsoft login screen during local
development — same flag and behavior as the Jamf tool.

### Group-based Autopilot profile assignment (`AUTOPILOT_ASSIGNMENT_MODE=groupTag`)
Many tenants assign Windows Autopilot deployment profiles to Entra ID groups rather than
individual devices. This tool does not create, convert, or otherwise manage those groups
itself — that's real production identity configuration, and an in-place static→dynamic
group conversion silently *replaces* membership rather than merging it, which can drop
existing devices out of policies/apps scoped to the same group. Instead:

1. For each Windows Autopilot deployment profile you want this mode to work with, create a
   **new** Dynamic Device security group in Entra ID with a membership rule matching Group
   Tag = that profile's exact display name, e.g. for a profile named `EMPLOYEE`:
   ```
   (device.devicePhysicalIds -any (_ -eq "[OrderID]:EMPLOYEE"))
   ```
2. Add that new group to the profile's existing **Assignments** (Intune admin center →
   Devices → Enrollment → Windows enrollment → Deployment Profiles → *profile* →
   Assignments) — **alongside** any existing static group, not replacing it.
3. Set `AUTOPILOT_ASSIGNMENT_MODE=groupTag` in `.env`. Devices assigned through this tool
   from then on get their Group Tag set and pick up the dynamic group (and therefore the
   profile) automatically — allow a few minutes for dynamic membership evaluation.
4. Existing devices in an old static group are unaffected and keep working as before;
   backfilling their Group Tag and retiring the static group is optional future cleanup,
   not required for this mode to work.

## Architecture
Two Bun processes, same as the Jamf tool:
- `client/` — Alpine.js SPA, bundled and served by `client/worker.ts`.
- `server/` — Bun-native REST API (`server/server.ts`), calling Microsoft Graph via
  `server/utils.ts` instead of the Jamf Pro API.

Route shape mirrors the Jamf tool with Intune-appropriate names:

| Jamf tool route | Intune tool route |
|---|---|
| `GET /api/prestages` / `/api/mobile-prestages` | `GET /api/enrollment-profiles/:platform` (`platform` = `windows` \| `apple`) |
| `GET /api/computers/:search` / `/api/mobiledevices/:search` | `GET /api/devices/:search` |
| `POST /api/change-prestage/:deviceType/:prestageId/:serialNumber` | `POST /api/change-enrollment-profile/:platform/:profileId/:serialNumber` |
| `PUT /api/update-info/:deviceType/:preloadId/:computerId` | `PUT /api/device-metadata/:serialNumber` |
| `DELETE /api/wipedevice/:computerId` | `DELETE /api/wipedevice/:deviceId` |
| `DELETE /api/retiredevice/:computerId/:serialNumber/:macAddress/:altMacAddress` | `DELETE /api/retiredevice/:deviceId/:serialNumber/:macAddress/:altMacAddress` |
| `/api/audit-log`, `/api/approvals*`, `/api/config` | unchanged |

## License
This project is licensed under the [GPL 3.0 license](LICENSE).
