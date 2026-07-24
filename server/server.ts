import logger from "./logger.ts";
import { writeAudit, getAuditLog, createApproval, getPendingApprovals, resolveApproval, getDeviceMetadata, upsertDeviceMetadata } from "./db.ts";
import * as utils from "./utils.ts";
import { CORS_HEADERS, type Platform } from "./utils.ts";
import { withAuth } from "./auth.ts";
import { register, withMetrics } from "./metrics.ts";

// Verified by withAuth from the caller's Entra token; X-User-Name is a legacy
// fallback only reachable if a route is ever added without the withAuth wrapper.
function getActor(req: Request): string {
  return (req as any).actor ?? req.headers.get('X-User-Name') ?? 'unknown';
}

function getIP(req: Request): string {
  return req.headers.get('X-Forwarded-For') ?? 'unknown';
}

const notFound = Bun.file("404.html");

const {
  SERVER_API_HOSTNAME,
  SERVER_API_PORT,

  // Decoupled from the two above: SERVER_API_HOSTNAME/PORT are the public address baked into
  // the browser bundle, while these control what the container actually binds to behind a
  // reverse proxy. Both fall back to the public values so single-host/no-proxy setups are unaffected.
  SERVER_BIND_HOST,
  SERVER_BIND_PORT,

  // Optional shared secret for scraping /metrics. Unset means the endpoint is
  // open — fine when it's only reachable over loopback (matching how this
  // service is published in production), but set it if that ever changes.
  METRICS_TOKEN,
} = process.env;

// TLS is only enabled when cert/key files are present (self-signed local/docker-compose dev).
// Behind a reverse proxy (e.g. Caddy in prod) no certs are provided and this serves plain HTTP.
const hasTls = await Bun.file("server.cert").exists() && await Bun.file("server.key").exists();

function isValidPlatform(value: string): value is Platform {
  return value === 'windows' || value === 'apple';
}

// @ts-ignore
const server: Bun.Server = Bun.serve({
  development: false,
  hostname: SERVER_BIND_HOST || SERVER_API_HOSTNAME || "localhost",
  port: SERVER_BIND_PORT || SERVER_API_PORT || 3001,
  ...(hasTls ? {
    tls: {
      key: Bun.file("server.key"),
      cert: Bun.file("server.cert"),
    },
  } : {}),
  routes: {
    "/api/enrollment-profiles/:platform": {
      GET: withMetrics('/api/enrollment-profiles/:platform', withAuth(async (req) => {
        const { platform } = req.params;
        if (!isValidPlatform(platform)) {
          return new Response('Invalid platform — expected "windows" or "apple"', { ...CORS_HEADERS, status: 400 });
        }
        try {
          const profiles = await utils.getEnrollmentProfiles(platform, req.graphToken);
          return new Response(JSON.stringify(profiles), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          logger.error({ err: error.message }, 'Error fetching enrollment profiles');
          return new Response('Error fetching enrollment profiles', { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    "/api/devices/:search": {
      GET: withMetrics('/api/devices/:search', withAuth(async (req) => {
        const { search } = req.params;
        logger.info(`[Device Search] Incoming search for: ${search}`);
        try {
          const results = await utils.searchDevices(search, req.graphToken);
          if (results.length === 0) {
            return new Response('No device found', { ...CORS_HEADERS, status: 404 });
          }
          return new Response(JSON.stringify(results), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          return new Response(`${error.message || 'Unknown error'}`, { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    "/api/change-enrollment-profile/:platform/:profileId/:serialNumber": {
      POST: withMetrics('/api/change-enrollment-profile/:platform/:profileId/:serialNumber', withAuth(async (req) => {
        const { serialNumber, profileId, platform } = req.params;
        if (!isValidPlatform(platform)) {
          return new Response('Invalid platform — expected "windows" or "apple"', { ...CORS_HEADERS, status: 400 });
        }
        logger.info(`Assigning ${platform} device with serial number: ${serialNumber} to enrollment profile: ${profileId}`);

        const serialRegex = /^[A-Z0-9-]{4,}$/i;
        if (!serialRegex.test(serialNumber)) {
          return new Response('Invalid serial number format', { ...CORS_HEADERS, status: 400 });
        }

        try {
          // First, find the current profile assignment and remove it if it differs
          // from the target — a device can only be validly claimed by one enrollment
          // profile at a time (same constraint as Jamf prestage scope membership).
          const current = await utils.getEnrollmentProfileAssignment(serialNumber, platform, req.graphToken);
          if (current.displayName !== 'Unassigned' && current.displayName !== 'N/A') {
            const profiles = await utils.getEnrollmentProfiles(platform, req.graphToken);
            const currentProfile = profiles.find(p => p.displayName === current.displayName);
            if (currentProfile && currentProfile.id !== profileId) {
              await utils.removeDeviceFromProfile(currentProfile.id, serialNumber, platform, req.graphToken).catch((err: any) => {
                logger.warn({ err: err.message }, 'Warning: failed to remove from current enrollment profile');
              });
            }
          }

          const url = new URL(req.url, `http://${req.headers.get('host') || 'localhost'}`);
          const dryRun = url.searchParams.get('dryRun') === 'true';

          const result = await utils.assignDeviceToProfile(profileId, serialNumber, platform, req.graphToken, dryRun);
          writeAudit({ action: 'enrollment_profile_change', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, details: { profileId, platform, dryRun }, result: 'success' });
          return new Response(JSON.stringify(result), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          writeAudit({ action: 'enrollment_profile_change', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, details: { profileId, platform }, result: 'error', error_detail: String(error.message) });
          logger.error({ err: error.message }, 'Assign enrollment profile error');
          return new Response(`Error assigning device to enrollment profile: ${error.message}`, { ...CORS_HEADERS, status: 500 });
        }
      })),
      DELETE: withMetrics('/api/change-enrollment-profile/:platform/:profileId/:serialNumber', withAuth(async (req) => {
        const { profileId, serialNumber, platform } = req.params;
        if (!isValidPlatform(platform)) {
          return new Response('Invalid platform — expected "windows" or "apple"', { ...CORS_HEADERS, status: 400 });
        }
        logger.info(`Removing ${platform} device with serial number: ${serialNumber} from enrollment profile: ${profileId}`);

        try {
          const result = await utils.removeDeviceFromProfile(profileId, serialNumber, platform, req.graphToken);
          return new Response(JSON.stringify(result), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          logger.error({ err: error.message }, 'Remove enrollment profile error');
          return new Response(`Error removing device from enrollment profile: ${error.message}`, { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    // Jamf's "Inventory Preload" equivalent — local-only metadata (username, covering
    // both username and email as one free-text field) since Graph has no native
    // pre-enrollment record for this. See db.ts device_metadata table.
    "/api/device-metadata/:serialNumber": {
      GET: withMetrics('/api/device-metadata/:serialNumber', withAuth(async (req) => {
        const serialNumber = decodeURIComponent(req.params.serialNumber);
        const metadata = getDeviceMetadata(serialNumber);
        return new Response(JSON.stringify(metadata ?? { serialNumber, username: null }), { ...CORS_HEADERS, status: 200 });
      })),
      PUT: withMetrics('/api/device-metadata/:serialNumber', withAuth(async (req) => {
        const serialNumber = decodeURIComponent(req.params.serialNumber);
        const body = await req.json() as { username?: string };

        logger.info({ serialNumber }, 'Updating device metadata');
        try {
          upsertDeviceMetadata({
            serialNumber,
            username: body.username ?? null,
          });
          writeAudit({ action: 'update_metadata', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, details: body, result: 'success' });
          return new Response(JSON.stringify({ ok: true }), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          writeAudit({ action: 'update_metadata', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, result: 'error', error_detail: String(error.message) });
          return new Response(`Error updating device metadata: ${error.message}`, { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    // Renames an already-enrolled Windows device via Graph's setDeviceName action —
    // see the comment above utils.ts: renameDevice for why this can't just be a PATCH.
    // Not instant: applies next time the device checks in.
    "/api/rename-device/:deviceId": {
      PUT: withMetrics('/api/rename-device/:deviceId', withAuth(async (req) => {
        const { deviceId } = req.params;
        const body = await req.json() as { name?: string };
        const newName = (body.name ?? '').trim();
        if (!newName) {
          return new Response('Missing new device name', { ...CORS_HEADERS, status: 400 });
        }

        logger.info({ deviceId, newName }, 'Renaming device');
        try {
          const result = await utils.renameDevice(deviceId, newName, req.graphToken);
          writeAudit({ action: 'rename_device', actor: getActor(req), ip: getIP(req), device_id: deviceId, details: { newName }, result: 'success' });
          return new Response(JSON.stringify(result), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          writeAudit({ action: 'rename_device', actor: getActor(req), ip: getIP(req), device_id: deviceId, details: { newName }, result: 'error', error_detail: String(error.message) });
          return new Response(`Error renaming device: ${error.message}`, { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    "/api/wipedevice/:deviceId": {
      DELETE: withMetrics('/api/wipedevice/:deviceId', withAuth(async (req) => {
        const { deviceId } = req.params;
        logger.info({ deviceId }, 'Wiping device');
        const result = await utils.wipeDevice(deviceId, req.graphToken);
        writeAudit({ action: 'wipe', actor: getActor(req), ip: getIP(req), device_id: deviceId, result: result.status === 200 ? 'success' : 'error' });
        return result;
      }))
    },

    "/api/retiredevice/:deviceId/:serialNumber/:macAddress/:altMacAddress": {
      DELETE: withMetrics('/api/retiredevice/:deviceId/:serialNumber/:macAddress/:altMacAddress', withAuth(async (req) => {
        const { deviceId, serialNumber, macAddress, altMacAddress } = req.params;
        logger.info({ deviceId, serialNumber }, 'Retiring device');

        try {
          const result = await utils.retireDevice(deviceId, serialNumber, req.graphToken, macAddress, altMacAddress);
          if (!result.ok) {
            return new Response(result.message ?? 'Retirement failed', { ...CORS_HEADERS, status: 500 });
          }
          writeAudit({ action: 'retire', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, device_id: deviceId, result: 'success' });
          return new Response('Device retired successfully', { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          writeAudit({ action: 'retire', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, device_id: deviceId, result: 'error', error_detail: String(error.message) });
          return new Response(error.message ?? 'Error retiring device', { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    "/api/audit-log": {
      GET: withMetrics('/api/audit-log', withAuth(async (req) => {
        const url = new URL(req.url, `http://${req.headers.get('host') || 'localhost'}`);
        const limit = Math.min(parseInt(url.searchParams.get('limit') ?? '100', 10) || 100, 500);
        const entries = getAuditLog(limit);
        return new Response(JSON.stringify(entries), { ...CORS_HEADERS, status: 200 });
      }))
    },

    "/api/approvals": {
      POST: withMetrics('/api/approvals', withAuth(async (req) => {
        try {
          const body = await req.json() as { action: string; justification?: string; deviceSerial: string; deviceId?: string; payload: object };
          const { action, justification, deviceSerial, deviceId, payload } = body;
          const requester = getActor(req);

          const id = createApproval({ action, requester, justification, device_serial: deviceSerial, device_id: deviceId, payload });
          logger.info({ action, requester, deviceSerial, justification }, 'Approval request created');
          return new Response(JSON.stringify({ id }), { ...CORS_HEADERS, status: 201 });
        } catch (error: any) {
          return new Response(JSON.stringify({ error: error.message }), { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    "/api/approvals/pending": {
      GET: withMetrics('/api/approvals/pending', withAuth(async () => {
        const pending = getPendingApprovals();
        return new Response(JSON.stringify({ count: pending.length, items: pending }), { ...CORS_HEADERS, status: 200 });
      }))
    },

    "/api/approvals/:id/approve": {
      POST: withMetrics('/api/approvals/:id/approve', withAuth(async (req) => {
        try {
          const id = parseInt(req.params.id, 10);
          const approver = getActor(req);

          const pending = getPendingApprovals();
          const approval = (pending as any[]).find((a: any) => a.id === id);
          if (!approval) {
            return new Response(JSON.stringify({ error: 'Approval not found or already resolved' }), { ...CORS_HEADERS, status: 404 });
          }

          if (approval.requester.toLowerCase() === approver.toLowerCase()) {
            return new Response(JSON.stringify({ error: 'A second admin must approve — you cannot approve your own request' }), { ...CORS_HEADERS, status: 400 });
          }

          resolveApproval(id, approver, 'approved');

          const payload = JSON.parse(approval.payload);
          if (!payload || typeof payload !== 'object') {
            return new Response(JSON.stringify({ error: 'Invalid approval payload' }), { ...CORS_HEADERS, status: 500 });
          }
          let actionResult: Response;
          if (approval.action === 'wipe') {
            if (!payload.deviceId) {
              return new Response(JSON.stringify({ error: 'Approval payload missing deviceId' }), { ...CORS_HEADERS, status: 500 });
            }
            // Executed with the APPROVER's own Graph token, not the original
            // requester's — the second admin's Entra/Intune RBAC is what actually
            // authorizes this action, not just their click in this UI.
            actionResult = await utils.wipeDevice(payload.deviceId, req.graphToken);
          } else if (approval.action === 'retire') {
            const { deviceId, serialNumber, macAddress, altMacAddress } = payload;
            if (!deviceId || !serialNumber) {
              return new Response(JSON.stringify({ error: 'Approval payload missing deviceId or serialNumber' }), { ...CORS_HEADERS, status: 500 });
            }
            const result = await utils.retireDevice(deviceId, serialNumber, req.graphToken, macAddress, altMacAddress);
            actionResult = new Response(
              result.ok ? JSON.stringify({ status: 'retired' }) : JSON.stringify({ error: result.message }),
              { ...CORS_HEADERS, status: result.ok ? 200 : 500 }
            );
          } else {
            actionResult = new Response('Unknown action', { ...CORS_HEADERS, status: 400 });
          }

          writeAudit({ action: approval.action, actor: `${approval.requester} (req) / ${approver} (appr)`, ip: getIP(req), device_serial: approval.device_serial, device_id: approval.device_id, result: actionResult.status < 300 ? 'success' : 'error' });
          logger.info({ action: approval.action, requester: approval.requester, approver, deviceSerial: approval.device_serial }, 'Approval executed');

          return new Response(JSON.stringify({ status: 'approved', actionStatus: actionResult.status }), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          return new Response(JSON.stringify({ error: error.message }), { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    "/api/approvals/:id/reject": {
      POST: withMetrics('/api/approvals/:id/reject', withAuth(async (req) => {
        try {
          const id = parseInt(req.params.id, 10);
          const approver = getActor(req);

          const pending = getPendingApprovals();
          const approval = (pending as any[]).find((a: any) => a.id === id);
          if (!approval) {
            return new Response(JSON.stringify({ error: 'Approval not found or already resolved' }), { ...CORS_HEADERS, status: 404 });
          }

          resolveApproval(id, approver, 'rejected');
          writeAudit({ action: `${approval.action}_rejected`, actor: approver, ip: getIP(req), device_serial: approval.device_serial, device_id: approval.device_id, result: 'success' });
          logger.info({ action: approval.action, requester: approval.requester, approver, deviceSerial: approval.device_serial }, 'Approval rejected');

          return new Response(JSON.stringify({ status: 'rejected' }), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          return new Response(JSON.stringify({ error: error.message }), { ...CORS_HEADERS, status: 500 });
        }
      }))
    },

    // Intentionally public/unauthenticated — the client needs this before it knows
    // whether auth is even required, and it exposes no data beyond a boolean flag.
    "/api/config": {
      GET: withMetrics('/api/config', async () => {
        const skip = process.env.SKIP_ENTRA_AUTH === 'true';
        return new Response(JSON.stringify({ skipEntraAuth: skip }), { ...CORS_HEADERS, status: 200 });
      })
    },
    "/api/*": {
      async OPTIONS() {
        return new Response('CORS preflight', CORS_HEADERS);
      }
    },

    // Scraped by Prometheus. Optionally gated by METRICS_TOKEN — see note above
    // where it's read from process.env.
    "/metrics": {
      async GET(req: Request) {
        if (METRICS_TOKEN && req.headers.get('X-Metrics-Token') !== METRICS_TOKEN) {
          return new Response('Unauthorized', { status: 401 });
        }
        return new Response(await register.metrics(), {
          status: 200,
          headers: { 'Content-Type': register.contentType },
        });
      }
    },

    "/*": () => new Response(notFound, { headers: { "Content-Type": "text/html" }, status: 404 }),
  },

  error() {
    return new Response("Error: Internal Server Error", { ...CORS_HEADERS, status: 500 });
  },
});

console.log(`Bun version: ${Bun.version_with_sha}`);
console.log(`Server listening on ${server.url}`);
