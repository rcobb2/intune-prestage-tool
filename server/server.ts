import logger from "./logger.ts";
import { writeAudit, getAuditLog, createApproval, getPendingApprovals, resolveApproval, getDeviceMetadata, upsertDeviceMetadata } from "./db.ts";
import * as utils from "./utils.ts";
import { CORS_HEADERS, type Platform } from "./utils.ts";
import { withAuth } from "./auth.ts";

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
      GET: withAuth(async (req) => {
        const { platform } = req.params;
        if (!isValidPlatform(platform)) {
          return new Response('Invalid platform — expected "windows" or "apple"', { ...CORS_HEADERS, status: 400 });
        }
        try {
          const profiles = await utils.getEnrollmentProfiles(platform);
          return new Response(JSON.stringify(profiles), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          logger.error({ err: error.message }, 'Error fetching enrollment profiles');
          return new Response('Error fetching enrollment profiles', { ...CORS_HEADERS, status: 500 });
        }
      })
    },

    "/api/devices/:search": {
      GET: withAuth(async (req) => {
        const { search } = req.params;
        logger.info(`[Device Search] Incoming search for: ${search}`);
        try {
          const results = await utils.searchDevices(search);
          if (results.length === 0) {
            return new Response('No device found', { ...CORS_HEADERS, status: 404 });
          }
          return new Response(JSON.stringify(results), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          return new Response(`${error.message || 'Unknown error'}`, { ...CORS_HEADERS, status: 500 });
        }
      })
    },

    "/api/change-enrollment-profile/:platform/:profileId/:serialNumber": {
      POST: withAuth(async (req) => {
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
          const current = await utils.getEnrollmentProfileAssignment(serialNumber, platform);
          if (current.displayName !== 'Unassigned' && current.displayName !== 'N/A') {
            const profiles = await utils.getEnrollmentProfiles(platform);
            const currentProfile = profiles.find(p => p.displayName === current.displayName);
            if (currentProfile && currentProfile.id !== profileId) {
              await utils.removeDeviceFromProfile(currentProfile.id, serialNumber, platform).catch((err: any) => {
                logger.warn({ err: err.message }, 'Warning: failed to remove from current enrollment profile');
              });
            }
          }

          const url = new URL(req.url, `http://${req.headers.get('host') || 'localhost'}`);
          const dryRun = url.searchParams.get('dryRun') === 'true';

          const result = await utils.assignDeviceToProfile(profileId, serialNumber, platform, dryRun);
          writeAudit({ action: 'enrollment_profile_change', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, details: { profileId, platform, dryRun }, result: 'success' });
          return new Response(JSON.stringify(result), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          writeAudit({ action: 'enrollment_profile_change', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, details: { profileId, platform }, result: 'error', error_detail: String(error.message) });
          logger.error({ err: error.message }, 'Assign enrollment profile error');
          return new Response(`Error assigning device to enrollment profile: ${error.message}`, { ...CORS_HEADERS, status: 500 });
        }
      }),
      DELETE: withAuth(async (req) => {
        const { profileId, serialNumber, platform } = req.params;
        if (!isValidPlatform(platform)) {
          return new Response('Invalid platform — expected "windows" or "apple"', { ...CORS_HEADERS, status: 400 });
        }
        logger.info(`Removing ${platform} device with serial number: ${serialNumber} from enrollment profile: ${profileId}`);

        try {
          const result = await utils.removeDeviceFromProfile(profileId, serialNumber, platform);
          return new Response(JSON.stringify(result), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          logger.error({ err: error.message }, 'Remove enrollment profile error');
          return new Response(`Error removing device from enrollment profile: ${error.message}`, { ...CORS_HEADERS, status: 500 });
        }
      })
    },

    // Jamf's "Inventory Preload" equivalent — local-only metadata (username, email,
    // building, room, asset tag) since Graph has no native pre-enrollment record for
    // these fields. See db.ts device_metadata table.
    "/api/device-metadata/:serialNumber": {
      GET: withAuth(async (req) => {
        const serialNumber = decodeURIComponent(req.params.serialNumber);
        const metadata = getDeviceMetadata(serialNumber);
        return new Response(JSON.stringify(metadata ?? { serialNumber, username: null, email: null, building: null, room: null, assetTag: null }), { ...CORS_HEADERS, status: 200 });
      }),
      PUT: withAuth(async (req) => {
        const serialNumber = decodeURIComponent(req.params.serialNumber);
        const body = await req.json() as { username?: string; email?: string; building?: string; room?: string; assetTag?: string };

        logger.info({ serialNumber }, 'Updating device metadata');
        try {
          upsertDeviceMetadata({
            serialNumber,
            username: body.username ?? null,
            email: body.email ?? null,
            building: body.building ?? null,
            room: body.room ?? null,
            assetTag: body.assetTag ?? null,
          });
          writeAudit({ action: 'update_metadata', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, details: body, result: 'success' });
          return new Response(JSON.stringify({ ok: true }), { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          writeAudit({ action: 'update_metadata', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, result: 'error', error_detail: String(error.message) });
          return new Response(`Error updating device metadata: ${error.message}`, { ...CORS_HEADERS, status: 500 });
        }
      })
    },

    "/api/wipedevice/:deviceId": {
      DELETE: withAuth(async (req) => {
        const { deviceId } = req.params;
        logger.info({ deviceId }, 'Wiping device');
        const result = await utils.wipeDevice(deviceId);
        writeAudit({ action: 'wipe', actor: getActor(req), ip: getIP(req), device_id: deviceId, result: result.status === 200 ? 'success' : 'error' });
        return result;
      })
    },

    "/api/retiredevice/:deviceId/:serialNumber/:macAddress/:altMacAddress": {
      DELETE: withAuth(async (req) => {
        const { deviceId, serialNumber, macAddress, altMacAddress } = req.params;
        logger.info({ deviceId, serialNumber }, 'Retiring device');

        try {
          const result = await utils.retireDevice(deviceId, serialNumber, macAddress, altMacAddress);
          if (!result.ok) {
            return new Response(result.message ?? 'Retirement failed', { ...CORS_HEADERS, status: 500 });
          }
          writeAudit({ action: 'retire', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, device_id: deviceId, result: 'success' });
          return new Response('Device retired successfully', { ...CORS_HEADERS, status: 200 });
        } catch (error: any) {
          writeAudit({ action: 'retire', actor: getActor(req), ip: getIP(req), device_serial: serialNumber, device_id: deviceId, result: 'error', error_detail: String(error.message) });
          return new Response(error.message ?? 'Error retiring device', { ...CORS_HEADERS, status: 500 });
        }
      })
    },

    "/api/audit-log": {
      GET: withAuth(async (req) => {
        const url = new URL(req.url, `http://${req.headers.get('host') || 'localhost'}`);
        const limit = Math.min(parseInt(url.searchParams.get('limit') ?? '100', 10) || 100, 500);
        const entries = getAuditLog(limit);
        return new Response(JSON.stringify(entries), { ...CORS_HEADERS, status: 200 });
      })
    },

    "/api/approvals": {
      POST: withAuth(async (req) => {
        try {
          const body = await req.json() as { action: string; justification?: string; deviceSerial: string; deviceId?: string; deviceAssetTag?: string; payload: object };
          const { action, justification, deviceSerial, deviceId, deviceAssetTag, payload } = body;
          const requester = getActor(req);

          const id = createApproval({ action, requester, justification, device_serial: deviceSerial, device_id: deviceId, device_asset_tag: deviceAssetTag, payload });
          logger.info({ action, requester, deviceSerial, justification }, 'Approval request created');
          return new Response(JSON.stringify({ id }), { ...CORS_HEADERS, status: 201 });
        } catch (error: any) {
          return new Response(JSON.stringify({ error: error.message }), { ...CORS_HEADERS, status: 500 });
        }
      })
    },

    "/api/approvals/pending": {
      GET: withAuth(async () => {
        const pending = getPendingApprovals();
        return new Response(JSON.stringify({ count: pending.length, items: pending }), { ...CORS_HEADERS, status: 200 });
      })
    },

    "/api/approvals/:id/approve": {
      POST: withAuth(async (req) => {
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
            actionResult = await utils.wipeDevice(payload.deviceId);
          } else if (approval.action === 'retire') {
            const { deviceId, serialNumber, macAddress, altMacAddress } = payload;
            if (!deviceId || !serialNumber) {
              return new Response(JSON.stringify({ error: 'Approval payload missing deviceId or serialNumber' }), { ...CORS_HEADERS, status: 500 });
            }
            const result = await utils.retireDevice(deviceId, serialNumber, macAddress, altMacAddress);
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
      })
    },

    "/api/approvals/:id/reject": {
      POST: withAuth(async (req) => {
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
      })
    },

    // Intentionally public/unauthenticated — the client needs this before it knows
    // whether auth is even required, and it exposes no data beyond a boolean flag.
    "/api/config": {
      async GET() {
        const skip = process.env.SKIP_ENTRA_AUTH === 'true';
        return new Response(JSON.stringify({ skipEntraAuth: skip }), { ...CORS_HEADERS, status: 200 });
      }
    },
    "/api/*": {
      async OPTIONS() {
        return new Response('CORS preflight', CORS_HEADERS);
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
