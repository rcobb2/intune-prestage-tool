import axios from 'axios';
import logger from './logger.ts';
import axiosRetry from 'axios-retry';
import { getDeviceMetadata } from './db.ts';

// Without this, a hung connection to Graph (or GLPI/Clearpass) would hang the request
// handler indefinitely — axios has no default timeout.
axios.defaults.timeout = 15000;

// Configure global retry for all axios requests (3 retries, exponential backoff).
// Only retry safe/idempotent methods on 5xx — never retry POST/DELETE (wipe, retire, delete).
axiosRetry(axios, {
  retries: 3,
  retryDelay: axiosRetry.exponentialDelay,
  retryCondition: (error) => {
    const method = (error.config?.method ?? '').toUpperCase();
    const safeMethods = ['GET', 'HEAD', 'OPTIONS', 'PUT'];
    const isSafe = safeMethods.includes(method);
    return axiosRetry.isNetworkOrIdempotentRequestError(error) ||
      (isSafe && !!error.response && error.response.status >= 500);
  },
});

// Trace every outbound API call with method, URL, status, and duration
axios.interceptors.request.use((config) => {
  (config as any)._startTime = Date.now();
  logger.debug({ method: config.method?.toUpperCase(), url: config.url }, 'Outbound API request');
  return config;
});
axios.interceptors.response.use(
  (response) => {
    const ms = Date.now() - ((response.config as any)._startTime ?? 0);
    logger.info({ method: response.config.method?.toUpperCase(), url: response.config.url, status: response.status, ms }, 'API response');
    return response;
  },
  (error) => {
    const ms = Date.now() - ((error.config as any)?._startTime ?? 0);
    logger.error({ method: error.config?.method?.toUpperCase(), url: error.config?.url, status: error.response?.status, ms }, 'API error');
    return Promise.reject(error);
  }
);

const {
  CLEARPASS_INSTANCE,
  CLEARPASS_CLIENT_ID,
  CLEARPASS_CLIENT_SECRET,

  GLPI_INSTANCE,
  GLPI_APP_TOKEN,
  GLPI_USER_TOKEN,

  CLIENT_HOSTNAME,
  CLIENT_PORT,
} = process.env;

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
// Autopilot deployment profile assignment/read operations are still beta-only in Graph.
const GRAPH_BETA = 'https://graph.microsoft.com/beta';

// The browser's Origin header includes the port whenever it isn't the scheme default
// (443 for https) — CORS requires an exact match, so this must mirror the redirect-URI
// logic in client/azure-auth.ts exactly, or every request gets silently blocked by the
// browser once CLIENT_PORT is anything other than 443/unset.
const clientOrigin = (!CLIENT_PORT || CLIENT_PORT === '443')
  ? `https://${CLIENT_HOSTNAME}`
  : `https://${CLIENT_HOSTNAME}:${CLIENT_PORT}`;

export const CORS_HEADERS: ResponseInit = {
  headers: {
    "Access-Control-Allow-Origin": clientOrigin,
    "Access-Control-Allow-Methods": "GET, POST, PUT, PATCH, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Accept, Authorization, X-User-Name",
    "Access-Control-Allow-Credentials": "false",
    "Accept": "application/json",
    "Content-Type": "application/json",
    "Cache-Control": "no-cache",
  },
};

export type Platform = 'windows' | 'apple';

// The closest Graph/Intune analog to a Jamf PrestageEnrollment: a Windows Autopilot
// deployment profile, or an Apple Automated Device Enrollment (ADE) profile. Both are
// applied to a device before/at Setup Assistant/OOBE time, same as a Jamf prestage.
export type EnrollmentProfile = {
  id: string;
  displayName: string;
  platform: Platform;
};

// Vendor-neutral device record returned to the client — the Graph/Intune equivalent of
// the Jamf tool's JAMFResponse type. `building`/`room`/`username`/`email`/`assetTag` come
// from the local device_metadata table (see db.ts) since Graph has no native equivalent
// of Jamf's Inventory Preload; everything else comes live from Graph.
export type DeviceRecord = {
  serialNumber: string;
  intuneDeviceId: string | null;
  azureAdDeviceId: string | null;
  autopilotId: string | null;
  name: string | null;
  model: string | null;
  platform: Platform | 'unknown';
  currentEnrollmentProfile: string;
  groupTag: string | null;
  assignedUserPrincipalName: string | null;
  macAddress: string | null;
  altMacAddress: string | null;
  username: string | null;
  email: string | null;
  building: string | null;
  room: string | null;
  assetTag: string | null;
};

// ============================================================================
// Microsoft Graph auth — there is no server-held credential here. Every function
// below takes the caller's own delegated Graph access token (verified by
// server/auth.ts, forwarded via req.graphToken) and uses it directly. Graph enforces
// whatever Entra/Intune RBAC role the signed-in admin actually holds; this server
// never has broader access than the person currently using it.
// ============================================================================

// ============================================================================
// GLPI / Clearpass — vendor-agnostic, reused as-is from the Jamf tool.
// ============================================================================

export async function getGLPIToken() {
  return await axios.get(`${GLPI_INSTANCE}/initSession/`, {
    headers: {
      'Content-Type': 'application/json',
      'App-Token': GLPI_APP_TOKEN,
      'Authorization': `user_token ${GLPI_USER_TOKEN}`,
    },
  });
}

export async function cleanupGLPI(session_token: string) {
  return await axios.get(`${GLPI_INSTANCE}/killSession/`, {
    headers: {
      'Content-Type': 'application/json',
      'App-Token': GLPI_APP_TOKEN,
      'Session-Token': session_token,
    },
  });
}

export async function getClearpassToken() {
  const clearpassResp = await axios.post(`${CLEARPASS_INSTANCE}/oauth`, {
    grant_type: "client_credentials",
    client_id: CLEARPASS_CLIENT_ID,
    client_secret: CLEARPASS_CLIENT_SECRET,
  });

  if (clearpassResp.status !== 200) {
    throw new Error(`Failed to retrieve Clearpass access token: ${clearpassResp.status} ${clearpassResp.data}`);
  }

  logger.info('Successfully retrieved Clearpass access token');
  return clearpassResp.data.access_token;
}

export async function deleteClearpassMAC(macAddress: string): Promise<any> {
  const token = await getClearpassToken();

  const response = await axios.delete(`${CLEARPASS_INSTANCE}/endpoint/mac-address/${macAddress}`, {
    headers: { Authorization: `Bearer ${token}` }
  });

  logger.info({ macAddress }, 'Successfully deleted MAC address from Clearpass');
  return response.data;
}

// ============================================================================
// Graph request helpers
// ============================================================================

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

// Follows @odata.nextLink until exhausted. extraHeaders is used for endpoints that
// require ConsistencyLevel: eventual (advanced query capabilities, e.g. contains()).
async function graphGetAllPages<T = any>(url: string, token: string, extraHeaders?: Record<string, string>): Promise<T[]> {
  const all: T[] = [];
  let next: string | undefined = url;
  while (next) {
    const response: any = await axios.get(next, {
      headers: { Authorization: `Bearer ${token}`, ...extraHeaders },
    });
    all.push(...(response.data.value ?? []));
    next = response.data['@odata.nextLink'];
  }
  return all;
}

function platformFromOperatingSystem(operatingSystem?: string | null): Platform | 'unknown' {
  const os = (operatingSystem ?? '').toLowerCase();
  if (os.includes('windows')) return 'windows';
  if (os.includes('mac') || os.includes('ios') || os.includes('ipad')) return 'apple';
  return 'unknown';
}

// Looks up a device's Windows Autopilot identity by serial number, with its assigned
// deployment profile expanded so callers get the displayName without a second round trip.
async function findAutopilotIdentityBySerial(serialNumber: string, token: string): Promise<any | null> {
  try {
    const results = await graphGetAllPages<any>(
      `${GRAPH_BETA}/deviceManagement/windowsAutopilotDeviceIdentities?$filter=contains(serialNumber,'${escapeODataString(serialNumber)}')&$expand=deploymentProfile`,
      token
    );
    return results[0] ?? null;
  } catch (err: any) {
    logger.warn({ err: err.message, serialNumber }, 'Windows Autopilot device identity lookup failed');
    return null;
  }
}

function buildRecordFromMetadata(serialNumber: string) {
  const metadata = getDeviceMetadata(serialNumber);
  return {
    username: metadata?.username ?? null,
    email: metadata?.email ?? null,
    building: metadata?.building ?? null,
    room: metadata?.room ?? null,
    assetTag: metadata?.assetTag ?? null,
  };
}

async function buildEnrolledDeviceRecord(managedDevice: any, token: string): Promise<DeviceRecord> {
  const platform = platformFromOperatingSystem(managedDevice.operatingSystem);
  const metadata = buildRecordFromMetadata(managedDevice.serialNumber);
  let autopilotId: string | null = null;
  let groupTag: string | null = null;
  let assignedUserPrincipalName: string | null = managedDevice.userPrincipalName ?? null;
  let currentEnrollmentProfile = 'N/A';

  if (platform === 'windows') {
    const identity = await findAutopilotIdentityBySerial(managedDevice.serialNumber, token);
    if (identity) {
      autopilotId = identity.id;
      groupTag = identity.groupTag ?? null;
      assignedUserPrincipalName = identity.assignedUserPrincipalName ?? assignedUserPrincipalName;
      currentEnrollmentProfile = identity.deploymentProfile?.displayName ?? 'Unassigned';
    } else {
      currentEnrollmentProfile = 'Unassigned';
    }
  }

  return {
    serialNumber: managedDevice.serialNumber,
    intuneDeviceId: managedDevice.id,
    // The GUID Graph calls azureADDeviceId on managedDevice — this is the Entra ID
    // device's `deviceId` property, NOT its directory object id. retireDevice() below
    // resolves the directory object id separately when it needs to delete the object.
    azureAdDeviceId: managedDevice.azureADDeviceId ?? null,
    autopilotId,
    name: managedDevice.deviceName ?? null,
    model: managedDevice.model ?? null,
    platform,
    currentEnrollmentProfile,
    groupTag,
    assignedUserPrincipalName,
    macAddress: managedDevice.wiFiMacAddress ?? null,
    // managedDevice does not expose a second MAC address the way Jamf's mobile device
    // detail did (wifi + bluetooth) — left null.
    altMacAddress: null,
    username: metadata.username ?? managedDevice.userPrincipalName ?? null,
    email: metadata.email ?? managedDevice.emailAddress ?? null,
    building: metadata.building,
    room: metadata.room,
    assetTag: metadata.assetTag,
  };
}

function buildAutopilotOnlyDeviceRecord(identity: any): DeviceRecord {
  return {
    serialNumber: identity.serialNumber,
    intuneDeviceId: identity.managedDeviceId ?? null,
    azureAdDeviceId: identity.azureActiveDirectoryDeviceId ?? null,
    autopilotId: identity.id,
    name: identity.displayName ?? null,
    model: identity.model ?? null,
    platform: 'windows',
    currentEnrollmentProfile: identity.deploymentProfile?.displayName ?? 'Unassigned',
    groupTag: identity.groupTag ?? null,
    assignedUserPrincipalName: identity.assignedUserPrincipalName ?? null,
    macAddress: null,
    altMacAddress: null,
    ...buildRecordFromMetadata(identity.serialNumber),
  };
}

// importedAppleDeviceIdentity field names below are a best-effort reading of the Graph
// beta schema (this corner of Graph is far less documented/stable than the Windows
// Autopilot APIs above) — verify against current Microsoft Learn docs before relying on
// this in production. The resource primarily tracks corporate device identifiers
// (serial number or IMEI) imported via an Apple Business/School Manager token; unlike
// managedDevice it does not carry rich attributes like model or MAC address.
function buildAppleOnlyDeviceRecord(identity: any): DeviceRecord {
  return {
    serialNumber: identity.serialNumber ?? identity.importedDeviceIdentifier ?? '',
    intuneDeviceId: null,
    azureAdDeviceId: null,
    autopilotId: null,
    name: identity.description ?? null,
    model: null,
    platform: 'apple',
    currentEnrollmentProfile: 'Unassigned',
    groupTag: null,
    assignedUserPrincipalName: null,
    macAddress: null,
    altMacAddress: null,
    ...buildRecordFromMetadata(identity.serialNumber ?? identity.importedDeviceIdentifier ?? ''),
  };
}

// ============================================================================
// Device search
//
// Graph has no single "search by anything" endpoint like Jamf's Classic API wildcard
// match. This fans out to enrolled devices first; only if nothing is enrolled does it
// fall back to pre-enrollment identities (Windows Autopilot / Apple ADE), mirroring the
// Jamf tool's own "search computers, then fall back to device-enrollments" shape.
//
// The managedDevices contains() filter below needs the tenant to support Graph's
// advanced query capabilities on this endpoint. If it doesn't, the request is caught
// and logged rather than allowed to crash the whole search — the pre-enrollment fallback
// still runs. If this warning shows up in your logs, switch to exact-match `eq` filters
// (serialNumber eq '{q}', deviceName eq '{q}') instead.
// ============================================================================
export async function searchDevices(query: string, token: string): Promise<DeviceRecord[]> {
  const escaped = escapeODataString(query.trim());

  let managedDevices: any[] = [];
  try {
    const filter = `contains(serialNumber,'${escaped}') or contains(deviceName,'${escaped}') or contains(userPrincipalName,'${escaped}') or contains(emailAddress,'${escaped}')`;
    managedDevices = await graphGetAllPages<any>(
      `${GRAPH_BASE}/deviceManagement/managedDevices?$filter=${encodeURIComponent(filter)}&$count=true`,
      token,
      { ConsistencyLevel: 'eventual' }
    );
  } catch (err: any) {
    logger.warn({ err: err.message }, 'managedDevices contains() filter failed — falling back to pre-enrollment search only; your tenant may need exact-match (eq) filters instead');
  }

  if (managedDevices.length > 0) {
    return Promise.all(managedDevices.map((md) => buildEnrolledDeviceRecord(md, token)));
  }

  const [autopilotMatches, appleMatches] = await Promise.all([
    graphGetAllPages<any>(
      `${GRAPH_BETA}/deviceManagement/windowsAutopilotDeviceIdentities?$filter=contains(serialNumber,'${escaped}')&$expand=deploymentProfile`,
      token
    ).catch((err: any) => {
      logger.warn({ err: err.message }, 'windowsAutopilotDeviceIdentities search failed');
      return [];
    }),
    graphGetAllPages<any>(
      `${GRAPH_BETA}/deviceManagement/importedAppleDeviceIdentities?$filter=contains(serialNumber,'${escaped}')`,
      token
    ).catch((err: any) => {
      logger.warn({ err: err.message }, 'importedAppleDeviceIdentities search failed (see field-name caveat on buildAppleOnlyDeviceRecord)');
      return [];
    }),
  ]);

  return [
    ...autopilotMatches.map(buildAutopilotOnlyDeviceRecord),
    ...appleMatches.map(buildAppleOnlyDeviceRecord),
  ];
}

// ============================================================================
// Enrollment profiles (Jamf prestage list equivalent)
// ============================================================================
export async function getEnrollmentProfiles(platform: Platform, token: string): Promise<EnrollmentProfile[]> {
  if (platform === 'windows') {
    const profiles = await graphGetAllPages<any>(`${GRAPH_BETA}/deviceManagement/windowsAutopilotDeploymentProfiles`, token);
    return profiles.map((p) => ({ id: p.id, displayName: p.displayName, platform: 'windows' as const }));
  }

  // Apple ADE profiles are scoped per Apple Business/School Manager (ABM/ASM) token —
  // there is no single flat list the way Jamf has one prestage list. Enumerate every
  // onboarding setting (one per ABM/ASM token) and flatten their enrollment profiles.
  // depOnboardingSettings/enrollmentProfiles nesting is a best-effort reading of Graph
  // beta — verify against current docs if this returns unexpected shapes.
  const settings = await graphGetAllPages<any>(`${GRAPH_BETA}/deviceManagement/depOnboardingSettings`, token);
  const profileLists = await Promise.all(
    settings.map((s: any) =>
      graphGetAllPages<any>(`${GRAPH_BETA}/deviceManagement/depOnboardingSettings/${s.id}/enrollmentProfiles`, token)
        .catch((err: any) => {
          logger.warn({ err: err.message, settingId: s.id }, 'Failed to list enrollment profiles for a depOnboardingSetting');
          return [];
        })
    )
  );
  return profileLists.flat().map((p: any) => ({ id: p.id, displayName: p.displayName, platform: 'apple' as const }));
}

// ============================================================================
// Current profile assignment for a device (Jamf getPrestageAssignments equivalent)
// ============================================================================
export async function getEnrollmentProfileAssignment(serialNumber: string, platform: Platform, token: string): Promise<{ serialNumber: string; displayName: string }> {
  if (platform === 'windows') {
    const identity = await findAutopilotIdentityBySerial(serialNumber, token);
    if (!identity) return { serialNumber, displayName: 'Unassigned' };
    return { serialNumber, displayName: identity.deploymentProfile?.displayName ?? 'Unassigned' };
  }

  // TODO: Apple ADE — resolving a device's current enrollment profile assignment needs
  // cross-referencing depOnboardingSettings/{id}/enrollmentProfiles/{id}/devices (or an
  // equivalent assignment collection); left as 'N/A' until that shape is confirmed
  // against your tenant, matching the Jamf tool's own "N/A" sentinel for unknown state.
  return { serialNumber, displayName: 'N/A' };
}

// ============================================================================
// Assign / reassign a device to an enrollment profile (Jamf addDeviceToPrestage
// equivalent — but the assignment MODEL is fundamentally different from Jamf's flat
// serial-number scope list, so this is not a drop-in port).
//
// Windows: uses Graph's documented per-device direct-assignment action — bind the
// target deploymentProfile onto the Autopilot identity via @odata.bind, then invoke the
// `assign` action to apply it. This is the closest Intune analog to Jamf's "PUT this one
// serial into this one prestage" semantics, and does not require pre-existing dynamic/
// static Entra ID groups. If your tenant instead assigns Autopilot profiles to groups via
// Group Tag, use updateDeviceProperties to set groupTag on the identity instead — the
// group's dynamic membership rule (and the profile-to-group assignment) already exists
// in that setup, so this call alone is sufficient there too.
//
// Apple ADE: Graph does not expose a per-device "assign" call the way Jamf's PUT scope
// endpoint does, and the exact assignment action for depOnboardingSettings profiles is
// not confirmed here — left unimplemented rather than guessed, since this mutates real
// enrollment configuration.
// ============================================================================
export async function assignDeviceToProfile(profileId: string, serialNumber: string, platform: Platform, token: string, dryRun?: boolean): Promise<any> {
  if (platform === 'windows') {
    const identity = await findAutopilotIdentityBySerial(serialNumber, token);
    if (!identity) {
      throw new Error(`No Windows Autopilot device identity found for serial ${serialNumber}`);
    }

    const bindUrl = `${GRAPH_BETA}/deviceManagement/windowsAutopilotDeviceIdentities/${identity.id}`;
    const bindBody = {
      'deploymentProfile@odata.bind': `${GRAPH_BETA}/deviceManagement/windowsAutopilotDeploymentProfiles/${profileId}`,
    };
    const assignUrl = `${GRAPH_BETA}/deviceManagement/windowsAutopilotDeviceIdentities/${identity.id}/assign`;

    if (dryRun) {
      return {
        dryRun: true,
        steps: [
          { method: 'PATCH', url: bindUrl, body: bindBody },
          { method: 'POST', url: assignUrl, body: {} },
        ],
      };
    }

    await axios.patch(bindUrl, bindBody, { headers: { Authorization: `Bearer ${token}` } });
    const assignResp = await axios.post(assignUrl, {}, { headers: { Authorization: `Bearer ${token}` } });
    return assignResp.data ?? { status: 'assigned' };
  }

  throw new Error('assignDeviceToProfile for platform "apple" is not implemented — see comment block above; the Graph action for per-device Apple ADE profile assignment is not confirmed and mutating writes should not be guessed at.');
}

// Windows Autopilot direct-assignment has no separate "unassign" action the way Jamf's
// scope/delete-multiple does — a device carries at most one bound deploymentProfile, and
// assignDeviceToProfile() above already re-binds atomically when moving to a new target.
// This best-effort attempts to clear the existing binding via a $ref delete for the
// standalone "remove" case; if your tenant's Graph version rejects it, the error is
// thrown (callers in server.ts already treat this as a non-fatal, logged step during
// reassignment — see the POST /api/change-enrollment-profile route).
export async function removeDeviceFromProfile(_profileId: string, serialNumber: string, platform: Platform, token: string): Promise<any> {
  if (platform === 'windows') {
    const identity = await findAutopilotIdentityBySerial(serialNumber, token);
    if (!identity) {
      throw new Error(`No Windows Autopilot device identity found for serial ${serialNumber}`);
    }
    const response = await axios.delete(
      `${GRAPH_BETA}/deviceManagement/windowsAutopilotDeviceIdentities/${identity.id}/deploymentProfile/$ref`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    return response.data ?? { status: 'unassigned' };
  }

  throw new Error('removeDeviceFromProfile for platform "apple" is not implemented — see assignDeviceToProfile comment block.');
}

// ============================================================================
// Wipe a device via Intune MDM command.
//
// keepEnrollmentData defaults to true: for a Windows Autopilot device this lets it
// re-provision itself automatically on next boot (Autopilot Reset) instead of dropping
// out of management — the closest Intune analog to how the Jamf tool always re-applies
// a prestage after erase.
//
// macOS Activation Lock bypass (Jamf's DEVICE_ERASE_PIN equivalent): Graph does not
// accept a caller-supplied PIN. Fetch the device's `activationLockBypassCode` via a
// separate GET before wiping and surface it to the tech — this function does not do
// that automatically since it changes the UX flow (the code must be shown BEFORE the
// wipe is confirmed, not after).
// ============================================================================
export async function wipeDevice(intuneDeviceId: string, token: string, options?: { keepEnrollmentData?: boolean; keepUserData?: boolean }): Promise<Response> {
  try {
    const body = {
      keepEnrollmentData: options?.keepEnrollmentData ?? true,
      keepUserData: options?.keepUserData ?? false,
    };
    await axios.post(`${GRAPH_BASE}/deviceManagement/managedDevices/${intuneDeviceId}/wipe`, body, {
      headers: { Authorization: `Bearer ${token}` }
    });
    return new Response(JSON.stringify({ status: 'wiped' }), { ...CORS_HEADERS, status: 200 });
  } catch (error: any) {
    const status = error?.response?.status ?? 500;
    const message = error?.response?.data?.error?.message ?? error.message ?? 'Error wiping device';
    return new Response(JSON.stringify(message), { ...CORS_HEADERS, status });
  }
}

// ============================================================================
// Full device retirement sequence: retire from Intune management → best-effort remove
// Windows Autopilot identity → best-effort delete Entra ID device object → GLPI update
// → Clearpass MAC removal.
//
// Steps after the Intune retire call are all non-fatal by design: once Intune retire has
// committed, a failure in directory cleanup or GLPI/Clearpass must never be surfaced as
// an overall failure to the caller — same principle as the Jamf tool's own retireDevice.
// ============================================================================
export async function retireDevice(
  intuneDeviceId: string,
  serialNumber: string,
  token: string,
  macAddress?: string,
  altMacAddress?: string,
): Promise<{ ok: boolean; message?: string }> {
  // Capture the Entra ID device GUID before retiring — the managedDevice object may
  // disappear once retired, and this is needed to clean up the Entra ID device object
  // afterward.
  let azureADDeviceGuid: string | null = null;
  try {
    const detail = await axios.get(`${GRAPH_BASE}/deviceManagement/managedDevices/${intuneDeviceId}?$select=azureADDeviceId`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    azureADDeviceGuid = detail.data.azureADDeviceId ?? null;
  } catch (err: any) {
    logger.warn({ err: err.message }, 'Could not read managed device detail before retire — Entra ID device cleanup may be skipped');
  }

  try {
    await axios.post(`${GRAPH_BASE}/deviceManagement/managedDevices/${intuneDeviceId}/retire`, {}, {
      headers: { Authorization: `Bearer ${token}` }
    });
  } catch (error: any) {
    const message = error?.response?.data?.error?.message ?? error.message;
    return { ok: false, message: `Intune retire failed: ${message}` };
  }

  // Best-effort: deregister the Windows Autopilot identity so the device doesn't
  // silently re-enroll on next boot. Skip if the device was never Autopilot-registered
  // (e.g. it's an Apple device, or was manually enrolled).
  try {
    const identity = await findAutopilotIdentityBySerial(serialNumber, token);
    if (identity) {
      await axios.delete(`${GRAPH_BETA}/deviceManagement/windowsAutopilotDeviceIdentities/${identity.id}`, {
        headers: { Authorization: `Bearer ${token}` }
      });
    }
  } catch (err: any) {
    logger.warn({ err: err.message, serialNumber }, 'Failed to delete Windows Autopilot device identity during retire — continuing');
  }

  // Best-effort: delete the Entra ID device object.
  try {
    if (azureADDeviceGuid) {
      const lookup = await axios.get(`${GRAPH_BASE}/devices?$filter=deviceId eq '${escapeODataString(azureADDeviceGuid)}'`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      const directoryObjectId = lookup.data.value?.[0]?.id;
      if (directoryObjectId) {
        await axios.delete(`${GRAPH_BASE}/devices/${directoryObjectId}`, {
          headers: { Authorization: `Bearer ${token}` }
        });
      }
    }
  } catch (err: any) {
    logger.warn({ err: err.message }, 'Failed to delete Entra ID device object during retire — continuing');
  }

  // GLPI state update (non-fatal after Intune retire has already committed)
  if (GLPI_INSTANCE && GLPI_APP_TOKEN) {
    try {
      const glpiTokenResp = await getGLPIToken();
      const sessionToken = glpiTokenResp.data.session_token;
      if (sessionToken) {
        const params = new URLSearchParams({
          'criteria[0][field]': '5',
          'criteria[0][searchtype]': 'contains',
          'criteria[0][value]': `^${serialNumber}$`,
        });
        const searchResp = await axios.get(`${GLPI_INSTANCE}/search/Computer`, {
          headers: { 'Content-Type': 'application/json', 'App-Token': GLPI_APP_TOKEN, 'Session-Token': sessionToken },
          params,
        });
        if (searchResp.data.totalcount === 1) {
          const computerIdGLPI = searchResp.data.data[0][2];
          await axios.put(`${GLPI_INSTANCE}/Computer/${computerIdGLPI}`, { input: { states_id: 18 } }, {
            headers: { 'Content-Type': 'application/json', 'App-Token': GLPI_APP_TOKEN, 'Session-Token': sessionToken },
          });
        } else {
          logger.warn({ serialNumber }, 'GLPI: device not found or multiple matches — skipping state update');
        }
        await cleanupGLPI(sessionToken).catch((err: any) => logger.warn({ err: err.message }, 'GLPI session cleanup failed'));
      }
    } catch (err: any) {
      logger.warn({ err: err.message }, 'GLPI retirement step failed — continuing');
    }
  }

  // Clearpass MAC removal (non-fatal)
  const removeMac = async (mac: string) => {
    try { await deleteClearpassMAC(mac); }
    catch (err: any) { logger.warn({ mac, err: err.message }, 'Clearpass MAC deletion failed'); }
  };
  if (macAddress) await removeMac(macAddress);
  if (altMacAddress) await removeMac(altMacAddress);

  return { ok: true };
}
