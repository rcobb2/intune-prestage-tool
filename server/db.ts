import { Database } from 'bun:sqlite';
import logger from './logger.ts';

const db = new Database(process.env.AUDIT_DB_PATH || '/app/audit.db', { create: true });

db.run(`CREATE TABLE IF NOT EXISTS audit_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  timestamp TEXT NOT NULL DEFAULT (datetime('now')),
  action TEXT NOT NULL,
  actor TEXT,
  ip TEXT,
  device_serial TEXT,
  device_id TEXT,
  details TEXT,
  result TEXT NOT NULL,
  error_detail TEXT
)`);

db.run(`CREATE TABLE IF NOT EXISTS pending_approvals (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  action TEXT NOT NULL,
  requester TEXT NOT NULL,
  justification TEXT,
  device_serial TEXT NOT NULL,
  device_id TEXT,
  payload TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending',
  approver TEXT,
  approved_at TEXT
)`);

// Graph/Intune has no equivalent of Jamf's "Inventory Preload" — there is no generic
// arbitrary-metadata record attachable to a device before it enrolls. This table is
// that equivalent, kept locally and joined onto live Graph data at search time.
// `username` covers both username and email (merged into one free-text field).
db.run(`CREATE TABLE IF NOT EXISTS device_metadata (
  serial_number TEXT PRIMARY KEY,
  username TEXT,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
)`);

// Migration: earlier versions of these tables had a separate `email` column (merged
// into `username` above), and `building`/`room`/`asset_tag` columns (removed
// entirely — no longer tracked). Drop them if a pre-existing table still has them —
// safe/no-op once already migrated, since each only acts when its column is present.
const deviceMetadataColumns = db.query(`PRAGMA table_info(device_metadata)`).all() as any[];
for (const col of ['email', 'building', 'room', 'asset_tag']) {
  if (deviceMetadataColumns.some((c) => c.name === col)) {
    db.run(`ALTER TABLE device_metadata DROP COLUMN ${col}`);
  }
}
const pendingApprovalsColumns = db.query(`PRAGMA table_info(pending_approvals)`).all() as any[];
if (pendingApprovalsColumns.some((c) => c.name === 'device_asset_tag')) {
  db.run(`ALTER TABLE pending_approvals DROP COLUMN device_asset_tag`);
}

export function writeAudit(entry: {
  action: string;
  actor?: string;
  ip?: string;
  device_serial?: string;
  device_id?: string;
  details?: object;
  result: 'success' | 'error';
  error_detail?: string;
}) {
  db.run(
    `INSERT INTO audit_log (action, actor, ip, device_serial, device_id, details, result, error_detail)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
    [
      entry.action,
      entry.actor ?? null,
      entry.ip ?? null,
      entry.device_serial ?? null,
      entry.device_id ?? null,
      entry.details ? JSON.stringify(entry.details) : null,
      entry.result,
      entry.error_detail ?? null,
    ]
  );
  const logFn = entry.result === 'error' ? logger.error.bind(logger) : logger.info.bind(logger);
  logFn({ audit: true, action: entry.action, actor: entry.actor, device_serial: entry.device_serial, device_id: entry.device_id, details: entry.details, error_detail: entry.error_detail }, `AUDIT: ${entry.action} [${entry.result}]`);
}

export function getAuditLog(limit = 100) {
  return db.query(`SELECT * FROM audit_log ORDER BY id DESC LIMIT ?`).all(limit);
}

export function createApproval(row: {
  action: string;
  requester: string;
  justification?: string;
  device_serial: string;
  device_id?: string;
  payload: object;
}) {
  const result = db.run(
    `INSERT INTO pending_approvals (action, requester, justification, device_serial, device_id, payload)
     VALUES (?, ?, ?, ?, ?, ?)`,
    [
      row.action,
      row.requester,
      row.justification ?? null,
      row.device_serial,
      row.device_id ?? null,
      JSON.stringify(row.payload),
    ]
  );
  return result.lastInsertRowid;
}

export function getPendingApprovals() {
  return db.query(`SELECT * FROM pending_approvals WHERE status = 'pending' ORDER BY created_at ASC`).all();
}

export function resolveApproval(id: number, approver: string, status: 'approved' | 'rejected') {
  db.run(
    `UPDATE pending_approvals SET status = ?, approver = ?, approved_at = datetime('now') WHERE id = ?`,
    [status, approver, id]
  );
  return db.query(`SELECT * FROM pending_approvals WHERE id = ?`).get(id) as any;
}

export type DeviceMetadata = {
  serialNumber: string;
  username: string | null;
};

export function getDeviceMetadata(serialNumber: string): DeviceMetadata | null {
  const row = db.query(`SELECT * FROM device_metadata WHERE serial_number = ?`).get(serialNumber) as any;
  if (!row) return null;
  return {
    serialNumber: row.serial_number,
    username: row.username,
  };
}

export function upsertDeviceMetadata(entry: DeviceMetadata) {
  db.run(
    `INSERT INTO device_metadata (serial_number, username, updated_at)
     VALUES (?, ?, datetime('now'))
     ON CONFLICT(serial_number) DO UPDATE SET
       username = excluded.username,
       updated_at = excluded.updated_at`,
    [entry.serialNumber, entry.username]
  );
}

export default db;
