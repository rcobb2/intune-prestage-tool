/// <reference lib="dom" />
import Alpine from 'alpinejs';
import AlpinePersist from "@alpinejs/persist";
import axios, { type AxiosResponse } from 'axios';
import AzureAuth, { authReady } from "./azure-auth.ts";

// Set up axios defaults
// SERVER_API_HOSTNAME/PORT let a reverse proxy (e.g. Caddy) put the client and API on the
// same public origin; unset, this falls back to the direct-to-container dev default.
const apiHost = process.env.SERVER_API_HOSTNAME || window.location.hostname;
const apiPort = process.env.SERVER_API_PORT || '8443';
const apiURL = `https://${apiHost}:${apiPort}/api`;
axios.defaults.baseURL = apiURL;

// Pass a dev placeholder for SKIP_ENTRA_AUTH mode; MSAL mode sets this after login via AzureAuth
if (process.env.SKIP_ENTRA_AUTH === 'true') {
  axios.defaults.headers.common['X-User-Name'] = 'dev-user';
}

type DeviceInfo = {
  serialNumber: string,
  intuneDeviceId: string | null,
  azureAdDeviceId: string | null,
  autopilotId: string | null,
  name: string | null,
  model: string | null,
  platform: 'windows' | 'apple' | 'unknown',
  currentEnrollmentProfile: string,
  groupTag: string | null,
  assignedUserPrincipalName: string | null,
  macAddress: string | null,
  altMacAddress: string | null,
  username: string | null,
  email: string | null,
  building: string | null,
  room: string | null,
  assetTag: string | null,
};

function createAlpineData() {
  return {
    theme: Alpine.$persist(process.env.THEME ?? 'dim'),
    searchData: '',
    searchType: Alpine.$persist('windows' as 'windows' | 'apple'),
    errorMessage: '',
    successMessage: '',
    dataList: [] as DeviceInfo[],
    dataListCopy: [] as DeviceInfo[],
    dataIndex: 0,
    updateToProfile: '' as string,
    showProfileDropdown: false,
    confirmModal: {
      title: '',
      lines: [] as string[],
      onConfirm: null as (() => Promise<void>) | null,
    },
    approvalModal: {
      open: false,
      action: '' as 'wipe' | 'retire' | '',
      device: null as DeviceInfo | null,
      justification: '',
      submitting: false,
      error: '',
    },
    approvalsPanel: {
      open: false,
      error: '',
    },
    pendingApprovals: [] as any[],
    pendingCount: 0,
    auditLog: [] as any[],
    auditLogOpen: false,
    _pollInterval: null as any,
    // SKIP_ENTRA_AUTH is inlined at build time by Bun (env: 'inline' in worker.ts)
    skipEntraAuth: process.env.SKIP_ENTRA_AUTH === 'true',

    get currentData() {
      return this.dataList[this.dataIndex] || {};
    },

    async init() {
      // Wait for a real session before polling an authenticated endpoint —
      // otherwise this fires immediately on page load, before sign-in completes.
      await authReady;
      await this.pollPending();
      this._pollInterval = setInterval(() => this.pollPending(), 30000);
    },

    async pollPending() {
      try {
        const resp = await axios.get('/approvals/pending');
        this.pendingApprovals = resp.data.items ?? [];
        this.pendingCount = resp.data.count ?? 0;
      } catch { /* non-fatal */ }
    },

    async loadAuditLog() {
      try {
        const resp = await axios.get('/audit-log?limit=100');
        this.auditLog = resp.data;
        this.auditLogOpen = true;
        (document.getElementById('auditLogDialog') as HTMLDialogElement).showModal();
      } catch { /* non-fatal */ }
    },

    async submitApproval() {
      const modal = this.approvalModal;
      if (!modal.device || !modal.justification.trim()) { modal.error = 'A justification is required.'; return; }
      modal.submitting = true;
      modal.error = '';
      try {
        const device = modal.device;
        await axios.post('/approvals', {
          action: modal.action,
          justification: modal.justification.trim(),
          deviceSerial: device.serialNumber,
          deviceId: device.intuneDeviceId ?? undefined,
          deviceAssetTag: device.assetTag ?? undefined,
          payload: { deviceId: device.intuneDeviceId, serialNumber: device.serialNumber, macAddress: device.macAddress, altMacAddress: device.altMacAddress },
        });
        (document.getElementById('approvalRequestDialog') as HTMLDialogElement).close();
        this.successMessage = `${modal.action === 'wipe' ? 'Wipe' : 'Retire'} request submitted. Awaiting second admin approval.`;
        await this.pollPending();
      } catch (err: any) {
        modal.error = err.response?.data?.error ?? 'Failed to submit request.';
      } finally {
        modal.submitting = false;
      }
    },

    async approveRequest(id: number) {
      const panel = this.approvalsPanel;
      panel.error = '';
      try {
        await axios.post(`/approvals/${id}/approve`);
        await this.pollPending();
        this.successMessage = 'Action approved and executed.';
        if (this.pendingCount === 0) (document.getElementById('approvalsDialog') as HTMLDialogElement).close();
      } catch (err: any) {
        panel.error = err.response?.data?.error ?? 'Failed to approve.';
      }
    },

    async rejectRequest(id: number) {
      const panel = this.approvalsPanel;
      panel.error = '';
      try {
        await axios.post(`/approvals/${id}/reject`);
        await this.pollPending();
        this.successMessage = 'Request rejected.';
        if (this.pendingCount === 0) (document.getElementById('approvalsDialog') as HTMLDialogElement).close();
      } catch (err: any) {
        panel.error = err.response?.data?.error ?? 'Failed to reject.';
      }
    },

    prev() {
      this.dataIndex = (this.dataIndex - 1 + this.dataList.length) % this.dataList.length;
      this.updateToProfile = '';
      this.showProfileDropdown = false;
    },
    next() {
      this.dataIndex = (this.dataIndex + 1) % this.dataList.length;
      this.updateToProfile = '';
      this.showProfileDropdown = false;
    },

    async search() {
      try {
        if (!this.searchData) {
          return;
        }

        const encodedSearch = encodeURIComponent(this.searchData.trim());
        const response = await axios.get(`/devices/${encodedSearch}`)
          .catch((error: any) => {
            console.error('Error fetching data:', error.response?.data || error.message);
            this.errorMessage = `An error occurred while searching for data. Error: ${error.response?.status ?? 'unknown'}`;
            throw error;
          });

        this.dataList = response.data;
        this.dataListCopy = JSON.parse(JSON.stringify(this.dataList));

        this.dataIndex = 0;
        this.errorMessage = '';
        this.successMessage = '';
        this.showProfileDropdown = false;
        this.updateToProfile = '';

        Alpine.nextTick(() => {
          const nextInput = document.querySelector('.datafield-input:not(:disabled)');
          if (nextInput instanceof HTMLInputElement) {
            nextInput.focus();
          }
        });
      } catch (error: any) {
        if (error.response && error.response.status === 404) {
          this.errorMessage = `No device found for: ${this.searchData}`;
        } else {
          this.errorMessage = `An error occurred while searching for data. Error: ${error.response?.status ?? 'unknown'}`;
        }
        this.dataList = [];
        this.dataListCopy = [];
        this.dataIndex = 0;
      }
    },

    async send() {
      const original = this.dataListCopy[this.dataIndex];
      const current = this.dataList[this.dataIndex];
      if (!current) { this.errorMessage = 'No data to update.'; this.successMessage = ''; return; }

      const EDITABLE = ['username', 'email', 'building', 'room', 'assetTag'] as const;
      const fieldLines: string[] = EDITABLE
        .filter(k => String((current as any)[k] ?? '') !== String((original as any)[k] ?? ''))
        .map(k => `${k}: "${(original as any)[k] ?? ''}" → "${(current as any)[k] ?? ''}"`);

      const hasProfileUpdate = this.updateToProfile !== ''
        && String(current.currentEnrollmentProfile ?? '') !== String(original.currentEnrollmentProfile ?? '');
      const profileLines: string[] = hasProfileUpdate
        ? [`Enrollment profile: "${original.currentEnrollmentProfile}" → "${current.currentEnrollmentProfile}"`]
        : [];

      const lines = [...fieldLines, ...profileLines];
      if (lines.length === 0) { this.errorMessage = 'No changes to update.'; this.successMessage = ''; return; }

      this.showConfirm('Confirm Changes', lines, async () => {
        try {
          if (hasProfileUpdate) {
            await axios.post(`/change-enrollment-profile/${this.searchType}/${this.updateToProfile}/${current.serialNumber}`);
          }
          if (fieldLines.length > 0) {
            Object.keys(current).forEach(key => {
              if ((current as any)[key] === null) (current as any)[key] = '';
            });
            await axios.put(`/device-metadata/${encodeURIComponent(current.serialNumber)}`, {
              username: current.username,
              email: current.email,
              building: current.building,
              room: current.room,
              assetTag: current.assetTag,
            });
          }
          this.dataList[this.dataIndex] = { ...current };
          this.dataListCopy[this.dataIndex] = { ...current };
          this.updateToProfile = '';
          this.errorMessage = '';
          this.successMessage = 'Data updated successfully.';
        } catch (error: any) {
          this.errorMessage = `An error occurred while sending data. Error: ${error.response?.status ?? 'unknown'}`;
        }
      });
    },

    async erase() {
      const current = this.dataList[this.dataIndex];
      if (!current) { this.errorMessage = 'No data to wipe.'; this.successMessage = ''; return; }
      this.approvalModal.action = 'wipe';
      this.approvalModal.device = current;
      this.approvalModal.justification = '';
      this.approvalModal.error = '';
      (document.getElementById('approvalRequestDialog') as HTMLDialogElement).showModal();
    },

    async retire() {
      const current = this.dataList[this.dataIndex];
      if (!current) { this.errorMessage = 'No data to retire.'; this.successMessage = ''; return; }
      this.approvalModal.action = 'retire';
      this.approvalModal.device = current;
      this.approvalModal.justification = '';
      this.approvalModal.error = '';
      (document.getElementById('approvalRequestDialog') as HTMLDialogElement).showModal();
    },

    showConfirm(title: string, lines: string[], callback: () => Promise<void>) {
      this.confirmModal.title = title;
      this.confirmModal.lines = lines;
      this.confirmModal.onConfirm = callback;
      (document.getElementById('confirmDialog') as HTMLDialogElement).showModal();
    },
  }
}

function fetchEnrollmentProfiles(params?: { getSearchType: () => 'windows' | 'apple'; getDataList: () => any[] }) {
  return {
    profiles: [],
    currentSearchType: '',

    async init(this: any) {
      if (params && typeof params.getDataList === 'function' && typeof params.getSearchType === 'function') {
        this.$watch(() => params.getDataList(), async (dataList: any[]) => {
          if (Array.isArray(dataList) && dataList.length > 0) {
            const searchType = params.getSearchType() || 'windows';
            if (searchType !== this.currentSearchType || this.profiles.length === 0) {
              this.currentSearchType = searchType;
              await this.loadProfiles(searchType);
            }
          }
        });
      }
    },

    async loadProfiles(searchType: string) {
      try {
        const response: AxiosResponse = await axios.get(`/enrollment-profiles/${searchType}`);
        response.data.sort((a: { displayName: string; }, b: { displayName: string; }) => a.displayName.localeCompare(b.displayName));
        this.profiles = response.data;
        console.log(`Loaded ${this.profiles.length} ${searchType} enrollment profiles`);
      } catch (error: any) {
        console.error('Error fetching enrollment profiles:', error.response?.data || error.message);
        this.profiles = [];
      }
    }
  }
}

// @ts-ignore
window.Alpine = Alpine;
Alpine.store('skipEntraAuth', process.env.SKIP_ENTRA_AUTH === 'true');

// Register Alpine components
Alpine.data('AzureAuth', AzureAuth);
Alpine.data('FetchEnrollmentProfiles', fetchEnrollmentProfiles);
Alpine.data('AlpineData', createAlpineData);

// Import Alpine plugins
Alpine.plugin(AlpinePersist);

// Start Alpine.js
Alpine.start();
