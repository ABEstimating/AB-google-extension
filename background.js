console.log('[AB] Service Worker started');
console.log('[AB] Service Worker version 2025-10-30-3');
// top of background.js
self.importScripts('xlsx.full.min.js');
self.XLSX_LOADED = typeof XLSX !== 'undefined';

try { importScripts('division_mapping_helper.js'); } catch (e) { console.warn('division-mapping helper not loaded', e); }

// load campaign automation worker (MV3 service worker allows importScripts)
try { importScripts('campaigns-bg.js'); } catch(e) { console.warn('campaigns-bg not loaded', e); }
try { importScripts('excel-summary-bg.js'); } catch(e) { console.warn('excel-summary not loaded', e); }


try { importScripts('static_project_cache.js'); } catch (e) { console.warn('[StaticCache] static_project_cache.js not loaded', e); }

// ============ INDEXEDDB DATABASE ============
const DB_NAME = 'AutoBuildersDB';
const DB_VERSION = 5;
const AUTO_SCAN_ENABLED = false;
const ENABLE_VERSION_HISTORY_LOOKUP = true;
let detachedPanelWindowId = null;
let detachedPanelTabId = null;
let suppressNextAttach = false;

if (!AUTO_SCAN_ENABLED && typeof chrome !== 'undefined' && chrome.alarms && typeof chrome.alarms.clear === 'function') {
  try {
    chrome.alarms.clear('AB_AUTO_SCAN').catch(() => {});
  } catch (error) {
    // ignore errors clearing alarm at startup
  }
}


chrome.windows.onRemoved.addListener(windowId => {
  if (windowId === detachedPanelWindowId) {
    detachedPanelWindowId = null;
    if (suppressNextAttach) {
      suppressNextAttach = false;
      return;
    }
    if (detachedPanelTabId) {
      chrome.tabs.sendMessage(detachedPanelTabId, { type: 'AB_PANEL_ATTACH' }).catch(() => {});
    }
  }
});

chrome.tabs.onRemoved.addListener(tabId => {
  if (tabId === detachedPanelTabId) {
    detachedPanelTabId = null;
  }
});

const DB = {
  db: null,

  async init() {
    if (this.db) return this.db;

    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        this.db = request.result;
        resolve(this.db);
      };

      request.onupgradeneeded = (event) => {
        const db = event.target.result;

        if (!db.objectStoreNames.contains('projects')) {
          db.createObjectStore('projects', { keyPath: 'id', autoIncrement: true });
        }

        if (!db.objectStoreNames.contains('contacts')) {
          const contactStore = db.createObjectStore('contacts', { keyPath: 'id', autoIncrement: true });
          contactStore.createIndex('company', 'company', { unique: false });
          contactStore.createIndex('state', 'state', { unique: false });
          contactStore.createIndex('division', 'division', { unique: false });
        }

        if (!db.objectStoreNames.contains('metadata')) {
          db.createObjectStore('metadata', { keyPath: 'key' });
        }

        let auditStore = null;
        const upgradeTx = event.target.transaction || null;

        if (!db.objectStoreNames.contains('contactAudit')) {
          auditStore = db.createObjectStore('contactAudit', { keyPath: 'id', autoIncrement: true });
        } else if (upgradeTx) {
          try {
            auditStore = upgradeTx.objectStore('contactAudit');
          } catch (error) {
            console.warn('[DB] Failed to access contactAudit during upgrade:', error);
          }
        }

        if (auditStore) {
          if (!auditStore.indexNames.contains('detectedAt')) {
            auditStore.createIndex('detectedAt', 'detectedAt', { unique: false });
          }
          if (auditStore.indexNames.contains('byWorkbookEmail')) {
            try {
              auditStore.deleteIndex('byWorkbookEmail');
            } catch (error) {
              console.warn('[DB] Failed to remove legacy audit index:', error);
            }
          }
          try {
            auditStore.createIndex('byWorkbookEmail', ['workbookKey', 'emailNormalized'], { unique: false });
          } catch (error) {
            console.warn('[DB] Failed to create audit workbook/email index:', error);
          }
        }
        let trackedStore = null;
        if (!db.objectStoreNames.contains('trackedAudit')) {
          trackedStore = db.createObjectStore('trackedAudit', { keyPath: 'id', autoIncrement: true });
        } else if (upgradeTx) {
          try {
            trackedStore = upgradeTx.objectStore('trackedAudit');
          } catch (error) {
            console.warn('[DB] Failed to access trackedAudit during upgrade:', error);
          }
        }

        if (trackedStore) {
          const ensureIndex = (name, keyPath, options = {}) => {
            if (!trackedStore.indexNames.contains(name)) {
              try {
                trackedStore.createIndex(name, keyPath, options);
              } catch (error) {
                console.warn('[DB] Failed to create trackedAudit index', name, error);
              }
            }
          };

          ensureIndex('byProject', 'projectKey');
          ensureIndex('byProjectDate', ['projectKey', 'detectedAt']);
          ensureIndex('byDate', 'detectedAt');
          ensureIndex('byWorkbook', 'workbookKey');
          ensureIndex('byWorkbookEmailDate', ['workbookKey', 'emailNormalized', 'detectedAt']);
        }
      };
    });
  },

  async saveProjects(projects) {
    const db = await this.init();
    const tx = db.transaction(['projects'], 'readwrite');
    const store = tx.objectStore('projects');

    await store.clear();

    for (const project of projects) {
      await store.add(project);
    }

    return tx.complete;
  },

  async saveContacts(contacts) {
    const db = await this.init();
    const tx = db.transaction(['contacts'], 'readwrite');
    const store = tx.objectStore('contacts');

    await store.clear();

    for (const contact of contacts) {
      await store.add(contact);
    }

    return tx.complete;
  },

  async getAllProjects() {
    const db = await this.init();
    const tx = db.transaction(['projects'], 'readonly');
    const store = tx.objectStore('projects');

    return new Promise((resolve, reject) => {
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  },

  async getAllContacts() {
    const db = await this.init();
    const tx = db.transaction(['contacts'], 'readonly');
    const store = tx.objectStore('contacts');

    return new Promise((resolve, reject) => {
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  },

  async saveMetadata(data) {
    const db = await this.init();
    const tx = db.transaction(['metadata'], 'readwrite');
    const store = tx.objectStore('metadata');

    await store.put({ key: 'settings', ...data });
    return tx.complete;
  },

  async getMetadata() {
    const db = await this.init();
    const tx = db.transaction(['metadata'], 'readonly');
    const store = tx.objectStore('metadata');

    return new Promise((resolve, reject) => {
      const request = store.get('settings');
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  },

  async appendContactAuditEntries(entries = []) {
    if (!Array.isArray(entries) || entries.length === 0) {
      return { inserted: 0 };
    }

    const db = await this.init();
    const tx = db.transaction(['contactAudit'], 'readwrite');
    const store = tx.objectStore('contactAudit');
    let index = null;
    try {
      index = store.index('byWorkbookEmail');
    } catch (_) {
      index = null;
    }

    let inserted = 0;

    const normalizeEntry = (raw) => {

      if (!raw || typeof raw !== 'object') return null;



      const emailNormalized = String(raw.emailNormalized || raw.email || '').trim().toLowerCase();

      if (!emailNormalized) return null;



      const candidateKeys = [

        raw.workbookKey,

        raw.workbookId,

        raw.workbookPath,

        raw.workbookName && raw.projectKey ? `${raw.projectKey}::${raw.workbookName}` : null,

        raw.workbookName,

        raw.projectKey

      ];



      let workbookKey = candidateKeys.find(value => {

        if (value === null || value === undefined) return false;

        const str = String(value).trim();

        return str.length > 0;

      });

      workbookKey = workbookKey ? String(workbookKey).trim().toLowerCase() : emailNormalized;



      let detectedAt = raw.detectedAt;

      try {

        detectedAt = detectedAt ? new Date(detectedAt).toISOString() : new Date().toISOString();

      } catch (_) {

        detectedAt = new Date().toISOString();

      }



      let lastModifiedDateTime = raw.lastModifiedDateTime || null;

      if (lastModifiedDateTime) {

        try {

          lastModifiedDateTime = new Date(lastModifiedDateTime).toISOString();

        } catch (_) {

          lastModifiedDateTime = String(raw.lastModifiedDateTime).trim() || null;

        }

      }



      const lastModifiedBy = raw.lastModifiedBy ? String(raw.lastModifiedBy).trim() : 'Unknown';



      const normalizedProjectKey = (() => {

        const projectCandidates = [

          raw.projectKey,

          raw.projectName,

          raw.projectDisplayName,

          raw.projectFolderName,

          raw.project

        ];

        for (const candidate of projectCandidates) {

          if (candidate === null || candidate === undefined) continue;

          const trimmed = String(candidate).trim();

          if (trimmed) {

            const key = getProjectCacheKey(trimmed);

            if (key) return key;

          }

        }

        return null;

      })();



      const finalProjectKey = (() => {

        if (normalizedProjectKey && normalizedProjectKey.trim()) return normalizedProjectKey;

        if (raw.projectKey) {

          const trimmed = String(raw.projectKey).trim();

          if (trimmed) return trimmed.toLowerCase();

        }

        return null;

      })();



      const source = raw.source ? String(raw.source).trim() || 'scan' : 'scan';



      return {

        ...raw,

        projectKey: finalProjectKey,

        lastModifiedBy,

        lastModifiedDateTime,

        detectedAt,

        emailNormalized,

        workbookKey,

        source

      };

    };



    const addPromises = entries.map(entry => {
      const record = normalizeEntry(entry);
      if (!record) return Promise.resolve();

      return new Promise((resolve) => {
        const addRecord = () => {
          try {
            const request = store.add(record);
            request.onsuccess = () => {
              inserted += 1;
              resolve();
            };
            request.onerror = () => {
              console.warn('[DB] Failed to append audit entry', request.error, record);
              resolve();
            };
          } catch (error) {
            console.warn('[DB] Unexpected error appending audit entry', error, record);
            resolve();
          }
        };

        if (!index) {
          addRecord();
          return;
        }

        let lookupRequest;
        try {
          lookupRequest = index.get([record.workbookKey, record.emailNormalized]);
        } catch (error) {
          console.warn('[DB] Audit lookup failed', error, record);
          addRecord();
          return;
        }

        lookupRequest.onsuccess = () => {
          if (lookupRequest.result) {
            resolve();
            return;
          }
          addRecord();
        };

        lookupRequest.onerror = () => {
          console.warn('[DB] Audit lookup error', lookupRequest.error, record);
          addRecord();
        };
      });
    });

    await Promise.all(addPromises);

    await new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });

    return { inserted };
  },

  async appendTrackedAuditEntries(entries = []) {
    if (!Array.isArray(entries) || entries.length === 0) {
      return { inserted: 0 };
    }

    const db = await this.init();
    const tx = db.transaction(['trackedAudit'], 'readwrite');
    const store = tx.objectStore('trackedAudit');
    let index = null;
    try {
      index = store.index('byWorkbookEmailDate');
    } catch (_) {
      index = null;
    }

    let inserted = 0;

    const normalizeEntry = (raw) => {
      if (!raw || typeof raw !== 'object') return null;

      const emailNormalized = String(raw.emailNormalized || raw.email || '').trim().toLowerCase();
      if (!emailNormalized) return null;

      const projectKey = String(raw.projectKey || raw.projectFolderName || raw.projectName || '').trim().toLowerCase();
      if (!projectKey) return null;

      let workbookKey = raw.workbookKey || raw.workbookId || raw.workbookPath || null;
      if (!workbookKey || String(workbookKey).trim().length === 0) {
        workbookKey = `${projectKey}::${emailNormalized}`;
      } else {
        workbookKey = String(workbookKey).trim().toLowerCase();
      }

      let detectedAt = raw.detectedAt;
      try {
        detectedAt = detectedAt ? new Date(detectedAt).toISOString() : new Date().toISOString();
      } catch (_) {
        detectedAt = new Date().toISOString();
      }

      let lastModifiedDateTime = raw.lastModifiedDateTime || null;
      if (lastModifiedDateTime) {
        try {
          lastModifiedDateTime = new Date(lastModifiedDateTime).toISOString();
        } catch (_) {
          lastModifiedDateTime = String(raw.lastModifiedDateTime).trim() || null;
        }
      }

      const company = raw.company || raw.companyName || '';
      const divisionName = raw.divisionName || raw.division || raw.divisionKey || '';
      const state = raw.state || raw.projectState || null;

      return {
        ...raw,
        projectKey,
        projectName: raw.projectName || raw.projectDisplayName || raw.project || '',
        workbookKey,
        workbookId: raw.workbookId || null,
        workbookName: raw.workbookName || null,
        workbookPath: raw.workbookPath || null,
        email: raw.email || '',
        emailNormalized,
        company,
        divisionName,
        divisionKey: raw.divisionKey || null,
        contactName: raw.contactName || raw.name || '',
        phone: raw.phone || '',
        detectedAt,
        lastModifiedBy: raw.lastModifiedBy ? String(raw.lastModifiedBy).trim() : 'Unknown',
        lastModifiedDateTime,
        state,
        source: raw.source ? String(raw.source).trim() || 'scan' : 'scan',
        createdAt: new Date().toISOString()
      };
    };

    const addPromises = entries.map(entry => {
      const record = normalizeEntry(entry);
      if (!record) return Promise.resolve();

      return new Promise(resolve => {
        const addRecord = () => {
          try {
            const request = store.add(record);
            request.onsuccess = () => {
              inserted += 1;
              resolve();
            };
            request.onerror = () => {
              console.warn('[DB] Failed to append tracked audit entry', request.error, record);
              resolve();
            };
          } catch (error) {
            console.warn('[DB] Unexpected error appending tracked audit entry', error, record);
            resolve();
          }
        };

        if (!index) {
          addRecord();
          return;
        }

        let lookupRequest;
        try {
          lookupRequest = index.get([record.workbookKey, record.emailNormalized, record.detectedAt]);
        } catch (error) {
          console.warn('[DB] Tracked audit lookup failed', error, record);
          addRecord();
          return;
        }

        lookupRequest.onsuccess = () => {
          if (lookupRequest.result) {
            resolve();
            return;
          }
          addRecord();
        };

        lookupRequest.onerror = () => {
          console.warn('[DB] Tracked audit lookup error', lookupRequest.error, record);
          addRecord();
        };
      });
    });

    await Promise.all(addPromises);

    await new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });

    return { inserted };
  },

  async getTrackedAuditEntries(options = {}) {
    const { projectKey = null, projectKeys = null, start = null, end = null, limit = 2000 } = options || {};

    const keySet = new Set();
    if (Array.isArray(projectKeys)) {
      projectKeys.filter(Boolean).forEach(value => keySet.add(String(value).trim().toLowerCase()));
    }
    if (projectKey) {
      keySet.add(String(projectKey).trim().toLowerCase());
    }
    const hasProjectFilter = keySet.size > 0;

    const db = await this.init();
    const tx = db.transaction(['trackedAudit'], 'readonly');
    const store = tx.objectStore('trackedAudit');

    const normalizeIso = (value, isEnd = false) => {
      if (!value) return null;
      try {
        if (typeof value === 'string') {
          if (value.includes('T')) {
            return value.endsWith('Z') ? value : `${value}Z`;
          }
          return `${value}${isEnd ? 'T23:59:59.999Z' : 'T00:00:00.000Z'}`;
        }
        if (value instanceof Date) {
          return value.toISOString();
        }
      } catch (error) {
        console.warn('[Audit] Failed to normalize tracked audit date', value, error);
      }
      return String(value);
    };

    const lower = normalizeIso(start, false);
    const upper = normalizeIso(end, true);

    const projectKeysArray = Array.from(keySet);
    const singleProject = projectKeysArray.length === 1 ? projectKeysArray[0] : null;

    let index = null;
    let range = null;

    if (singleProject) {
      try {
        index = store.index('byProjectDate');
        if (lower && upper) {
          range = IDBKeyRange.bound([singleProject, lower], [singleProject, upper]);
        } else if (lower) {
          range = IDBKeyRange.lowerBound([singleProject, lower]);
        } else if (upper) {
          range = IDBKeyRange.upperBound([singleProject, upper]);
        } else {
          range = IDBKeyRange.bound([singleProject, ''], [singleProject, '\\uffff']);
        }
      } catch (error) {
        console.warn('[DB] Tracked audit byProjectDate index failed', error);
        index = null;
      }
    }

    if (!index) {
      try {
        index = store.index('byDate');
      } catch (_) {
        index = store;
      }

      if (lower && upper) {
        range = IDBKeyRange.bound(lower, upper);
      } else if (lower) {
        range = IDBKeyRange.lowerBound(lower);
      } else if (upper) {
        range = IDBKeyRange.upperBound(upper);
      }
    }

    return await new Promise((resolve, reject) => {
      const results = [];
      const request = index.openCursor(range, 'prev');

      request.onsuccess = (event) => {
        const cursor = event.target.result;
        if (!cursor) {
          resolve(results);
          return;
        }

        const value = cursor.value;
        if (hasProjectFilter && !keySet.has(String(value.projectKey || '').trim().toLowerCase())) {
          cursor.continue();
          return;
        }

        results.push(value);
        if (results.length >= limit) {
          resolve(results);
          return;
        }

        cursor.continue();
      };

      request.onerror = () => reject(request.error);
    });
  },

  async replaceTrackedAuditEntries(entries = []) {
    const db = await this.init();
    const tx = db.transaction(['trackedAudit'], 'readwrite');
    const store = tx.objectStore('trackedAudit');

    await new Promise((resolve, reject) => {
      const clearRequest = store.clear();
      clearRequest.onsuccess = () => resolve();
      clearRequest.onerror = () => reject(clearRequest.error);
    });

    if (Array.isArray(entries)) {
      entries.forEach(entry => {
        try {
          store.add(entry);
        } catch (error) {
          console.warn('[DB] Failed to restore tracked audit entry', error, entry);
        }
      });
    }

    await new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  },

  async deleteTrackedAuditForProjects(projectKeys = []) {
    if (!Array.isArray(projectKeys) || projectKeys.length === 0) {
      return 0;
    }

    const normalizedKeys = projectKeys
      .map(key => getProjectCacheKey(key || ''))
      .filter(key => typeof key === 'string' && key.length);

    if (!normalizedKeys.length) {
      return 0;
    }

    const db = await this.init();
    const tx = db.transaction(['trackedAudit'], 'readwrite');
    const store = tx.objectStore('trackedAudit');
    let deleted = 0;

    const deleteForKey = (projectKey) => {
      return new Promise((resolve, reject) => {
        let cursorRequest;
        try {
          const index = store.index('byProject');
          cursorRequest = index.openCursor(IDBKeyRange.only(projectKey));
        } catch (error) {
          reject(error);
          return;
        }

        cursorRequest.onsuccess = (event) => {
          const cursor = event.target.result;
          if (cursor) {
            cursor.delete();
            deleted += 1;
            cursor.continue();
          } else {
            resolve();
          }
        };

        cursorRequest.onerror = () => reject(cursorRequest.error);
      });
    };

    for (const key of normalizedKeys) {
      await deleteForKey(key);
    }

    await new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });

    return deleted;
  },

  async getAllTrackedAuditEntries() {
    const db = await this.init();
    const tx = db.transaction(['trackedAudit'], 'readonly');
    const store = tx.objectStore('trackedAudit');

    return await new Promise((resolve, reject) => {
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  },


  async getContactAuditEntries(options = {}) {
    const { start, end, projectKey, user, limit = 1000 } = options || {};
    const db = await this.init();
    const tx = db.transaction(['contactAudit'], 'readonly');
    const store = tx.objectStore('contactAudit');
    let index;
    try {
      index = store.index('detectedAt');
    } catch (_) {
      index = store;
    }

    const normalizeIso = (value, isEnd = false) => {
      if (!value) return null;
      try {
        if (typeof value === 'string') {
          if (value.includes('T')) {
            return value.endsWith('Z') ? value : `${value}Z`;
          }
          return `${value}${isEnd ? 'T23:59:59.999Z' : 'T00:00:00.000Z'}`;
        }
        if (value instanceof Date) {
          return value.toISOString();
        }
      } catch (error) {
        console.warn('[Audit] Failed to normalize date value', value, error);
      }
      return String(value);
    };

    const lower = normalizeIso(start, false);
    const upper = normalizeIso(end, true);
    let range = null;
    if (lower && upper) {
      range = IDBKeyRange.bound(lower, upper);
    } else if (lower) {
      range = IDBKeyRange.lowerBound(lower);
    } else if (upper) {
      range = IDBKeyRange.upperBound(upper);
    }

    return await new Promise((resolve, reject) => {
      const results = [];
      const request = index.openCursor(range, 'prev');

      request.onsuccess = (event) => {
        const cursor = event.target.result;
        if (!cursor) {
          resolve(results);
          return;
        }

        const value = cursor.value;
        if (projectKey && value.projectKey !== projectKey) {
          cursor.continue();
          return;
        }
        if (user && (value.lastModifiedBy || '').toLowerCase() !== user.toLowerCase()) {
          cursor.continue();
          return;
        }

        results.push(value);
        if (results.length >= limit) {
          resolve(results);
          return;
        }

        cursor.continue();
      };

      request.onerror = () => reject(request.error);
    });
  },

  clearAll: async function () {
    const db = await this.init();
    const tx = db.transaction(['projects', 'contacts', 'metadata', 'contactAudit'], 'readwrite');

    await tx.objectStore('projects').clear();
    await tx.objectStore('contacts').clear();
    await tx.objectStore('metadata').clear();
    try {
      await tx.objectStore('contactAudit').clear();
    } catch (error) {
      console.warn('[DB] contactAudit clear skipped:', error);
    }

    return tx.complete;
  }

};

// ============ END INDEXEDDB DATABASE ============

const CACHE_FOLDER_SEGMENTS = ['Z - Extension Cache', 'Database'];
const CACHE_MANIFEST_NAME = 'manifest.json';

// Load XLSX library - MUST be loaded synchronously at the top
let XLSX_LOADED = false;

try {
  importScripts('xlsx.full.min.js');
if (!self.AB_CAMPAIGN_BG_LOADED) importScripts('campaigns-bg.js');

  XLSX_LOADED = true;
  console.log('XLSX library loaded successfully');
} catch (error) {
  console.error('Failed to load XLSX library:', error);
  console.error('Make sure xlsx.full.min.js is in the extension root directory');
}

// Verify XLSX is actually loaded
if (typeof XLSX !== 'undefined' && XLSX.read && XLSX.utils) {
  XLSX_LOADED = true;
  console.log('XLSX library verified and ready');
} else {
  console.error('XLSX library not properly loaded or incomplete');
  XLSX_LOADED = false;
}

// Azure AD Configuration
const AZURE_CONFIG = {
  clientId: 'a9be13b9-b750-4dd6-a28b-9bf0169c44bb',
  tenantId: '73403434-959f-42fd-990a-9b774abe1489',
  redirectUri: chrome.identity.getRedirectURL(),
  scopes: [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Mail.Send',
    'https://graph.microsoft.com/Mail.Send.Shared',
    'offline_access'
  ]
};;

const GRAPH_API = 'https://graph.microsoft.com/v1.0';

const DELTA_KEY_PREFIX = 'delta:';
const DELTA_FILE_INDEX_KEY = `${DELTA_KEY_PREFIX}fileIndex`;

function sanitizeStorageToken(value) {
  if (value === undefined || value === null) return 'unknown';
  const token = String(value).trim();
  if (!token) return 'unknown';
  return token.replace(/[^a-zA-Z0-9]/g, '_') || 'unknown';
}

function buildDeltaStorageKey(folderRef = {}) {
  const siteToken = sanitizeStorageToken(folderRef.siteId || folderRef.siteUrl || folderRef.site || folderRef.siteName || 'default');
  const driveToken = sanitizeStorageToken(folderRef.driveId || folderRef.drive || 'drive');
  const folderToken = sanitizeStorageToken(folderRef.folderId || folderRef.id || 'folder');
  return `${DELTA_KEY_PREFIX}${siteToken}:${driveToken}:${folderToken}`;
}

const deltaStore = {
  async getDeltaLink(folderRef = {}) {
    const key = buildDeltaStorageKey(folderRef);
    const stored = await chrome.storage.local.get(key);
    return stored[key] || null;
  },

  async setDeltaLink(folderRef = {}, deltaLink) {
    const key = buildDeltaStorageKey(folderRef);
    if (!deltaLink) {
      await chrome.storage.local.remove(key);
    } else {
      await chrome.storage.local.set({ [key]: deltaLink });
    }
  },

  async clearDeltaLink(folderRef = {}) {
    const key = buildDeltaStorageKey(folderRef);
    await chrome.storage.local.remove(key);
  },

  async getFileIndexSnapshot() {
    const stored = await chrome.storage.local.get(DELTA_FILE_INDEX_KEY);
    const snapshot = stored[DELTA_FILE_INDEX_KEY];
    if (!snapshot || typeof snapshot !== 'object') {
      return {};
    }
    return { ...snapshot };
  },

  async saveFileIndexSnapshot(snapshot = {}) {
    await chrome.storage.local.set({ [DELTA_FILE_INDEX_KEY]: snapshot });
  },

  async updateFileEntry(itemId, metadata = {}) {
    if (!itemId) return;
    const snapshot = await this.getFileIndexSnapshot();
    snapshot[itemId] = { ...(snapshot[itemId] || {}), ...metadata };
    await this.saveFileIndexSnapshot(snapshot);
  },

  async removeFileEntry(itemId) {
    if (!itemId) return;
    const snapshot = await this.getFileIndexSnapshot();
    if (snapshot[itemId]) {
      delete snapshot[itemId];
      await this.saveFileIndexSnapshot(snapshot);
    }
  }
};

const SHAREPOINT_USER_OVERRIDES = {
  'bandrews@autobuilders.net': 'Blake Andrews',
  'sdepkon@autobuilders.net': 'Sharon Depkon',
  'rbrown@autobuilders.net': 'Rick Brown'
};

const SHAREPOINT_SITES = [
  'autobuildersbpos.sharepoint.com:/sites/estimating',
  'autobuildersbpos.sharepoint.com:/sites/questestimatingfolder'
];


const CACHE_BOOTSTRAP_DELAY_MINUTES = 1;
const PERIODIC_SCAN_INTERVAL_MINUTES = 30;
const BASELINE_MAX_AGE_HOURS = 12;
const INCREMENTAL_MAX_AGE_MINUTES = 60;
const SCAN_ALARM_NAME = 'AB_AUTO_SCAN';
const AUDIT_DAILY_ALARM_NAME = 'AB_CONTACT_AUDIT_DAILY';
const AUDIT_DAILY_RUN_HOUR = 23;
const AUDIT_DAILY_RUN_MINUTE = 55;
const MAX_AUDIT_FETCH_LIMIT = 5000;
const AUDIT_LOG_DEFAULT_LIMIT = 1000;


function normalizeProjectName(value) {
  if (!value) return '';
  return String(value)
    .toLowerCase()
    .replace(/projects?$/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function compactProjectName(value) {
  return normalizeProjectName(value).replace(/\s+/g, '');
}

function parseIsoDate(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function minutesToMs(minutes) {
  return Math.max(0, Number(minutes || 0)) * 60 * 1000;
}

console.log('Extension ID:', chrome.runtime.id);
console.log('Expected Redirect URI:', chrome.identity.getRedirectURL());

let accessToken = null;
let tokenExpiry = null;
let cachedProjects = [];
let cachedContacts = [];
let staticProjectsCache = [];
let staticContactsCache = [];
let staticZipLocationCache = {};
const staticProjectKeySet = new Set();
const staticActiveProjectKeySet = new Set();
const STATIC_CACHE_SOURCE = (typeof self !== 'undefined' && self.STATIC_PROJECT_CACHE) ? self.STATIC_PROJECT_CACHE : null;
const PROJECT_CACHE_LIMIT = 5000;
const PROJECT_CODE_PATTERN = /(\d{2})[-\s_]?(\d{2,3})/;
let lastScanTime = null;
let lastBaselineScan = null;
let lastIncrementalScan = null;
let currentProjectFile = null;
let currentWorkbookSession = null;
let deltaScanInProgress = false;
let trackedEntriesScanInProgress = false;
let divisionMapping = {};

const EXCLUDED_SHEETS = ['Spreadsheet','Summary','Cover','Key','Bid Results','Notes','Template','Instructions','Index','Equipment','Shop Equipment'];
const SPECIAL_SHEETS = []; // sheets whose data starts on row 4

const FALLBACK_DIVISION_ORDER = [
  'Demo', 'Demo & Saw Cutting', 'Demolition',
  'Site Concrete', 'Concrete', 'Concrete, Reinforcement & Formwork',
  'Masonry', 'Masonry & Precast',
  'Steel', 'Misc Steel', 'Structural Steel',
  'Rough Carpentry', 'Millwork', 'Casework', 'Wood & Plastics',
  'Waterproofing', 'Roofing', 'Insulation', 'Thermal & Moisture',
  'Storefront', 'Glazing', 'Doors & Windows', 'Doors and Windows',
  'Drywall', 'Metal Framing & Drywall', 'Metal Framing and Drywall',
  'Flooring', 'Tile', 'Acoustical Ceilings', 'Ceilings',
  'Paint', 'Painting', 'Painting & Wallcovering',
  'Specialties', 'Signage', 'Building Signage',
  'Equipment', 'Food Service Equipment', 'Shop Equipment',
  'Furnishings', 'Window Treatments', 'Car Wash Equipment',
  'Plumbing', 'Mechanical', 'HVAC', 'Fire Protection', 'Elec', 'Electrical', 'Fire Alarm'
];
const FALLBACK_START_DIVISION = 'Demo';
const FALLBACK_END_DIVISION = 'Fire Alarm';
const FALLBACK_DIVISION_NAME_SET = new Set(FALLBACK_DIVISION_ORDER.map(name => String(name || '').trim().toLowerCase()));
const MIN_EXPECTED_DIVISIONS = 30;

function isDivisionMappingIncomplete(mapping = null, divisions = null) {
  const mapKeys = mapping && typeof mapping === 'object' ? Object.keys(mapping) : [];
  const divisionList = Array.isArray(divisions) ? divisions : [];
  if (mapKeys.length === 0 || divisionList.length === 0) return true;
  if (mapKeys.length < MIN_EXPECTED_DIVISIONS || divisionList.length < MIN_EXPECTED_DIVISIONS) return true;

  let fallbackMatches = 0;
  for (const key of mapKeys) {
    const normalized = String(key || '').trim().toLowerCase();
    if (FALLBACK_DIVISION_NAME_SET.has(normalized)) {
      fallbackMatches++;
    }
  }

  if (fallbackMatches >= Math.max(5, mapKeys.length - 5)) {
    return true;
  }

  return false;
}

const DIVISION_CODE_MIN = 2 * 1000 + 100; // 02-100
const DIVISION_CODE_MAX = 16 * 1000 + 670; // 16-670

function getDivisionSiteKey(project = null) {
  const siteValue = (project && (project.site || project.siteUrl || project.siteURL)) || (currentProjectFile && currentProjectFile.site) || '';
  if (typeof siteValue === 'string' && siteValue.toLowerCase().includes('quest')) {
    return 'Quest';
  }
  return 'AutoBuilders';
}

function parseDivisionCode(label) {
  if (!label || typeof label !== 'string') return null;
  const match = label.trim().match(/^(\d{2})[-\s]?(\d{2,3})/);
  if (!match) return null;
  return {
    major: parseInt(match[1], 10),
    minor: parseInt(match[2], 10)
  };
}

function getDivisionCodeValue(codeParts) {
  if (!codeParts) return null;
  return codeParts.major * 1000 + codeParts.minor;
}

function isDivisionCodeWithinRange(value) {
  if (value === null || value === undefined) return false;
  return value >= DIVISION_CODE_MIN && value <= DIVISION_CODE_MAX;
}

function getFallbackOrderIndex(sheetName) {
  if (!sheetName) return -1;
  const lower = String(sheetName).trim().toLowerCase();
  return FALLBACK_DIVISION_ORDER.findIndex(name => name.toLowerCase() === lower);
}

const FALLBACK_START_INDEX = getFallbackOrderIndex(FALLBACK_START_DIVISION);
const FALLBACK_END_INDEX = getFallbackOrderIndex(FALLBACK_END_DIVISION);

function isSheetWithinFallbackRange(sheetName) {
  const idx = getFallbackOrderIndex(sheetName);
  if (idx === -1) return false;
  const start = FALLBACK_START_INDEX === -1 ? 0 : FALLBACK_START_INDEX;
  const end = FALLBACK_END_INDEX === -1 ? FALLBACK_DIVISION_ORDER.length - 1 : FALLBACK_END_INDEX;
  return idx >= start && idx <= end;
}

function getDivisionDisplayInfo(sheetName, projectContext = null) {
  const siteKey = getDivisionSiteKey(projectContext);
  let resolvedLabel = null;
  if (typeof resolveDivisionFolder === 'function') {
    try {
      resolvedLabel = resolveDivisionFolder(sheetName, siteKey);
    } catch (error) {
      console.warn('[DivisionMapping] resolveDivisionFolder failed for', sheetName, error);
    }
  }

  const codeParts = parseDivisionCode(resolvedLabel);
  const codeValue = getDivisionCodeValue(codeParts);
  const fallbackIndex = getFallbackOrderIndex(sheetName);

  return {
    displayName: resolvedLabel || sheetName,
    codeValue,
    fallbackIndex
  };
}

const projectRuntimeCache = new Map();
let metadataCache = {};
let zipLocationCache = {};
let trackedProjectsConfig = [];
function getTrackedProjectsConfigSnapshot() {
  return Array.isArray(trackedProjectsConfig)
    ? trackedProjectsConfig.map(entry => ({ ...entry }))
    : [];
}

function resolveTrackedProjectContext(projectKey) {
  if (!projectKey) return null;
  const normalizedKey = getProjectCacheKey(projectKey);
  if (!normalizedKey) return null;
  if (Array.isArray(cachedProjects) && cachedProjects.length) {
    return cachedProjects.find(project => {
      const candidateKey = project.projectKey
        || getProjectCacheKey(project.projectFolderName || project.displayName || project.name || '');
      return candidateKey === normalizedKey;
    }) || null;
  }
  return null;
}

function normalizeTrackedProjectSelection(raw = {}) {
  if (!raw || typeof raw !== 'object') return null;

  const keyInput = raw.projectKey || raw.key || raw.project || raw.projectFolderName || raw.projectName || raw.name;
  const normalizedKey = keyInput ? getProjectCacheKey(keyInput) : null;
  const context = resolveTrackedProjectContext(normalizedKey || raw.projectName || raw.projectFolderName || raw.name);
  const projectKey = normalizedKey || (context ? context.projectKey : null);

  if (!projectKey) {
    return null;
  }

  const nowIso = new Date().toISOString();
  const projectName = String(raw.projectName || raw.name || (context && (context.displayName || context.name || context.projectFolderName)) || '').trim() || projectKey;
  const projectNumber = raw.projectNumber || (context && context.projectNumber) || null;
  const siteUrl = raw.siteUrl || raw.site || (context && (context.site || context.siteUrl || context.siteURL)) || null;
  const stateName = raw.stateName || raw.state || (context && (context.stateName || context.state)) || null;
  const driveId = raw.driveId || (context && context.driveId) || null;
  const stateFolderId = raw.stateFolderId || (context && context.stateFolderId) || null;
  const stateFolderName = raw.stateFolderName || (context && context.stateFolderName) || null;
  const workbookId = raw.workbookId || raw.fileId || (context && (context.fileId || context.workbookId)) || null;
  const workbookPath = raw.workbookPath || (context && (context.workbookPath || context.filePath)) || null;

  return {
    projectKey,
    projectName,
    projectNumber,
    siteUrl,
    stateName,
    stateFolderId,
    stateFolderName,
    driveId,
    workbookId,
    workbookPath,
    createdAt: raw.createdAt || (context && context.createdAt) || nowIso,
    updatedAt: nowIso
  };
}

async function setTrackedProjectsConfig(projects = []) {
  const existingMap = new Map(Array.isArray(trackedProjectsConfig) ? trackedProjectsConfig.map(entry => [entry.projectKey, entry]) : []);
  const normalizedList = [];
  const seen = new Set();

  if (Array.isArray(projects)) {
    projects.forEach(item => {
      const normalized = normalizeTrackedProjectSelection(item);
      if (!normalized || seen.has(normalized.projectKey)) {
        return;
      }
      const previous = existingMap.get(normalized.projectKey);
      if (previous) {
        normalized.createdAt = previous.createdAt || normalized.createdAt;
      }
      normalizedList.push(normalized);
      seen.add(normalized.projectKey);
    });
  }

  trackedProjectsConfig = normalizedList;
  metadataCache = metadataCache || {};
  metadataCache.trackedProjects = trackedProjectsConfig;
  try {
    await DB.saveMetadata(metadataCache);
  } catch (error) {
    console.error('[Metadata] Failed to persist tracked projects configuration:', error);
  }
  return getTrackedProjectsConfigSnapshot();
}

function getTrackedProjectKeys() {
  return Array.isArray(trackedProjectsConfig)
    ? trackedProjectsConfig.map(entry => entry.projectKey)
    : [];
}

function getProjectCacheKey(projectName) {
  return compactProjectName(projectName || '');
}

function normalizeProjectRecord(project) {
  if (!project || typeof project !== 'object') return project;

  if (project.jobName && typeof project.jobName === 'string') {
    const jobNameTrimmed = project.jobName.trim();
    if (jobNameTrimmed) {
      if (!project.displayName || project.displayName === project.name || project.displayName === project.projectFolderName) {
        project.displayName = jobNameTrimmed;
      }
    }
  }

  if (!project.displayName) {
    project.displayName = project.projectFolderName || project.name || project.projectName || project.fileName || '';
    if (typeof project.displayName === 'string') {
      project.displayName = project.displayName.trim();
    }
  } else if (typeof project.displayName === 'string') {
    project.displayName = project.displayName.trim();
  }

  if (!project.projectFolderName) {
    if (project.projectName) {
      project.projectFolderName = project.projectName;
    } else if (project.name) {
      project.projectFolderName = project.name;
    } else if (project.displayName) {
      project.projectFolderName = project.displayName;
    }
  }

  if (project.folderName && !project.projectFolderName) {
    project.projectFolderName = project.folderName;
  }

  if (!project.projectFolderName) {
    project.projectFolderName = project.displayName || project.name || '';
  }

  if (!project.projectKey && project.projectFolderName) {
    project.projectKey = getProjectCacheKey(project.projectFolderName);
  } else if (!project.projectKey && project.name) {
    project.projectKey = getProjectCacheKey(project.name);
  }

  if (!project.projectNumber && project.displayName) {
    project.projectNumber = extractProjectNumber(project.displayName);
  }

  return project;
}

function getProjectSortName(project) {
  if (!project || typeof project !== 'object') return '';
  const candidates = [
    project.displayName,
    project.name,
    project.projectFolderName,
    project.projectName,
    project.fileName
  ];
  for (const candidate of candidates) {
    if (!candidate) continue;
    const value = String(candidate).trim();
    if (value) return value;
  }
  return '';
}

function getProjectCodeValueFromProject(project) {
  if (!project || typeof project !== 'object') return null;
  const candidates = [
    project.projectNumber,
    project.projectFolderName,
    project.displayName,
    project.name,
    project.projectName,
    project.fileName
  ];
  for (const candidate of candidates) {
    if (!candidate) continue;
    const value = String(candidate).trim();
    if (!value) continue;
    const match = value.match(PROJECT_CODE_PATTERN);
    if (match) {
      return parseInt(match[1], 10) * 1000 + parseInt(match[2], 10);
    }
  }
  return null;
}

function getProjectRecencyTimestamp(project) {
  if (!project || typeof project !== 'object') return 0;
  const candidates = [
    project.lastRefreshed,
    project.lastDiscovered,
    project.lastUpdated,
    project.updatedAt,
    project.createdAt
  ];
  let latest = 0;
  for (const candidate of candidates) {
    if (candidate === null || candidate === undefined) continue;
    let value = 0;
    if (typeof candidate === 'number') {
      value = candidate;
    } else {
      const parsed = Date.parse(candidate);
      if (!Number.isNaN(parsed)) value = parsed;
    }
    if (value > latest) latest = value;
  }
  return latest;
}

function compareCachedProjectsDesc(a, b) {
  if (!a && !b) return 0;
  if (!a) return 1;
  if (!b) return -1;
  const timeA = getProjectRecencyTimestamp(a);
  const timeB = getProjectRecencyTimestamp(b);
  if (timeA && timeB && timeA !== timeB) return timeB - timeA;
  if (timeA && !timeB) return -1;
  if (!timeA && timeB) return 1;
  const codeA = getProjectCodeValueFromProject(a);
  const codeB = getProjectCodeValueFromProject(b);
  if (codeA !== null && codeB !== null && codeA !== codeB) return codeB - codeA;
  if (codeA !== null && codeB === null) return -1;
  if (codeA === null && codeB !== null) return 1;
  const nameA = getProjectSortName(a);
  const nameB = getProjectSortName(b);
  return nameA.localeCompare(nameB, undefined, { sensitivity: 'base' });
}

function buildProjectKeySet(projects) {
  const keySet = new Set();
  (projects || []).forEach(project => {
    if (!project || typeof project !== 'object') return;
    if (typeof project.projectKey === 'string' && project.projectKey) {
      keySet.add(project.projectKey);
    }
    const candidates = [
      project.projectFolderName,
      project.projectFolder,
      project.displayName,
      project.name,
      project.projectName,
      project.fileName,
      project.projectNumber
    ];
    candidates.forEach(candidate => {
      if (!candidate) return;
      const value = String(candidate).trim();
      if (!value) return;
      const key = getProjectCacheKey(value);
      if (key) keySet.add(key);
    });
  });
  return keySet;
}

function enforceProjectRetentionLimit(limit = PROJECT_CACHE_LIMIT) {
  if (!Array.isArray(cachedProjects) || cachedProjects.length <= limit) return;

  const sorted = cachedProjects.filter(Boolean).sort(compareCachedProjectsDesc);
  const retained = sorted.slice(0, limit);
  const retainKeys = buildProjectKeySet(retained);
  const removedCount = cachedProjects.length - retained.length;
  cachedProjects = retained;

  if (Array.isArray(cachedContacts) && cachedContacts.length) {
    cachedContacts = cachedContacts.filter(contact => {
      const candidates = [
        contact.projectKey,
        contact.projectFolderName ? getProjectCacheKey(contact.projectFolderName) : null,
        contact.project ? getProjectCacheKey(contact.project) : null,
        contact.projectName ? getProjectCacheKey(contact.projectName) : null,
        contact.projectFolder ? getProjectCacheKey(contact.projectFolder) : null,
        contact.projectDisplayName ? getProjectCacheKey(contact.projectDisplayName) : null
      ];
      return candidates.some(key => key && retainKeys.has(key));
    });
  }

  console.log(`[Cache] Retained ${retained.length} most recent projects${removedCount > 0 ? `, trimmed ${removedCount}` : ''}.`);
}


function initializeStaticProjectCache() {
  if (initializeStaticProjectCache._initialized) {
    return;
  }
  initializeStaticProjectCache._initialized = true;

  if (!STATIC_CACHE_SOURCE) {
    console.log('[StaticCache] No bundled static cache found (static_project_cache.js not provided)');
    return;
  }

  try {
    const payload = STATIC_CACHE_SOURCE || {};

    staticProjectKeySet.clear();
    staticActiveProjectKeySet.clear();


    const projects = Array.isArray(payload.projects) ? payload.projects : [];
    staticProjectsCache = projects.map(project => {
      const clone = { ...project };
      normalizeProjectRecord(clone);
      clone.cacheSource = clone.cacheSource || 'static';
      const normalizedKey = getProjectCacheKey(clone.projectKey || clone.projectFolderName || clone.name || clone.displayName || '');
      if (normalizedKey) {
        clone.projectKey = normalizedKey;
        staticProjectKeySet.add(normalizedKey);
      }
      return clone;
    }).filter(Boolean);

    const contacts = Array.isArray(payload.contacts) ? payload.contacts : [];
    staticContactsCache = contacts.map(contact => normalizeStaticContact(contact)).filter(Boolean);

    if (payload.zipCache && typeof payload.zipCache === 'object') {
      staticZipLocationCache = { ...payload.zipCache };
    } else {
      staticZipLocationCache = {};
    }

    if (Array.isArray(payload.activeProjectKeys)) {
      payload.activeProjectKeys.forEach(key => {
        const normalizedKey = getProjectCacheKey(key || '');
        if (normalizedKey) {
          staticActiveProjectKeySet.add(normalizedKey);
        }
      });
    }

    console.log(`[StaticCache] Bundled cache loaded: ${staticProjectsCache.length} projects, ${staticContactsCache.length} contacts, ${staticActiveProjectKeySet.size} active keys`);
  } catch (error) {
    console.error('[StaticCache] Failed to initialize bundled cache:', error);
    staticProjectsCache = [];
    staticContactsCache = [];
    staticZipLocationCache = {};
    staticProjectKeySet.clear();
  }
}

function normalizeStaticContact(contact) {
  if (!contact || typeof contact !== 'object') {
    return null;
  }

  const normalized = { ...contact };

  if (!normalized.project && (normalized.projectName || normalized.projectFolderName || normalized.projectDisplayName)) {
    normalized.project = normalized.projectName || normalized.projectFolderName || normalized.projectDisplayName;
  }

  const normalizedProjectKey = getProjectCacheKey(
    normalized.projectKey
    || normalized.projectFolderName
    || normalized.project
    || normalized.projectName
    || ''
  );

  if (normalizedProjectKey) {
    normalized.projectKey = normalizedProjectKey;
  }

  if (normalized.email) {
    normalized.email = String(normalized.email).trim();
    normalized.emailNormalized = normalized.email.toLowerCase();
  }

  if (!normalized.company && normalized.companyName) {
    normalized.company = normalized.companyName;
  }

  if (normalized.company) {
    normalized.company = String(normalized.company).trim();
  }

  if (!normalized.division && normalized.divisionName) {
    normalized.division = normalized.divisionName;
  }

  if (!normalized.sheetName && normalized.division) {
    normalized.sheetName = normalized.division;
  }

  normalized.cacheSource = 'static';
  return normalized;
}

function buildContactDedupKey(contact) {
  const projectKey = getProjectCacheKey(contact?.projectKey || contact?.project || contact?.projectName || contact?.projectFolderName || '');
  const divisionKey = String(contact?.divisionKey || contact?.division || contact?.sheetName || '').trim().toLowerCase();
  const email = String(contact?.email || contact?.emailNormalized || contact?.Email || '').trim().toLowerCase();
  const company = String(contact?.company || contact?.companyName || '').trim().toLowerCase();
  const name = String(contact?.name || contact?.contactName || '').trim().toLowerCase();
  return [projectKey, divisionKey, email, company, name].join('|');
}

function buildCombinedProjectList() {
  const combined = new Map();

  const addProject = (project, source) => {
    if (!project || typeof project !== 'object') return;
    const candidate = { ...project };
    normalizeProjectRecord(candidate);
    candidate.cacheSource = source;
    const key = getProjectCacheKey(candidate.projectKey || candidate.projectFolderName || candidate.name || candidate.displayName || '');
    if (!key) return;
    candidate.projectKey = key;

    const existing = combined.get(key);
    if (existing) {
      combined.set(key, {
        ...existing,
        ...candidate,
        cacheSource: existing.cacheSource === 'dynamic' ? existing.cacheSource : candidate.cacheSource
      });
    } else {
      combined.set(key, candidate);
    }
  };

  staticProjectsCache.forEach(project => addProject(project, 'static'));
  (cachedProjects || []).forEach(project => addProject(project, 'dynamic'));

  return Array.from(combined.values());
}

function buildCombinedContactList() {
  const combined = new Map();

  const addContact = (contact, source) => {
    if (!contact || typeof contact !== 'object') return;
    const candidate = { ...contact };
    if (!candidate.projectKey) {
      candidate.projectKey = getProjectCacheKey(candidate.project || candidate.projectName || candidate.projectFolderName || '');
    }
    if (candidate.email) {
      candidate.email = String(candidate.email).trim();
      candidate.emailNormalized = candidate.email.toLowerCase();
    }
    candidate.cacheSource = source;
    const dedupKey = buildContactDedupKey(candidate);
    if (!dedupKey || dedupKey === '||||') return;

    const existing = combined.get(dedupKey);
    if (existing) {
      combined.set(dedupKey, {
        ...existing,
        ...candidate,
        cacheSource: candidate.cacheSource === 'dynamic' ? 'dynamic' : existing.cacheSource
      });
    } else {
      combined.set(dedupKey, candidate);
    }
  };

  staticContactsCache.forEach(contact => addContact(contact, 'static'));
  (cachedContacts || []).forEach(contact => addContact(contact, 'dynamic'));

  return Array.from(combined.values());
}

function buildBundledCacheSnapshot() {
  const projects = buildCombinedProjectList();
  const contacts = buildCombinedContactList();
  const zipCache = { ...staticZipLocationCache, ...(zipLocationCache || {}) };

  return {
    projects,
    contacts,
    zipCache,
    staticProjectsCount: staticProjectsCache.length,
    dynamicProjectsCount: (cachedProjects || []).length,
    staticContactsCount: staticContactsCache.length,
    dynamicContactsCount: (cachedContacts || []).length,
    activeProjectKeys: Array.from(staticActiveProjectKeySet.values())
  };
}
function extractProjectNumber(value) {
  if (!value) return '';
  const match = String(value).trim().match(/^(\d{2})[-\s]?(\d{2,3})/);
  if (!match) return '';
  return `${match[1]}-${match[2]}`;
}

function deriveProjectDisplayName(project = {}, jobInfo = {}) {
  const jobName = typeof jobInfo.jobName === 'string' ? jobInfo.jobName.trim() : '';
  if (jobName) return jobName;

  const candidates = [
    project.displayName,
    project.projectName,
    project.projectFolderName,
    project.name,
    project.fileName
  ];

  for (const candidate of candidates) {
    if (candidate && String(candidate).trim()) {
      return String(candidate).trim();
    }
  }

  return '';
}

function mergeJobInfo(existing = {}, incoming = {}) {
  const result = { ...(existing || {}) };
  const source = incoming || {};
  const fields = ['jobName', 'address', 'squareFootage', 'bidDate', 'city', 'state', 'zip'];
  for (const field of fields) {
    const incomingValue = source[field];
    if (incomingValue === undefined || incomingValue === null || incomingValue === '') continue;
    if (!result[field] || String(incomingValue).length >= String(result[field]).length) {
      result[field] = incomingValue;
    }
  }
  return result;
}

function dedupeContacts(contacts = []) {
  const seen = new Set();
  const result = [];
  contacts.forEach(contact => {
    if (!contact) return;
    const divisionKey = contact.divisionKey || contact.division || '';
    const companyKey = normalizeMatchValue(contact.company || contact.companyNormalized || '');
    const emailKey = String(contact.email || '').trim().toLowerCase();
    const phoneKey = normalizePhoneDigits(contact.phone || contact.phoneDigits);
    const uniqueKey = `${divisionKey}|${companyKey}|${emailKey}|${phoneKey}`;
    if (seen.has(uniqueKey)) return;
    seen.add(uniqueKey);
    result.push(contact);
  });
  return result;
}

function getProjectFolderKey(project = {}) {
  const folderName = project.projectFolderName || project.folderName || project.name || project.displayName || '';
  return getProjectCacheKey(folderName);
}
function getOrCreateProjectRuntime(projectName) {
  const key = getProjectCacheKey(projectName);
  if (!projectRuntimeCache.has(key)) {
    projectRuntimeCache.set(key, { key, name: projectName || '' });
  }
  const runtime = projectRuntimeCache.get(key);
  if (projectName && !runtime.name) {
    runtime.name = projectName;
  }
  return runtime;
}

// State abbreviation to full name mapping
const STATE_ABBREVIATIONS = {
  'AL': 'Alabama', 'AK': 'Alaska', 'AZ': 'Arizona', 'AR': 'Arkansas',
  'CA': 'California', 'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware',
  'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii', 'ID': 'Idaho',
  'IL': 'Illinois', 'IN': 'Indiana', 'IA': 'Iowa', 'KS': 'Kansas',
  'KY': 'Kentucky', 'LA': 'Louisiana', 'ME': 'Maine', 'MD': 'Maryland',
  'MA': 'Massachusetts', 'MI': 'Michigan', 'MN': 'Minnesota', 'MS': 'Mississippi',
  'MO': 'Missouri', 'MT': 'Montana', 'NE': 'Nebraska', 'NV': 'Nevada',
  'NH': 'New Hampshire', 'NJ': 'New Jersey', 'NM': 'New Mexico', 'NY': 'New York',
  'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio', 'OK': 'Oklahoma',
  'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island', 'SC': 'South Carolina',
  'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas', 'UT': 'Utah',
  'VT': 'Vermont', 'VA': 'Virginia', 'WA': 'Washington', 'WV': 'West Virginia',
  'WI': 'Wisconsin', 'WY': 'Wyoming', 'DC': 'District of Columbia'
};

// Helper to get full state name
function getFullStateName(stateInput) {
  if (!stateInput) return '';
  
  const trimmed = stateInput.trim().toUpperCase();
  
  if (STATE_ABBREVIATIONS[trimmed]) {
    return STATE_ABBREVIATIONS[trimmed];
  }
  
  for (const [abbr, fullName] of Object.entries(STATE_ABBREVIATIONS)) {
    if (fullName.toUpperCase() === trimmed) {
      return fullName;
    }
  }
  
  return stateInput;
}

// ============ INDEXEDDB CACHE MANAGEMENT ============

async function saveCacheToIndexedDB() {
  try {
    console.log('[IndexedDB] Saving cache...');

    metadataCache = metadataCache || {};
    metadataCache.lastScanTime = lastScanTime;
    metadataCache.lastBaselineScan = lastBaselineScan;
    metadataCache.lastIncrementalScan = lastIncrementalScan;
    metadataCache.totalProjects = cachedProjects.length;
    metadataCache.totalContacts = cachedContacts.length;
    metadataCache.zipCache = zipLocationCache;
    metadataCache.trackedProjects = trackedProjectsConfig;

    await DB.saveMetadata(metadataCache);

    await DB.saveProjects(cachedProjects);
    await DB.saveContacts(cachedContacts);

    console.log('[IndexedDB] Cache saved successfully');
    console.log(`[IndexedDB] Saved ${cachedProjects.length} projects, ${cachedContacts.length} contacts`);

    return { success: true };
  } catch (error) {
    console.error('[IndexedDB] Error saving cache:', error);
    return { success: false, error: error.message };
  }
}

async function loadCacheFromIndexedDB() {
  try {
    console.log('[IndexedDB] Loading cache...');

    metadataCache = await DB.getMetadata() || {};
    lastScanTime = metadataCache.lastScanTime;
    lastBaselineScan = metadataCache.lastBaselineScan || null;
    lastIncrementalScan = metadataCache.lastIncrementalScan || null;
    zipLocationCache = {
      ...staticZipLocationCache,
      ...(metadataCache.zipCache || {})
    };
    trackedProjectsConfig = Array.isArray(metadataCache.trackedProjects) ? metadataCache.trackedProjects : [];

    if (lastScanTime) {
      console.log('[IndexedDB] Loaded metadata, last scan:', lastScanTime);
    }

    cachedProjects = (await DB.getAllProjects()).map(normalizeProjectRecord);
    console.log('[IndexedDB] Loaded', cachedProjects.length, 'projects');

    cachedContacts = await DB.getAllContacts();
    enforceProjectRetentionLimit();
    console.log('[IndexedDB] Loaded', cachedContacts.length, 'contacts');

    return {
      success: true,
      projectsLoaded: cachedProjects.length,
      contactsLoaded: cachedContacts.length
    };
  } catch (error) {
    console.error('[IndexedDB] Error loading cache:', error);
    return { success: false, error: error.message };
  }
}

async function clearCache() {
  try {
    console.log('[IndexedDB] Clearing cache...');
    
    await DB.clearAll();
    
    cachedProjects = [];
    cachedContacts = [];
    lastScanTime = null;
    lastBaselineScan = null;
    lastIncrementalScan = null;
    
    console.log('[IndexedDB] Cache cleared successfully');
    return { success: true };
  } catch (error) {
    console.error('[IndexedDB] Error clearing cache:', error);
    return { success: false, error: error.message };
  }
}

async function exportCacheToFile() {
  try {
    console.log('[Export] Creating backup file...');
    
    const backupData = {
      version: '1.0',
      timestamp: new Date().toISOString(),
      projects: cachedProjects,
      contacts: cachedContacts,
      lastScanTime: lastScanTime,
      totalProjects: cachedProjects.length,
      totalContacts: cachedContacts.length
    };
    
    const jsonString = JSON.stringify(backupData, null, 2);
    const dataUrl = 'data:application/json;charset=utf-8,' + encodeURIComponent(jsonString);
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 10);
    const filename = `AB_Cache_Backup_${timestamp}.json`;
    
    const downloadId = await chrome.downloads.download({
      url: dataUrl,
      filename: filename,
      saveAs: true
    });
    
    console.log('[Export] Download started with ID:', downloadId);
    console.log('[Export] Cache backup saved:', filename);
    
    return { success: true, filename };
  } catch (error) {
    console.error('[Export] Error saving cache backup:', error);
    return { success: false, error: error.message };
  }
}

async function importCacheFromFile(fileContent) {
  try {
    console.log('[Import] Loading cache from backup file...');
    
    const backupData = JSON.parse(fileContent);
    
    if (!backupData.contacts || !Array.isArray(backupData.contacts)) {
      throw new Error('Invalid backup file format');
    }
    
    await clearCache();
    
    cachedProjects = (backupData.projects || []).map(normalizeProjectRecord);
    cachedContacts = backupData.contacts || [];

    enforceProjectRetentionLimit();
    lastScanTime = backupData.lastScanTime || null;
    lastBaselineScan = backupData.lastBaselineScan || null;
    lastIncrementalScan = backupData.lastIncrementalScan || null;
    
    await saveCacheToIndexedDB();

    console.log('[Import] Cache loaded from backup file');
    console.log('[Import] Restored:', cachedContacts.length, 'contacts,', cachedProjects.length, 'projects');
    
    return {
      success: true,
      contactsLoaded: cachedContacts.length,
      projectsLoaded: cachedProjects.length
    };
  } catch (error) {
    console.error('[Import] Error loading cache from file:', error);
    return { success: false, error: error.message };
  }
}

// ============ END INDEXEDDB CACHE MANAGEMENT ============

// Helper functions for PKCE authentication
function generateRandomString(length) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

async function generateCodeChallenge() {
  const codeVerifier = generateRandomString(128);
  const encoder = new TextEncoder();
  const data = encoder.encode(codeVerifier);
  const hash = await crypto.subtle.digest('SHA-256', data);
  const codeChallenge = btoa(String.fromCharCode(...new Uint8Array(hash)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=/g, '');
  return { codeVerifier, codeChallenge };
}

async function authenticateWithMicrosoft() {
  try {
    console.log('Starting Microsoft authentication...');
    const redirectUri = chrome.identity.getRedirectURL();
    console.log('Chrome redirect URI:', redirectUri);
    
    const stored = await chrome.storage.local.get(['accessToken', 'tokenExpiry']);
    if (stored.accessToken && stored.tokenExpiry && stored.tokenExpiry > Date.now()) {
      accessToken = stored.accessToken;
      tokenExpiry = stored.tokenExpiry;
      console.log('Using stored valid token');
      return { success: true, token: accessToken };
    }

    const pkce = await generateCodeChallenge();
    await chrome.storage.local.set({ codeVerifier: pkce.codeVerifier });

    const authUrl = new URL(`https://login.microsoftonline.com/${AZURE_CONFIG.tenantId}/oauth2/v2.0/authorize`);
    authUrl.searchParams.set('client_id', AZURE_CONFIG.clientId);
    authUrl.searchParams.set('response_type', 'code');
    authUrl.searchParams.set('redirect_uri', redirectUri);
    authUrl.searchParams.set('scope', AZURE_CONFIG.scopes.join(' '));
    authUrl.searchParams.set('code_challenge', pkce.codeChallenge);
    authUrl.searchParams.set('code_challenge_method', 'S256');
    authUrl.searchParams.set('state', generateRandomString(32));
    authUrl.searchParams.set('prompt', 'select_account');

    const responseUrl = await chrome.identity.launchWebAuthFlow({
      url: authUrl.toString(),
      interactive: true
    });

    const urlParams = new URL(responseUrl).searchParams;
    const authCode = urlParams.get('code');
    const error = urlParams.get('error');

    if (error) {
      throw new Error(`Auth error: ${error} - ${urlParams.get('error_description')}`);
    }
    if (!authCode) {
      throw new Error('No authorization code received');
    }

    const tokenResponse = await exchangeCodeForToken(authCode, pkce.codeVerifier, redirectUri);
    
    if (tokenResponse.access_token) {
      accessToken = tokenResponse.access_token;
      tokenExpiry = Date.now() + (tokenResponse.expires_in * 1000);
      
      await chrome.storage.local.set({
        accessToken: accessToken,
        tokenExpiry: tokenExpiry,
        refreshToken: tokenResponse.refresh_token
      });

      console.log('Authentication successful');
      return { success: true, token: accessToken };
    } else {
      throw new Error('No access token received');
    }
  } catch (error) {
    console.error('Authentication failed:', error);
    await chrome.storage.local.remove(['accessToken', 'tokenExpiry', 'refreshToken']);
    return { success: false, error: error.message };
  }
}

async function exchangeCodeForToken(authCode, codeVerifier, redirectUri) {
  console.log('[Auth] Starting token exchange');
  const tokenUrl = `https://login.microsoftonline.com/${AZURE_CONFIG.tenantId}/oauth2/v2.0/token`;
  console.log('[Auth] Token URL:', tokenUrl);
  
  const body = new URLSearchParams({
    client_id: AZURE_CONFIG.clientId,
    scope: AZURE_CONFIG.scopes.join(' '),
    code: authCode,
    redirect_uri: redirectUri,
    grant_type: 'authorization_code',
    code_verifier: codeVerifier
  });

  console.log('[Auth] Posting to token endpoint...');
  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  console.log('[Auth] Token response status:', response.status);
  if (!response.ok) {
    const errorText = await response.text();
    console.error('[Auth] Token error:', errorText);
    throw new Error(`Token exchange failed: ${response.status} - ${errorText}`);
  }

  console.log('[Auth] Parsing token response...');
  const result = await response.json();
  console.log('[Auth] Token received successfully');
  return result;
}

async function callGraphAPI(endpoint, options = {}) {
  self.callGraphAPI = callGraphAPI;
  if (!accessToken || !tokenExpiry || tokenExpiry < Date.now()) {
    const authResult = await authenticateWithMicrosoft();
    if (!authResult.success) {
      throw new Error('Authentication required');
    }
  }

  const url = endpoint.startsWith('https://') 
    ? endpoint 
    : `https://graph.microsoft.com/v1.0${endpoint}`;
  
  const response = await fetch(url, {
    method: options.method || 'GET',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });

  if (response.status === 401) {
    accessToken = null;
    tokenExpiry = null;
    await chrome.storage.local.remove(['accessToken', 'tokenExpiry']);
    
    const authResult = await authenticateWithMicrosoft();
    if (!authResult.success) {
      throw new Error('Re-authentication failed');
    }
    
    return callGraphAPI(endpoint, options);
  }

  if (!response.ok) {
    const errorText = await response.text();
    const error = new Error(`API call failed: ${response.status} - ${errorText}`);
    error.status = response.status;
    error.body = errorText;
    error.endpoint = url;
    throw error;
  }

  const contentType = response.headers.get('content-type');
  if (contentType && contentType.includes('application/json')) {
    return await response.json();
  } else {
    return await response.arrayBuffer();
  }

}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function executeWithRetry(operation, options = {}) {
  const { retries = 1, delayMs = 1200, retryStatuses = [423, 429], label = 'operation' } = options;
  let attempt = 0;
  while (true) {
    try {
      return await operation();
    } catch (error) {
      const status = error && typeof error === 'object' ? error.status : undefined;
      if (attempt >= retries || !status || !retryStatuses.includes(status)) {
        throw error;
      }
      attempt += 1;
      console.warn(`[Retry] ${label} failed with status ${status}. Retrying in ${delayMs}ms...`);
      await delay(delayMs);
    }
  }
}

function getWorkbookRequestHeaders(additional = {}) {
  const headers = { ...additional };
  if (
    currentWorkbookSession &&
    currentWorkbookSession.id &&
    currentProjectFile &&
    currentWorkbookSession.itemId === currentProjectFile.id
  ) {
    headers['workbook-session-id'] = currentWorkbookSession.id;
  }
  return headers;
}

function normalizeEmail(value) {
  if (!value) return '';
  const trimmed = String(value).trim().toLowerCase();
  return trimmed.includes('@') ? trimmed : '';
}

function resolveSharePointUserName(value, details = {}) {
  const email = normalizeEmail(value);
  const detailEmail = normalizeEmail(details.email);
  const displayName = (details.displayName || details.name || '').trim();

  if (displayName) {
    return displayName;
  }

  const effectiveEmail = email || detailEmail;
  if (effectiveEmail && SHAREPOINT_USER_OVERRIDES[effectiveEmail]) {
    return SHAREPOINT_USER_OVERRIDES[effectiveEmail];
  }

  if (details.displayName) {
    return String(details.displayName).trim();
  }
  if (details.name) {
    return String(details.name).trim();
  }

  if (effectiveEmail) {
    const override = SHAREPOINT_USER_OVERRIDES[effectiveEmail];
    if (override) return override;
    return effectiveEmail;
  }

  return value || 'Unknown';
}

async function createWorkbookSession(driveId, itemId, options = {}) {
  if (!driveId || !itemId) {
    throw new Error('Missing driveId or itemId for workbook session');
  }
  const { persistChanges = false } = options;
  const result = await executeWithRetry(
    () => callGraphAPI(
      `/drives/${driveId}/items/${itemId}/workbook/createSession`,
      { method: 'POST', body: { persistChanges } }
    ),
    { retries: 1, delayMs: 1500, retryStatuses: [423], label: 'createSession' }
  );
  if (!result || !result.id) {
    throw new Error('Failed to create workbook session');
  }
  return result.id;
}

async function closeWorkbookSession(driveId, itemId, sessionId) {
  if (!driveId || !itemId || !sessionId) return;
  try {
    await callGraphAPI(`/drives/${driveId}/items/${itemId}/workbook/closeSession`, {
      method: 'POST',
      headers: { 'workbook-session-id': sessionId }
    });
  } catch (error) {
    if (error && (error.status === 404 || error.status === 410)) {
      return;
    }
    console.warn('[Excel API] Failed to close workbook session:', error);
  }
}

async function getSiteByUrl(siteUrl) {
  const match = siteUrl.match(/([^:]+):(.+)/);
  if (!match) throw new Error('Invalid site URL format');
  
  const hostname = match[1];
  const sitePath = match[2];
  
  return await callGraphAPI(`/sites/${hostname}:${sitePath}`);
}

async function getDrives(siteId) {
  return await callGraphAPI(`/sites/${siteId}/drives`);
}

async function getFolders(driveId, parentPath = '') {
  let endpoint = `/drives/${driveId}/root`;
  if (parentPath) {
    endpoint = `/drives/${driveId}/items/${parentPath}`;
  }
  endpoint += '/children?$filter=folder ne null';
  
  const result = await callGraphAPI(endpoint);
  return result.value || [];
}

async function getFiles(driveId, folderId) {
  const endpoint = `/drives/${driveId}/items/${folderId}/children`;
  const result = await callGraphAPI(endpoint);
  
  return (result.value || []).filter(file => 
    file.name && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.xlsm'))
  );
}

function parseExcel(arrayBuffer) {
  try {
    if (typeof XLSX === 'undefined' || !XLSX.read) {
      console.error('XLSX library not loaded or incomplete');
      return { sheetNames: [], sheets: {} };
    }
    
    const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
    const result = {
      sheetNames: workbook.SheetNames,
      sheets: {}
    };
    
    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      result.sheets[sheetName] = sheet;
    });
    
    return result;
  } catch (error) {
    console.error('Error parsing Excel:', error);
    return { sheetNames: [], sheets: {} };
  }
}

function getCellValue(sheet, cellRef) {
  if (!sheet || !cellRef) return '';
  const cell = sheet[cellRef];
  if (!cell) return '';
  return cell.v || '';
}

function parseAddress(address) {
  if (!address) return { city: '', state: '', zip: '' };
  
  address = address.replace(/\s+/g, ' ').trim();
  
  const patterns = [
    /([^,]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/i,
    /([^,]+)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$/i,
  ];
  
  for (const pattern of patterns) {
    const match = address.match(pattern);
    if (match) {
      return {
        city: match[1].trim(),
        state: getFullStateName(match[2]),
        zip: match[3]
      };
    }
  }
  
  return { city: '', state: '', zip: '' };
}

async function scanSingleState(siteUrl, stateName, progressCallback = null, trackedProjectsToScan = null) {
  const newProjects = [];
  const newContacts = [];
  const processedFiles = new Set();
  const existingProjectEmailCache = new Map();
  const auditEntriesBuffer = [];
  const trackedProjectsSnapshot = getTrackedProjectsConfigSnapshot();
  const trackedProjectMap = new Map(trackedProjectsSnapshot.map(config => [config.projectKey, config]));
  const trackedProjectKeySet = new Set(trackedProjectMap.keys());
  const trackedAuditBuffer = [];
  
  // If specific projects are provided (from rescan), create a map for quick lookup
  const trackedProjectsForThisScan = trackedProjectsToScan || trackedProjectsSnapshot;
  const projectNamesToScan = new Set();
  trackedProjectsForThisScan.forEach(proj => {
    if (proj && proj.projectFolderName) {
      projectNamesToScan.add(String(proj.projectFolderName).toLowerCase().trim());
    }
  });
  
  const getExistingEmailsForProject = (key) => {
    if (!key) return new Set();
    if (existingProjectEmailCache.has(key)) {
      return existingProjectEmailCache.get(key);
    }
    const normalizedKey = key;
    const emailSet = new Set();
    (cachedContacts || []).forEach(contact => {
      if (!contact) return;
      const contactProjectKey = contact.projectKey || contact.projectFolderName || contact.project || '';
      if (contactProjectKey === normalizedKey) {
        const email = (contact.email || '').trim().toLowerCase();
        if (email) emailSet.add(email);
      }
    });
    existingProjectEmailCache.set(normalizedKey, emailSet);
    return emailSet;
  };
  const projectAggregates = new Map();

  if (!XLSX_LOADED || typeof XLSX === 'undefined' || !XLSX.read || !XLSX.utils) {
    console.error('XLSX library check failed');
    return {
      success: false,
      error: 'Excel parsing library not loaded. Make sure xlsx.full.min.js is in your extension directory and reload the extension.'
    };
  }

  const updateProgress = (current, total, status = 'Processing') => {
    if (progressCallback) {
      progressCallback({
        current,
        total,
        percentage: total > 0 ? Math.round((current / total) * 100) : 0,
        status
      });
    }
  };

  const divisionSheets = FALLBACK_DIVISION_ORDER;
  const excludedSheetsSet = new Set(EXCLUDED_SHEETS.map(name => String(name || '').toLowerCase()));

  try {
    console.log(`\n========== Scanning state: ${stateName} from ${siteUrl} ==========`);

    const site = await getSiteByUrl(siteUrl);
    if (!site) throw new Error(`Could not access site: ${siteUrl}`);

    const drives = await getDrives(site.id);
    const sharedDocs = drives.value?.find(d =>
      d.name === 'Documents' ||
      d.name === 'Shared Documents' ||
      (d.name || '').includes('Shared')
    );

    if (!sharedDocs) throw new Error(`No shared documents found in ${siteUrl}`);

    const topFolders = await getFolders(sharedDocs.id);

    const stateFolder = topFolders.find(folder => {
      if (!folder?.name) return false;
      const folderName = folder.name.toLowerCase();
      const searchName = stateName.toLowerCase();

      if (folderName === searchName) return true;
      if (folderName === `${searchName} projects`) return true;
      if (folderName === `${searchName} project`) return true;
      if (folderName.startsWith(`${searchName} `)) return true;

      return false;
    });

    if (!stateFolder) throw new Error(`State folder "${stateName}" not found`);

    console.log(`Found state folder: ${stateFolder.name}`);

    const projectFolders = await getFolders(sharedDocs.id, stateFolder.id);
    console.log(`Found ${projectFolders.length} project folders in state folder`);
    
    // Filter to only tracked projects if specific projects were provided
    let foldersToProcess = projectFolders;
    if (trackedProjectsToScan && projectNamesToScan.size > 0) {
      foldersToProcess = projectFolders.filter(folder => {
        if (!folder?.name) return false;
        return projectNamesToScan.has(folder.name.toLowerCase().trim());
      });
      console.log(`Filtered to ${foldersToProcess.length} tracked project folder(s) to scan`);
    }
    
    let totalFiles = 0;
    const projectFileMap = new Map();

    updateProgress(0, 1, 'Counting files...');

    for (const projectFolder of foldersToProcess) {
      if (!projectFolder?.name || projectFolder.name.startsWith('.')) continue;
      const excelFiles = await getFiles(sharedDocs.id, projectFolder.id);
      const validFiles = (excelFiles || [])
        .filter(file => file && !String(file.name || '').startsWith('~$') && file.size <= 15728640);
      if (validFiles.length === 0) continue;
      projectFileMap.set(projectFolder.id, { folder: projectFolder, files: validFiles });
      totalFiles += validFiles.length;
    }

    console.log(`Total files to process: ${totalFiles}`);
    updateProgress(0, totalFiles, 'Starting scan...');

    let processedCount = 0;
    const startTime = Date.now();

    for (const [folderId, { folder: projectFolder, files: excelFiles }] of projectFileMap) {
      if (!projectFolder?.name || projectFolder.name.startsWith('.')) continue;

      const projectName = projectFolder.name;
      const projectAggregateKey = `${sharedDocs.id}:${projectFolder.id}`;

      let projectRecord = newProjects.find(p => p.projectFolderId === projectFolder.id);
      if (!projectRecord) {
        projectRecord = {
          name: projectName,
          displayName: projectName,
          projectFolderName: projectName,
          projectFolderId: projectFolder.id,
          state: getFullStateName(stateName),
          site: siteUrl,
          driveId: sharedDocs.id,
          contactsByDivision: {},
          divisions: [],
          projectKey: getProjectCacheKey(projectName)
        };
        newProjects.push(projectRecord);
      }

      let projectAggregate = projectAggregates.get(projectAggregateKey);
      if (!projectAggregate) {
        projectAggregate = {
          project: projectRecord,
          contactsByDivision: {},
          divisionMap: new Map(),
          jobInfo: null,
          latestFile: null
        };
        projectAggregates.set(projectAggregateKey, projectAggregate);
      }

      console.log(`Found ${excelFiles.length} Excel files in project: ${projectName}`);

      for (const file of excelFiles) {
        const fileKey = `${file.id}-${file.name}`;
        if (processedFiles.has(fileKey)) continue;
        processedFiles.add(fileKey);

        processedCount++;
        const elapsed = Date.now() - startTime;
        const avgTimePerFile = elapsed / Math.max(processedCount, 1);
        const remainingFiles = Math.max(totalFiles - processedCount, 0);
        const estimatedTimeRemaining = Math.round((avgTimePerFile * remainingFiles) / 1000);
        const timeString = estimatedTimeRemaining > 60
          ? `${Math.floor(estimatedTimeRemaining / 60)}m ${estimatedTimeRemaining % 60}s`
          : `${estimatedTimeRemaining}s`;

        updateProgress(
          processedCount,
          totalFiles,
          `Processing ${file.name} (${processedCount}/${totalFiles}) - Est. ${timeString} remaining`
        );

        if (file.size > 15728640) {
          console.log(`Skipping large file (over 15MB): ${file.name}`);
          continue;
        }

        try {
          const content = await callGraphAPI(`/drives/${sharedDocs.id}/items/${file.id}/content`);
          const workbook = parseExcel(content);
          if (!workbook.sheetNames || workbook.sheetNames.length === 0) continue;

          const jobInfo = {
            jobName: '',
            address: '',
            squareFootage: '',
            bidDate: '',
            city: '',
            state: getFullStateName(stateName),
            zip: ''
          };

          if (workbook.sheets['Spreadsheet']) {
            const spreadsheet = workbook.sheets['Spreadsheet'];
            jobInfo.jobName = getCellValue(spreadsheet, 'C4');
            jobInfo.address = getCellValue(spreadsheet, 'C5');
            jobInfo.squareFootage = getCellValue(spreadsheet, 'C7');
            jobInfo.bidDate = getCellValue(spreadsheet, 'A10');

            const parsed = parseAddress(jobInfo.address);
            jobInfo.city = parsed.city || jobInfo.city;
            jobInfo.state = parsed.state || jobInfo.state;
            jobInfo.zip = parsed.zip;
          }

          projectAggregate.jobInfo = mergeJobInfo(projectAggregate.jobInfo, jobInfo);

          if (!projectAggregate.latestFile || new Date(file.lastModifiedDateTime || 0) > new Date(projectAggregate.latestFile.lastModifiedDateTime || 0)) {
            projectAggregate.latestFile = file;
          }

          const allSheetNames = workbook.sheetNames || [];
          const nonExcludedSheets = allSheetNames.filter(name => {
            if (!name) return false;
            return !excludedSheetsSet.has(String(name).toLowerCase());
          });

          const candidateSheets = nonExcludedSheets.filter(name => {
            const lower = String(name).toLowerCase();
            return divisionSheets.some(div => lower.includes(div.toLowerCase()));
          });

          const sheetsToScan = candidateSheets.length > 0 ? candidateSheets : nonExcludedSheets;
          const uniqueSheetsToScan = Array.from(new Set(sheetsToScan.filter(Boolean)));

          if (uniqueSheetsToScan.length === 0) continue;

          console.log(`Scanning ${uniqueSheetsToScan.length} division sheets in ${file.name}:`, uniqueSheetsToScan.join(', '));

          for (const sheetName of uniqueSheetsToScan) {
            const sheet = workbook.sheets[sheetName];
            if (!sheet) continue;

            const divisionInfo = getDivisionDisplayInfo(sheetName, projectRecord);

            if (divisionInfo.codeValue !== null) {
              if (!isDivisionCodeWithinRange(divisionInfo.codeValue)) {
                continue;
              }
            } else if (!isSheetWithinFallbackRange(sheetName)) {
              continue;
            }

            const divisionKey = sheetName.trim();
            const divisionDisplayName = divisionInfo.displayName || sheetName;
            const divisionOrder = typeof divisionInfo.codeValue === 'number'
              ? divisionInfo.codeValue
              : (divisionInfo.fallbackIndex !== -1 ? divisionInfo.fallbackIndex + 1 : Number.MAX_SAFE_INTEGER);
            const startsOnRow4 = SPECIAL_SHEETS.includes(sheetName);

            if (!projectAggregate.divisionMap.has(divisionKey)) {
              projectAggregate.divisionMap.set(divisionKey, {
                key: divisionKey,
                name: divisionDisplayName,
                displayName: divisionDisplayName,
                sheetName,
                startsOnRow4,
                codeValue: divisionInfo.codeValue,
                order: divisionOrder
              });
            }

            if (!projectAggregate.contactsByDivision[divisionKey]) {
              projectAggregate.contactsByDivision[divisionKey] = [];
            }

            for (let col = 6; col < 50; col++) {
              const colLetter = XLSX.utils.encode_col(col);
              const companyName = getCellValue(sheet, `${colLetter}3`);
              const headerLower = String(companyName || '').toLowerCase();

              if (!companyName ||
                  headerLower === 'company' ||
                  headerLower === 'company name' ||
                  String(companyName).toUpperCase() === 'ABGCS') {
                continue;
              }

              const contactName = getCellValue(sheet, `${colLetter}4`);
              const phone = getCellValue(sheet, `${colLetter}5`);
              const email = getCellValue(sheet, `${colLetter}6`);
              const bidAmount = getCellValue(sheet, `${colLetter}7`);

              const normalizedContact = {
                company: String(companyName).trim(),
                name: contactName ? String(contactName).trim() : '',
                phone: phone ? String(phone).trim() : '',
                email: email ? String(email).trim() : '',
                bidAmount: bidAmount ? String(bidAmount).trim() : '',
                division: divisionDisplayName,
                divisionKey,
                sheetName,
                project: projectName,
                state: jobInfo.state,
                city: jobInfo.city,
                zip: jobInfo.zip,
                jobName: jobInfo.jobName,
                address: jobInfo.address,
                squareFootage: jobInfo.squareFootage,
                bidDate: jobInfo.bidDate,
                source: file.name,
                companyNormalized: normalizeMatchValue(companyName),
                phoneDigits: normalizePhoneDigits(phone),
                emailDomain: extractEmailDomain(email)
              };

              if (normalizedContact.company) {
                projectAggregate.contactsByDivision[divisionKey].push(normalizedContact);
              }
            }
          }
        } catch (err) {
          console.error(`Error processing ${file.name}:`, err);
        }
      }
    }

    const scanTimestamp = new Date().toISOString();
    const previousProjectCount = cachedProjects.length;
    const previousContactCount = cachedContacts.length;
    const removalKeys = new Set();

    newContacts.length = 0;

    for (const aggregate of projectAggregates.values()) {
      const projectRecord = aggregate.project;
      if (!projectRecord) continue;

      if (aggregate.latestFile) {
        projectRecord.fileId = aggregate.latestFile.id;
        projectRecord.fileName = aggregate.latestFile.name;
        projectRecord.fileModified = aggregate.latestFile.lastModifiedDateTime || projectRecord.fileModified || null;
      }

      const jobInfo = mergeJobInfo({}, aggregate.jobInfo || {});
      jobInfo.state = jobInfo.state || projectRecord.state || getFullStateName(stateName);

      const displayName = deriveProjectDisplayName(projectRecord, jobInfo);
      const projectNumber = extractProjectNumber(displayName);
      const projectFolderName = projectRecord.projectFolderName || projectRecord.name || displayName;
      const projectKey = getProjectCacheKey(projectFolderName || displayName || projectRecord.name || '');

      projectRecord.displayName = displayName;
      projectRecord.jobName = jobInfo.jobName || displayName;
      projectRecord.projectNumber = projectNumber || projectRecord.projectNumber || null;
      projectRecord.projectFolderName = projectFolderName;
      projectRecord.projectKey = projectKey;
      projectRecord.jobInfo = jobInfo;
      projectRecord.state = jobInfo.state || projectRecord.state || getFullStateName(stateName);
      projectRecord.city = jobInfo.city || projectRecord.city || '';
      projectRecord.zip = jobInfo.zip || projectRecord.zip || '';
      projectRecord.bidDate = jobInfo.bidDate || projectRecord.bidDate || '';
      projectRecord.address = jobInfo.address || projectRecord.address || '';
      projectRecord.squareFootage = jobInfo.squareFootage || projectRecord.squareFootage || '';
      projectRecord.lastScanned = scanTimestamp;
      projectRecord.site = projectRecord.site || siteUrl;
      projectRecord.driveId = projectRecord.driveId || sharedDocs.id;

      const divisionEntries = Array.from(aggregate.divisionMap.values()).map(entry => ({
        ...entry,
        displayName: entry.displayName || entry.name || entry.sheetName || entry.key
      }));

      divisionEntries.sort((a, b) => {
        const orderA = typeof a.codeValue === 'number' ? a.codeValue : (typeof a.order === 'number' ? a.order : Number.MAX_SAFE_INTEGER);
        const orderB = typeof b.codeValue === 'number' ? b.codeValue : (typeof b.order === 'number' ? b.order : Number.MAX_SAFE_INTEGER);
        if (orderA !== orderB) return orderA - orderB;
        return (a.displayName || '').localeCompare(b.displayName || '', undefined, { sensitivity: 'base' });
      });

      const sortedDivisions = divisionEntries.map(entry => ({
        key: entry.key,
        name: entry.displayName,
        displayName: entry.displayName,
        sheetName: entry.sheetName,
        startsOnRow4: entry.startsOnRow4,
        codeValue: entry.codeValue,
        order: entry.order
      }));

      const normalizedContactsByDivision = {};
      let projectContactCount = 0;

      sortedDivisions.forEach((division, index) => {
        const contacts = Array.isArray(aggregate.contactsByDivision[division.key])
          ? aggregate.contactsByDivision[division.key]
          : [];
        const deduped = dedupeContacts(contacts).map(contact => ({
          ...contact,
          division: division.displayName,
          divisionKey: division.key,
          divisionOrder: index,
          project: displayName,
          projectDisplayName: displayName,
          projectNumber,
          projectFolderName,
          projectKey,
          state: contact.state || projectRecord.state,
          site: contact.site || projectRecord.site || siteUrl
        }));
        const existingEmails = getExistingEmailsForProject(projectKey);
        const fileInfo = aggregate.latestFile || {};
        const detectedAt = new Date().toISOString();
        const lastModifiedBy = (fileInfo && fileInfo.lastModifiedBy && fileInfo.lastModifiedBy.user && (fileInfo.lastModifiedBy.user.displayName || fileInfo.lastModifiedBy.user.email))
          || (fileInfo && fileInfo.lastModifiedBy && fileInfo.lastModifiedBy.application && fileInfo.lastModifiedBy.application.displayName)
          || projectRecord.lastModifiedBy
          || 'Unknown';
        const lastModifiedDateTime = (fileInfo && fileInfo.lastModifiedDateTime) ? fileInfo.lastModifiedDateTime : (projectRecord.fileModified || null);
        const workbookId = (fileInfo && fileInfo.id) ? fileInfo.id : (projectRecord.fileId || null);
        const workbookName = (fileInfo && fileInfo.name) ? fileInfo.name : (projectRecord.fileName || projectFolderName);
        const workbookPath = (fileInfo && fileInfo.parentReference && fileInfo.parentReference.path) ? fileInfo.parentReference.path : ((fileInfo && fileInfo.webUrl) ? fileInfo.webUrl : projectFolderName);
        const driveIdForEntry = (fileInfo && fileInfo.parentReference && fileInfo.parentReference.driveId) ? fileInfo.parentReference.driveId : (projectRecord.driveId || sharedDocs.id);
        const auditDivisionName = division.displayName || division.name || division.sheetName || division.key;
        const auditSiteUrl = projectRecord.site || siteUrl;

        const workbookKeyForAudit = (() => {

          const candidates = [
            workbookId,
            workbookPath,
            workbookName && projectKey ? `${projectKey}::${workbookName}` : null,
            projectKey
          ];

          for (const candidate of candidates) {

            if (candidate === null || candidate === undefined) continue;

            const trimmed = String(candidate).trim();

            if (trimmed) {

              return trimmed.toLowerCase();

            }

          }

          return projectKey || '';

        })();



        const auditNewEntries = [];



        deduped.forEach(contact => {

          const emailValue = (contact.email || '').trim().toLowerCase();

          if (!emailValue) return;

          if (existingEmails.has(emailValue)) return;



          existingEmails.add(emailValue);



          auditNewEntries.push({

            projectKey,

            projectName: displayName,

            projectNumber: projectNumber || null,

            workbookId,

            workbookName,

            workbookPath,

            workbookKey: workbookKeyForAudit,

            driveId: driveIdForEntry,

            siteUrl: auditSiteUrl,

            state: contact.state || projectRecord.state || jobInfo.state || null,

            divisionKey: division.key,

            divisionName: auditDivisionName,

            email: contact.email || '',

            emailNormalized: emailValue,

            company: contact.company || '',

            contactName: contact.name || '',

            phone: contact.phone || '',

            detectedAt,

            lastModifiedBy,

            lastModifiedDateTime,

            source: 'scan'

          });

        });



        if (auditNewEntries.length) {          auditEntriesBuffer.push(...auditNewEntries);          if (trackedProjectKeySet.has(projectKey)) {            auditNewEntries.forEach(entry => {              const trackedConfig = trackedProjectMap.get(projectKey) || {};              trackedAuditBuffer.push({                ...entry,                trackedAt: new Date().toISOString(),                projectState: entry.state || trackedConfig.stateName || null              });            });          }        }
        normalizedContactsByDivision[division.key] = deduped;
        projectContactCount += deduped.length;
        newContacts.push(...deduped);
      });

      projectRecord.divisions = sortedDivisions;
      projectRecord.contactsByDivision = normalizedContactsByDivision;
      projectRecord.contactCount = projectContactCount;

      const normalizedProject = normalizeProjectRecord({ ...projectRecord });
      const existingIndex = cachedProjects.findIndex(existing => getProjectCacheKey(existing.projectFolderName || existing.name || existing.displayName) === projectKey);
      if (existingIndex >= 0) {
        cachedProjects[existingIndex] = { ...cachedProjects[existingIndex], ...normalizedProject };
      } else {
        cachedProjects.push(normalizedProject);
      }

      removalKeys.add(projectKey);
      if (displayName) removalKeys.add(getProjectCacheKey(displayName));
      if (projectRecord.name) removalKeys.add(getProjectCacheKey(projectRecord.name));
    }

    if (removalKeys.size > 0) {
      cachedContacts = cachedContacts.filter(contact => {
        const keyCandidates = [
          contact.projectKey,
          contact.projectFolderName ? getProjectCacheKey(contact.projectFolderName) : null,
          contact.project ? getProjectCacheKey(contact.project) : null,
          contact.projectName ? getProjectCacheKey(contact.projectName) : null,
          contact.projectFolder ? getProjectCacheKey(contact.projectFolder) : null
        ];
        return !keyCandidates.some(key => key && removalKeys.has(key));
      });
    }

    cachedContacts.push(...newContacts);

    enforceProjectRetentionLimit();

    const projectDelta = Math.max(0, cachedProjects.length - previousProjectCount);
    const contactDelta = Math.max(0, cachedContacts.length - previousContactCount);

    let cacheHeartbeat = null;
    try {
      updateProgress(totalFiles, totalFiles, 'Saving cache data...');
      cacheHeartbeat = setInterval(() => {
        updateProgress(totalFiles, totalFiles, 'Saving cache data...');
      }, 5000);
      await saveCacheToIndexedDB();
    } finally {
      if (cacheHeartbeat) {
        clearInterval(cacheHeartbeat);
        cacheHeartbeat = null;
      }
    }

    if (auditEntriesBuffer.length) {
      let auditHeartbeat = null;
      try {
        updateProgress(totalFiles, totalFiles, 'Recording audit trail...');
        auditHeartbeat = setInterval(() => {
          updateProgress(totalFiles, totalFiles, 'Recording audit trail...');
        }, 5000);
        const { inserted = 0 } = await DB.appendContactAuditEntries(auditEntriesBuffer);
        if (inserted) {
          console.log(`[Audit] Logged ${inserted} new contacts for ${stateName}`);
        }
      } catch (error) {
        console.error('[Audit] Failed to persist contact audit entries:', error);
      } finally {
        if (auditHeartbeat) {
          clearInterval(auditHeartbeat);
          auditHeartbeat = null;
        }
        auditEntriesBuffer.length = 0;
      }
    }

    if (trackedAuditBuffer.length) {
      try {
        const { inserted: trackedInserted = 0 } = await DB.appendTrackedAuditEntries(trackedAuditBuffer);
        if (trackedInserted) {
          console.log(`[Audit] Tracked ${trackedInserted} new contacts for monitored projects`);
        }
      } catch (error) {
        console.error('[Audit] Failed to persist tracked audit entries:', error);
      } finally {
        trackedAuditBuffer.length = 0;
      }
    }

    if (totalFiles === 0) {
      updateProgress(1, 1, 'Scan complete (no files)');
    } else {
      updateProgress(totalFiles, totalFiles, 'Scan complete!');
    }

    console.log(`[Scan] State ${stateName} complete: ${projectAggregates.size} projects processed, ${newContacts.length} contacts captured`);

    return {
      success: true,
      state: getFullStateName(stateName),
      newProjects: projectDelta,
      newContacts: contactDelta,
      totalProjects: cachedProjects.length,
      totalContacts: cachedContacts.length,
      scanTime: scanTimestamp
    };
  } catch (error) {
    console.error(`Error scanning state ${stateName}:`, error);
    return { success: false, error: error.message || String(error || 'Unknown error') };
  }
}

async function getAvailableStates() {
  const states = [];

  for (const siteUrl of SHAREPOINT_SITES) {
    try {
      const site = await getSiteByUrl(siteUrl);
      if (!site) continue;

      const drives = await getDrives(site.id);
      const sharedDocs = drives.value?.find(d =>
        d.name === 'Documents' ||
        d.name === 'Shared Documents' ||
        (d.name || '').includes('Shared')
      );

      if (!sharedDocs) continue;

      const topFolders = await getFolders(sharedDocs.id);
      const stateFolders = topFolders.filter(folder => {
        const name = folder.name || '';
        if (!name) return false;
        if (/^\d/.test(name)) return false;
        if (name.startsWith('.')) return false;
        if (name.toLowerCase().includes('archive')) return false;
        return true;
      });

      stateFolders.forEach(folder => {
        const stateName = folder.name.replace(/ Projects?$/i, '').trim();
        const fullName = folder.name || stateName;
        states.push({
          name: stateName,
          fullName,
          site: siteUrl,
          driveId: sharedDocs.id,
          folderId: folder.id
        });
      });
    } catch (error) {
      console.error(`Error getting states from ${siteUrl}:`, error);
    }
  }

  return states;
}
function getScanStatistics() {
  const stats = {
    totalProjects: cachedProjects.length,
    totalContacts: cachedContacts.length,
    uniqueCompanies: [...new Set(cachedContacts.map(c => c.company))].length,
    uniqueDivisions: [...new Set(cachedContacts.map(c => c.division))].length,
    lastScanTime: lastScanTime,
    contactsByState: {},
    contactsByDivision: {}
  };
  
  cachedContacts.forEach(contact => {
    const state = contact.state || 'Unknown';
    stats.contactsByState[state] = (stats.contactsByState[state] || 0) + 1;
  });
  
  cachedContacts.forEach(contact => {
    const division = contact.division || 'Unknown';
    stats.contactsByDivision[division] = (stats.contactsByDivision[division] || 0) + 1;
  });
  
  stats.lastAuditSummary = metadataCache.lastAuditSummary || null;
  return stats;
}

async function hydrateCacheFromRemote(options = {}) {
  const { triggeredBy = 'auto', silent = false } = options || {};
  try {
    const result = await downloadCacheFromSharePoint();
    if (!silent) {
      console.log(`[Cache] Remote cache hydrated (${triggeredBy}) - ${result.projectsLoaded || 0} projects, ${result.contactsLoaded || 0} contacts`);
    }
    return { success: true, result };
  } catch (error) {
    const message = (error && error.message) || String(error || 'Unknown error');
    if (!silent) {
      if (/manifest not found/i.test(message)) {
        console.warn('[Cache] Remote cache manifest not found during hydration');
      } else {
        console.warn(`[Cache] Remote cache hydration failed (${triggeredBy}):`, error);
      }
    }
    return { success: false, error: message };
  }
}

function schedulePeriodicScanAlarm(initialDelayMinutes = CACHE_BOOTSTRAP_DELAY_MINUTES) {
  if (!AUTO_SCAN_ENABLED) {
    console.log('[Scan] Auto scan disabled; periodic scan alarm not scheduled');
    if (typeof chrome !== 'undefined' && chrome.alarms && typeof chrome.alarms.clear === 'function') {
      try {
        chrome.alarms.clear('AB_AUTO_SCAN').catch(() => {});
      } catch (_) {
        // ignore
      }
    }
    return;
  }

  const delay = Math.max(1, Number.isFinite(Number(initialDelayMinutes)) ? Number(initialDelayMinutes) : CACHE_BOOTSTRAP_DELAY_MINUTES);
  try {
    chrome.alarms.create(SCAN_ALARM_NAME, {
      delayInMinutes: delay,
      periodInMinutes: Math.max(5, PERIODIC_SCAN_INTERVAL_MINUTES)
    });
    console.log(`[Scan] Scheduled periodic scan alarm every ${PERIODIC_SCAN_INTERVAL_MINUTES} minutes (delay ${delay}m)`);
  } catch (error) {
    console.error('[Scan] Failed to schedule periodic scan alarm:', error);
  }
}

async function runAutomaticScan(options = {}) {
  const { baseline = false, triggeredBy = 'auto', upload = true } = options || {};

  if (!AUTO_SCAN_ENABLED && triggeredBy !== 'manual') {
    console.log('[Scan] Auto scan disabled; skipping automatic scan triggered by', triggeredBy);
    return { success: false, reason: 'disabled' };
  }

  if (globalScanInProgress) {
    console.warn('[Scan] Automatic scan skipped because another scan is running');
    return { success: false, reason: 'in_progress' };
  }

  const states = await getAvailableStates();
  if (!states || states.length === 0) {
    console.warn('[Scan] No states available for automatic scan');
    return { success: false, reason: 'no_states' };
  }

  globalScanInProgress = true;
  const summary = {
    baseline,
    triggeredBy,
    states: [],
    failures: []
  };

  try {
    for (const stateEntry of states) {
      const siteUrl = stateEntry.site || stateEntry.siteUrl || stateEntry.siteURL;
      const stateName = stateEntry.name || stateEntry.state || stateEntry.fullName;
      if (!siteUrl || !stateName) continue;

      console.log(`[Scan] ${baseline ? 'Baseline' : 'Incremental'} scan for ${stateName} (${siteUrl})`);
      const stateResult = await scanSingleState(siteUrl, stateName, null);
      if (stateResult && stateResult.success) {
        summary.states.push({
          state: stateResult.state || stateName,
          newProjects: stateResult.newProjects || 0,
          newContacts: stateResult.newContacts || 0,
          totalProjects: stateResult.totalProjects || cachedProjects.length,
          totalContacts: stateResult.totalContacts || cachedContacts.length
        });
      } else {
        const errorMessage = stateResult?.error || 'Unknown scan error';
        summary.states.push({ state: stateName, error: errorMessage });
        summary.failures.push({ state: stateName, error: errorMessage });
      }
    }

    const finishedAt = new Date().toISOString();
    lastScanTime = finishedAt;
    if (baseline) {
      lastBaselineScan = finishedAt;
      lastIncrementalScan = finishedAt;
    } else {
      lastIncrementalScan = finishedAt;
    }

    await saveCacheToIndexedDB();

    if (upload) {
      try {
        const uploadResult = await uploadCacheToSharePoint();
        summary.remoteManifest = uploadResult?.manifest || null;
      } catch (uploadError) {
        console.warn('[Scan] Upload to SharePoint failed:', uploadError);
        summary.uploadError = uploadError && uploadError.message ? uploadError.message : String(uploadError || 'Upload failed');
      }
    }

    summary.success = summary.failures.length === 0;
    return summary;
  } catch (error) {
    console.error('[Scan] Automatic scan failed:', error);
    return { success: false, error: error.message || String(error || 'Unknown error') };
  } finally {
    globalScanInProgress = false;
  }
}

function needsBaselineScan() {
  if (!lastBaselineScan) return true;
  const baselineDate = parseIsoDate(lastBaselineScan);
  if (!baselineDate) return true;
  return (Date.now() - baselineDate.getTime()) > (BASELINE_MAX_AGE_HOURS * 60 * 60 * 1000);
}

function needsIncrementalScan() {
  if (!lastIncrementalScan) return true;
  const incrementalDate = parseIsoDate(lastIncrementalScan);
  if (!incrementalDate) return true;
  return (Date.now() - incrementalDate.getTime()) > minutesToMs(INCREMENTAL_MAX_AGE_MINUTES);
}

async function bootstrapCache(options = {}) {
  const { triggeredBy = 'bootstrap', hydrateRemote = true } = options || {};

  await loadCacheFromIndexedDB();

  if (hydrateRemote) {
    await hydrateCacheFromRemote({ triggeredBy, silent: triggeredBy === 'startup' });
  }

  if (AUTO_SCAN_ENABLED) {
    schedulePeriodicScanAlarm(triggeredBy === 'onInstalled' ? CACHE_BOOTSTRAP_DELAY_MINUTES : 5);
  } else {
    console.log('[Scan] Auto scan disabled; periodic alarm not scheduled during bootstrap');
    if (typeof chrome !== 'undefined' && chrome.alarms && typeof chrome.alarms.clear === 'function') {
      try {
        chrome.alarms.clear('AB_AUTO_SCAN').catch(() => {});
      } catch (_) {
        // ignore
      }
    }
  }
  scheduleDailyAuditRun();
}

async function exportContactsToExcel(contacts) {
  try {
    if (typeof XLSX === 'undefined' || !XLSX.utils) {
      throw new Error('Excel library not loaded');
    }
    
    const wb = XLSX.utils.book_new();
    
    const wsData = [
      ['Company', 'Contact Name', 'Email', 'Phone', 'Division/Trade', 'Project', 
       'City', 'State', 'Zip', 'Bid Amount', 'Job Name', 'Address', 'Bid Date']
    ];
    
    contacts.forEach(contact => {
      wsData.push([
        contact.company || '',
        contact.name || '',
        contact.email || '',
        contact.phone || '',
        contact.division || '',
        contact.project || '',
        contact.city || '',
        contact.state || '',
        contact.zip || '',
        contact.bidAmount || '',
        contact.jobName || '',
        contact.address || '',
        contact.bidDate || ''
      ]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    ws['!cols'] = [
      {wch: 25}, {wch: 20}, {wch: 25}, {wch: 15}, {wch: 20},
      {wch: 30}, {wch: 15}, {wch: 10}, {wch: 10}, {wch: 12},
      {wch: 30}, {wch: 35}, {wch: 12}
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, 'Contacts');
    
    const excelBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    
    const base64 = btoa(
      new Uint8Array(excelBuffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
    );
    const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const filename = `AB_Contacts_Export_${timestamp}.xlsx`;
    
    await chrome.downloads.download({
      url: dataUrl,
      filename: filename,
      saveAs: true
    });
    
    return { success: true };
    
  } catch (error) {
    console.error('Export error:', error);
    return { success: false, error: error.message };
  }
}

async function findExcelFileForFolder(driveId, folderId, depth = 0) {
  if (depth > 2) return null;
  try {
    const files = await getFiles(driveId, folderId);
    const validFiles = (files || []).filter(file => {
      if (!file || !file.name || file.name.startsWith('~$')) return false;
      if (file.size && file.size > 15728640) return false;
      return true;
    });
    if (validFiles.length > 0) {
      validFiles.sort((a, b) => new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0));
      return validFiles[0];
    }

    const childFolders = await getFolders(driveId, folderId);
    for (const folder of childFolders) {
      const nested = await findExcelFileForFolder(driveId, folder.id, depth + 1);
      if (nested) return nested;
    }
  } catch (error) {
    console.error('[Discovery] Error locating files for folder', folderId, error);
  }
  return null;
}

function scoreProjectMatch(targetNormalized, targetCompact, candidateNormalized, candidateCompact) {
  if (!candidateNormalized) return 0;
  if (candidateNormalized === targetNormalized) return 120;
  if (candidateCompact === targetCompact) return 115;
  if (candidateNormalized.startsWith(targetNormalized) || targetNormalized.startsWith(candidateNormalized)) return 105;
  if (candidateCompact.startsWith(targetCompact) || targetCompact.startsWith(candidateCompact)) return 100;
  if (candidateNormalized.includes(targetNormalized) || targetNormalized.includes(candidateNormalized)) return 90;
  if (candidateCompact.includes(targetCompact) || targetCompact.includes(candidateCompact)) return 85;
  return 0;
}

async function discoverProjectLocation(projectName) {
  const normalizedTarget = normalizeProjectName(projectName);
  if (!normalizedTarget) return null;
  const compactTarget = compactProjectName(projectName);

  let bestMatch = null;

  for (const siteUrl of SHAREPOINT_SITES) {
    try {
      const site = await getSiteByUrl(siteUrl);
      if (!site) continue;

      const drivesResponse = await getDrives(site.id);
      const driveItems = drivesResponse?.value || [];
      const sharedDocs = driveItems.find(d => {
        const name = d.name || '';
        return name === 'Documents' || name === 'Shared Documents' || name.includes('Shared');
      });

      if (!sharedDocs) continue;

      const stateFolders = await getFolders(sharedDocs.id);
      const filteredStateFolders = stateFolders.filter(folder => {
        const name = folder.name || '';
        if (!name) return false;
        if (/^\d/.test(name)) return false;
        if (name.startsWith('.')) return false;
        if (name.toLowerCase().includes('archive')) return false;
        return true;
      });

      for (const stateFolder of filteredStateFolders) {
        const projectFolders = await getFolders(sharedDocs.id, stateFolder.id);
        for (const projectFolder of projectFolders) {
          const candidateNormalized = normalizeProjectName(projectFolder.name);
          if (!candidateNormalized) continue;
          const candidateCompact = candidateNormalized.replace(/\s+/g, '');
          const score = scoreProjectMatch(
            normalizedTarget,
            compactTarget,
            candidateNormalized,
            candidateCompact
          );

          if (score <= 0) continue;

          if (!bestMatch || score > bestMatch.score) {
            bestMatch = {
              score,
              siteUrl,
              driveId: sharedDocs.id,
              stateFolder,
              projectFolder
            };

            if (score >= 120) {
              break;
            }
          }
        }

        if (bestMatch && bestMatch.score >= 120) {
          break;
        }
      }

      if (bestMatch && bestMatch.score >= 120) {
        break;
      }
    } catch (error) {
      console.error(`[Discovery] Error searching for project in ${siteUrl}:`, error);
    }
  }

  if (!bestMatch || bestMatch.score < 85) {
    return null;
  }

  const chosenFile = await findExcelFileForFolder(bestMatch.driveId, bestMatch.projectFolder.id);
  if (!chosenFile) {
    return null;
  }

  const stateName = (bestMatch.stateFolder.name || '').replace(/ Projects?$/i, '').trim();

  console.log('[Discovery] Found project folder', bestMatch.projectFolder.name, 'in', bestMatch.siteUrl, 'score:', bestMatch.score);

  return {
    name: bestMatch.projectFolder.name || projectName,
    displayName: bestMatch.projectFolder.name || projectName,
    state: stateName,
    site: bestMatch.siteUrl,
    driveId: bestMatch.driveId,
    fileId: chosenFile.id,
    fileName: chosenFile.name,
    projectFolderId: bestMatch.projectFolder.id,
    stateFolderId: bestMatch.stateFolder.id,
    projectFolderName: bestMatch.projectFolder.name,
    stateFolderName: bestMatch.stateFolder.name,
    lastDiscovered: new Date().toISOString()
  };
}

async function ensureProjectInCache(projectName) {
  const normalizedTarget = normalizeProjectName(projectName);
  const compactTarget = compactProjectName(projectName);

  // If we already have a runtime project loaded and it matches, reuse it
  if (currentProjectFile && currentProjectFile.projectName) {
    const currentNameNorm = normalizeProjectName(currentProjectFile.projectName);
    const currentCompact = compactProjectName(currentProjectFile.projectName);
    if (currentNameNorm === normalizedTarget || currentCompact === compactTarget) {
      const runtimeProject = {
        name: currentProjectFile.projectName,
        displayName: currentProjectFile.projectName,
        projectFolderName: currentProjectFile.projectFolderName || currentProjectFile.projectFolderId,
        projectFolderId: currentProjectFile.projectFolderId,
        stateFolderId: currentProjectFile.stateFolderId,
        driveId: currentProjectFile.driveId,
        fileId: currentProjectFile.id,
        fileName: currentProjectFile.name,
        site: currentProjectFile.site,
        state: currentProjectFile.state
      };
      const runtime = getOrCreateProjectRuntime(runtimeProject.name);
      runtime.project = runtimeProject;
      return { project: runtimeProject, source: 'runtime' };
    }
  }

  const matchesTarget = (project) => {
    const candidates = [
      project?.name,
      project?.displayName,
      project?.projectFolderName,
      project?.projectName
    ].filter(Boolean);
    return candidates.some(candidate => {
      const norm = normalizeProjectName(candidate);
      if (norm && norm === normalizedTarget) return true;
      const comp = compactProjectName(candidate);
      return comp && comp === compactTarget;
    });
  };

  const existingIndex = cachedProjects.findIndex(matchesTarget);

  let project = existingIndex >= 0 ? cachedProjects[existingIndex] : null;
  let source = 'cache';

  // If we have a usable cached project (with drive/file info), return it
  if (project && project.driveId && project.fileId && project.projectFolderId) {
    const runtime = getOrCreateProjectRuntime(project.name || project.projectFolderName || project.displayName || projectName);
    runtime.project = project;
    return { project, source };
  }

  // Otherwise attempt discovery
  const discovered = await discoverProjectLocation(projectName);
  if (!discovered) {
    return { project, source: 'missing' };
  }

  if (project) {
    Object.assign(project, discovered);
    source = 'updated';
  } else {
    project = { ...discovered };
    cachedProjects.push(project);
    enforceProjectRetentionLimit();
    source = 'discovered';
  }

  project.displayName = project.displayName || project.projectFolderName || project.name || projectName;
  if (!project.lastDiscovered) {
    project.lastDiscovered = new Date().toISOString();
  }
  project.lastRefreshed = new Date().toISOString();

  const runtime = getOrCreateProjectRuntime(project.displayName || project.name || projectName);
  runtime.project = project;

  try {
    await DB.saveProjects(cachedProjects);
  } catch (error) {
    console.error('[Discovery] Failed to persist updated projects cache:', error);
  }

  return { project, source };
}
async function loadProjectWorkbook(projectName, options = {}) {
  const { includeContacts = false, forceRefresh = false, forceContactRefresh = false } = options;

  try {
    console.log('[Excel API] Loading project workbook:', projectName);

    const { project, source } = await ensureProjectInCache(projectName);

    if (!project) {
      throw new Error('Project not found in SharePoint. Please verify the name.');
    }

    if (!project.driveId || !project.fileId) {
      throw new Error('Project file location not available.');
    }

    const resolvedName = project.name || projectName;
    const runtime = getOrCreateProjectRuntime(resolvedName);
    runtime.project = project;

    let mappingSource = 'cache';

    const hasProjectMapping = project.divisionMapping && project.divisions && Object.keys(project.divisionMapping).length > 0;
    const hasRuntimeMapping = runtime.divisionMapping && runtime.divisions && Object.keys(runtime.divisionMapping).length > 0;

    const projectMappingComplete = hasProjectMapping && !isDivisionMappingIncomplete(project.divisionMapping, project.divisions);
    const runtimeMappingComplete = hasRuntimeMapping && !isDivisionMappingIncomplete(runtime.divisionMapping, runtime.divisions);

    if (forceRefresh || (!projectMappingComplete && !runtimeMappingComplete)) {
      await buildDivisionMapping({ project, runtime, forceRefresh: true });
      mappingSource = 'graph';
    } else if (projectMappingComplete) {
      divisionMapping = project.divisionMapping;
      runtime.divisionMapping = project.divisionMapping;
      runtime.divisions = project.divisions;
      runtime.lastDivisionSync = Date.now();
      mappingSource = 'project-cache';
    } else if (runtimeMappingComplete) {
      divisionMapping = runtime.divisionMapping;
      mappingSource = 'runtime-cache';
    } else {
      await buildDivisionMapping({ project, runtime, forceRefresh });
      mappingSource = 'graph';
    }

    currentProjectFile = {
      id: project.fileId,
      name: project.fileName || `${resolvedName}.xlsx`,
      driveId: project.driveId,
      projectName: resolvedName,
      projectFolderId: project.projectFolderId,
      stateFolderId: project.stateFolderId,
      site: project.site,
      state: project.state
    };

    project.displayName = project.displayName || project.name || project.projectFolderName || project.fileName || resolvedName;

    console.log('[Excel API] Using project record from', source, ':', project);
    console.log('[Excel API] Using driveId:', project.driveId, 'fileId:', project.fileId);

    let contactsByDivision;
    if (includeContacts) {
      const shouldRefreshContacts = forceContactRefresh
        || !runtime.contactsByDivision
        || !runtime.lastContactsSync
        || (Date.now() - (runtime.lastContactsSync || 0) > 5 * 60 * 1000)
        || !project.contactsByDivision;

      if (shouldRefreshContacts) {
        contactsByDivision = await loadDivisionContactsInternal({ project, runtime, forceRefresh: forceContactRefresh });
      } else {
        contactsByDivision = runtime.contactsByDivision || project.contactsByDivision;
      }
    }

    return {
      success: true,
      fileId: currentProjectFile.id,
      fileName: currentProjectFile.name,
      divisions: runtime.divisions || project.divisions || [],
      project,
      source,
      divisionSource: mappingSource,
      contactsByDivision: includeContacts ? contactsByDivision : undefined
    };

  } catch (error) {
    console.error('[Excel API] Error loading workbook:', error);
    return { success: false, error: error.message };
  }
}

async function buildDivisionMapping(options = {}) {
  const { project = null, runtime = null, forceRefresh = false, includeContacts = false } = options;

  if ((!currentProjectFile || !currentProjectFile.id || !currentProjectFile.driveId) && project && project.fileId && project.driveId) {
    currentProjectFile = {
      id: project.fileId,
      name: project.fileName || `${project.name || 'project'}.xlsx`,
      driveId: project.driveId,
      projectName: project.name || '',
      projectFolderId: project.projectFolderId,
      stateFolderId: project.stateFolderId,
      site: project.site,
      state: project.state
    };
  }

  if (!currentProjectFile || !currentProjectFile.id || !currentProjectFile.driveId) {
    throw new Error('No project file loaded - cannot build division mapping');
  }

  try {
    console.log('[Excel API] Building division mapping...', forceRefresh ? '(forced)' : '');

    const sheetsData = await callGraphAPI(
      `/drives/${currentProjectFile.driveId}/items/${currentProjectFile.id}/workbook/worksheets`,
      { headers: getWorkbookRequestHeaders() }
    );
    const sheets = sheetsData.value || [];

    console.log('[Excel API] Found', sheets.length, 'worksheets');

    const projectContext = project || (runtime && runtime.project) || null;
    const mappingEntries = [];

    for (const sheet of sheets) {
      const sheetName = sheet.name;
      if (!sheetName) {
        continue;
      }
      if (EXCLUDED_SHEETS.some(entry => entry.toLowerCase() === sheetName.toLowerCase())) {
        continue;
      }

      const info = getDivisionDisplayInfo(sheetName, projectContext);
      if (info.codeValue !== null) {
        if (!isDivisionCodeWithinRange(info.codeValue)) {
          continue;
        }
      } else if (!isSheetWithinFallbackRange(sheetName)) {
        continue;
      }

      const orderValue = info.codeValue !== null
        ? info.codeValue
        : (info.fallbackIndex !== -1 ? info.fallbackIndex + 1 : Number.MAX_SAFE_INTEGER);

      mappingEntries.push({
        key: sheetName,
        sheetId: sheet.id,
        sheetName,
        displayName: info.displayName,
        startsOnRow4: SPECIAL_SHEETS.includes(sheetName),
        codeValue: info.codeValue,
        order: orderValue
      });
    }

    mappingEntries.sort((a, b) => {
      if (a.order !== b.order) return a.order - b.order;
      return a.displayName.localeCompare(b.displayName, undefined, { sensitivity: 'base' });
    });

    const mapping = {};
    const divisions = mappingEntries.map(entry => {
      mapping[entry.key] = {
        key: entry.key,
        sheetId: entry.sheetId,
        sheetName: entry.sheetName,
        displayName: entry.displayName,
        startsOnRow4: entry.startsOnRow4,
        codeValue: entry.codeValue,
        order: entry.order,
        lastSync: Date.now()
      };

      return {
        key: entry.key,
        name: entry.displayName,
        displayName: entry.displayName,
        sheetName: entry.sheetName,
        startsOnRow4: entry.startsOnRow4,
        codeValue: entry.codeValue,
        order: entry.order
      };
    });

    divisionMapping = mapping;

    if (runtime) {
      runtime.divisionMapping = mapping;
      runtime.divisions = divisions;
      runtime.lastDivisionSync = Date.now();
    }

    if (project) {
      project.divisionMapping = mapping;
      project.divisions = divisions;
      project.lastDivisionSync = new Date().toISOString();
      try {
        await DB.saveProjects(cachedProjects);
      } catch (error) {
        console.error('[Excel API] Error persisting project division mapping:', error);
      }
    }

    console.log('[Excel API] Division mapping complete:', divisions.length, 'divisions mapped');

    if (includeContacts) {
      await loadDivisionContactsInternal({ project, runtime, forceRefresh: true });
    }

    return { mapping, divisions };

  } catch (error) {
    console.error('[Excel API] Error building division mapping:', error);
    throw error;
  }
}

function normalizeMatchValue(value) {
  if (!value) return '';
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizePhoneDigits(value) {
  if (!value) return '';
  return String(value).replace(/[^0-9]/g, '');
}

function extractEmailDomain(email) {
  if (!email || typeof email !== 'string' || !email.includes('@')) return '';
  return email.split('@')[1].toLowerCase();
}

async function loadDivisionContactsInternal({ project = null, runtime = null, forceRefresh = false } = {}) {
  const contactsByDivision = {};

  for (const [divisionKey, division] of Object.entries(divisionMapping || {})) {
    if (!division) continue;
    if (!forceRefresh && division.contacts && division.contacts.length > 0) {
      contactsByDivision[divisionKey] = division.contacts;
      continue;
    }

    const contacts = await readDivisionContactsForMapping(division);
    division.contacts = contacts;
    contactsByDivision[divisionKey] = contacts;
  }

  if (runtime) {
    runtime.contactsByDivision = contactsByDivision;
    runtime.lastContactsSync = Date.now();
  }

  if (project) {
    project.contactsByDivision = contactsByDivision;
    project.lastContactsSync = new Date().toISOString();
    try {
      await DB.saveProjects(cachedProjects);
    } catch (error) {
      console.error('[Excel API] Error persisting project contacts cache:', error);
    }
  }

  return contactsByDivision;
}

async function readDivisionContactsForMapping(division) {
  const contacts = [];
  const sheetName = division.sheetName;
  const startsOnRow4 = division.startsOnRow4;
  const startRow = startsOnRow4 ? 4 : 3;
  const endRow = startRow + 4; // company, name, phone, email, bid amount
  const displayName = division.displayName || division.name || sheetName;
  const divisionKey = division.key || division.sheetName;

  const rangeAddress = `G${startRow}:AZ${endRow}`;
  const rangeUrl = `${GRAPH_API}/drives/${currentProjectFile.driveId}/items/${currentProjectFile.id}/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${rangeAddress}')`;

  try {
    const rangeData = await callGraphAPI(rangeUrl, { headers: getWorkbookRequestHeaders() });
    const rows = rangeData.values || [];
    if (rows.length === 0) return contacts;

    const companyRow = rows[0] || [];
    const nameRow = rows[1] || [];
    const phoneRow = rows[2] || [];
    const emailRow = rows[3] || [];
    const bidRow = rows[4] || [];

    const maxColumns = Math.max(companyRow.length, nameRow.length, phoneRow.length, emailRow.length, bidRow.length);

    for (let i = 0; i < maxColumns; i++) {
      const company = (companyRow[i] || '').toString().trim();
      const contactName = (nameRow[i] || '').toString().trim();
      const phone = (phoneRow[i] || '').toString().trim();
      const email = (emailRow[i] || '').toString().trim();
      const bidAmount = (bidRow[i] || '').toString().trim();

      if (!company && !email && !phone && !contactName) {
        continue;
      }

      const columnIndex = 6 + i; // Column G is index 6
      const columnLetter = getColumnLetter(columnIndex);

      contacts.push({
        divisionKey,
        sheetName,
        division: displayName,
        columnLetter,
        company,
        companyNormalized: normalizeMatchValue(company),
        contactName,
        name: contactName,
        phone,
        phoneDigits: normalizePhoneDigits(phone),
        email,
        emailDomain: extractEmailDomain(email),
        bidAmount,
        lastSeen: new Date().toISOString()
      });
    }
  } catch (error) {
    console.error('[Excel API] Failed to read contacts for division', sheetName, error);
  }

  return contacts;
}

// ============ CONTACT AUDIT WITH VERSION HISTORY ============

// Helpers for tracked entries scan
function buildProjectDisplayLabel(project = {}, fallbackKey = '') {
  const toString = (value) => {
    if (value === null || value === undefined) return '';
    return typeof value === 'string' ? value.trim() : String(value).trim();
  };

  const folder = toString(project.projectFolderName);
  const number = toString(project.projectNumber);
  const display = toString(project.displayName);
  const name = toString(project.name);

  let primary = folder || number || display || name || toString(fallbackKey) || toString(project.projectKey) || toString(project.key);
  if (!primary) {
    return 'Unknown Project';
  }

  const secondaryCandidates = [display, name];
  for (const candidate of secondaryCandidates) {
    if (!candidate) continue;
    const lowerCandidate = candidate.toLowerCase();
    const lowerPrimary = primary.toLowerCase();
    if (lowerCandidate === lowerPrimary) continue;
    if (lowerCandidate.startsWith(lowerPrimary)) continue;
    primary = `${primary} ${candidate}`.trim();
    break;
  }

  return primary || 'Unknown Project';
}

function pickFirstValidDate(...values) {
  for (const value of values) {
    if (!value) continue;
    let date;
    if (value instanceof Date) {
      date = value;
    } else if (typeof value === 'number') {
      date = new Date(value);
    } else {
      const trimmed = String(value).trim();
      if (!trimmed) continue;
      date = new Date(trimmed);
    }
    if (!Number.isNaN(date.getTime())) {
      return date;
    }
  }
  return null;
}

function pickFirstNonEmpty(...values) {
  for (const value of values) {
    if (value === null || value === undefined) continue;
    const text = String(value).trim();
    if (text) {
      return text;
    }
  }
  return '';
}

function buildProjectFileContext(project = {}) {
  if (!project || typeof project !== 'object') {
    return {};
  }

  const driveId = project.driveId
    || project.parentDriveId
    || (project.parentReference && project.parentReference.driveId)
    || (project.drive && project.drive.id)
    || project.siteDriveId
    || null;

  const fileId = project.fileId
    || project.id
    || project.itemId
    || (project.file && project.file.id)
    || null;

  const fileName = project.fileName
    || project.displayName
    || project.projectFolderName
    || project.name
    || project.projectName
    || '';

  const lastModifiedDateTime = project.fileModified
    || project.lastModifiedDateTime
    || project.modifiedTime
    || project.lastUpdated
    || project.lastRefreshed
    || project.detectedAt
    || null;

  let lastModifiedByCandidate = project.lastModifiedByDetails
    || project.lastModifiedBy
    || project.modifiedBy
    || null;

  let lastModifiedBy = null;
  if (lastModifiedByCandidate) {
    if (typeof lastModifiedByCandidate === 'object') {
      lastModifiedBy = lastModifiedByCandidate;
    } else {
      const email = String(lastModifiedByCandidate).trim();
      if (email) {
        lastModifiedBy = { user: { email } };
      }
    }
  }

  const parentReference = project.parentReference || null;
  const webUrl = project.webUrl || project.fileUrl || project.siteUrl || null;

  return {
    driveId,
    id: fileId,
    name: fileName,
    fileName,
    lastModifiedDateTime,
    lastModifiedBy,
    parentReference,
    webUrl
  };
}


/**
 * Scan selected projects and collect all current entries with version history info
 */
async function scanSelectedProjectsForTrackedEntries(selectedProjectKeys = []) {
  try {
    console.log('[TrackedEntries] Starting scan for projects:', selectedProjectKeys);

    if (!selectedProjectKeys || !selectedProjectKeys.length) {
      throw new Error('No projects selected. Please select at least one project before scanning.');
    }

    let token = accessToken;
    if (!token) {
      const stored = await chrome.storage.local.get(['accessToken']);
      token = stored.accessToken;
    }
    if (!token) {
      throw new Error('No access token available. Please authenticate first.');
    }

    const projectsToScan = Array.from(selectedProjectKeys);
    const trackedEntries = {};
    const versionLookupCache = new Map();

    const projectCache = new Map();
    projectsToScan.forEach(projectKey => {
      const project = resolveTrackedProjectContext(projectKey);
      if (project) {
        projectCache.set(projectKey, project);
      } else {
        console.warn('[TrackedEntries] Cached project missing for key', projectKey);
      }
    });

    const totalContactsToScan = Array.from(projectCache.values()).reduce((sum, project) => {
      const contactsByDivision = project?.contactsByDivision || {};
      const count = Object.values(contactsByDivision).reduce((c, list) => c + (Array.isArray(list) ? list.length : 0), 0);
      return sum + count;
    }, 0) || projectsToScan.length;

    let processedContacts = 0;
    const scanStartTime = Date.now();

    const updateProgress = (status, currentProject) => {
      const percent = Math.min(100, Math.round((processedContacts / totalContactsToScan) * 100));
      const elapsedSeconds = Math.max(1, (Date.now() - scanStartTime) / 1000);
      const avgPerContact = processedContacts ? (elapsedSeconds / processedContacts) : 0;
      const remainingContacts = Math.max(0, totalContactsToScan - processedContacts);
      const etaSeconds = avgPerContact ? Math.round(remainingContacts * avgPerContact) : 0;
      const etaLabel = etaSeconds > 120 ? `${Math.round(etaSeconds / 60)}m` : `${Math.max(0, etaSeconds)}s`;

      broadcastScanProgress({
        status,
        percent,
        currentProject,
        totalEntriesFound: processedContacts,
        eta: etaLabel
      });
    };

    updateProgress('Preparing scan...', '');

    for (const projectKey of projectsToScan) {
      const project = projectCache.get(projectKey);
      if (!project) {
        console.warn(`[TrackedEntries] Project not found: ${projectKey}`);
        continue;
      }

      const projectLabel = buildProjectDisplayLabel(project, projectKey);
      const projectFileInfo = buildProjectFileContext(project);

      updateProgress(`Scanning ${projectLabel}...`, projectLabel);

      let trackedHistoryByEmail = null;
      try {
        const historyEntries = await DB.getTrackedAuditEntries({ projectKey, limit: 5000 });
        if (Array.isArray(historyEntries) && historyEntries.length) {
          trackedHistoryByEmail = new Map();
          historyEntries.forEach(entry => {
            const emailNormalized = String(entry.emailNormalized || entry.email || '').trim().toLowerCase();
            if (!emailNormalized) return;
            const detectedDate = pickFirstValidDate(entry.detectedAt, entry.lastModifiedDateTime, entry.trackedAt);
            const timestamp = detectedDate ? new Date(detectedDate).getTime() : Number.MAX_SAFE_INTEGER;
            const existing = trackedHistoryByEmail.get(emailNormalized);
            if (!existing || timestamp < existing.timestamp) {
              trackedHistoryByEmail.set(emailNormalized, { entry, timestamp });
            }
          });
        }
      } catch (error) {
        console.warn('[TrackedEntries] Failed to load tracked audit history for', projectKey, error);
      }

      const simplifiedVersionInfo = await getContactAddedInfoSimplified(projectFileInfo);
      const projectLastModifiedByDisplay = pickFirstNonEmpty(
        simplifiedVersionInfo?.addedByDisplayName,
        projectFileInfo.lastModifiedBy?.user?.displayName,
        projectFileInfo.lastModifiedBy?.application?.displayName
      );

      const projectLastModifiedByEmailRaw = pickFirstNonEmpty(
        simplifiedVersionInfo?.addedByEmail,
        simplifiedVersionInfo?.addedBy,
        projectFileInfo.lastModifiedBy?.user?.email,
        project.lastModifiedBy,
        project.modifiedBy
      );
      const projectLastModifiedByEmail = normalizeEmail(projectLastModifiedByEmailRaw);

      const projectLastModifiedBy = pickFirstNonEmpty(
        projectLastModifiedByDisplay,
        projectLastModifiedByEmail
      );

      const projectEntries = [];
      const contactsByDivision = project.contactsByDivision || {};
      const projectDivisions = Array.isArray(project.divisions) ? project.divisions : [];
      const driveId = projectFileInfo.driveId;
      const fileId = projectFileInfo.id;

      for (const [divisionKey, contacts] of Object.entries(contactsByDivision)) {
        if (!Array.isArray(contacts)) continue;

        const normalizedDivisionKey = String(divisionKey || '').toLowerCase();
        const divisionInfo = projectDivisions.find(div => {
          const divKey = String(div?.key || '').toLowerCase();
          const sheetName = String(div?.sheetName || '').toLowerCase();
          const name = String(div?.name || div?.displayName || '').toLowerCase();
          return (divKey && divKey === normalizedDivisionKey)
            || (sheetName && sheetName === normalizedDivisionKey)
            || (name && name === normalizedDivisionKey);
        }) || null;

        for (const contact of contacts) {
          try {
            const rawEmail = typeof contact.email === 'string' ? contact.email : '';
            const emailValue = rawEmail.trim();
            const emailNormalized = emailValue ? emailValue.toLowerCase() : '';
            const hasEmail = Boolean(emailNormalized);
            const trackedSnapshotInfo = (hasEmail && trackedHistoryByEmail) ? trackedHistoryByEmail.get(emailNormalized) : null;
            const trackedSnapshot = trackedSnapshotInfo ? trackedSnapshotInfo.entry : null;

            let versionLookupInfo = null;
            const sheetNameCandidate = (contact.sheetName || divisionInfo?.sheetName || divisionInfo?.name || divisionKey || '').trim();
            const canLookupVersion = ENABLE_VERSION_HISTORY_LOOKUP && Boolean(driveId && fileId && token && sheetNameCandidate && hasEmail);
            const trackedHasAuthor = Boolean(trackedSnapshot?.addedBy || trackedSnapshot?.lastModifiedBy || trackedSnapshot?.sharePointUser);
            const trackedHasDating = Boolean(trackedSnapshot?.detectedAt || trackedSnapshot?.lastModifiedDateTime || trackedSnapshot?.trackedAt);
            const hasMeaningfulContactDate = Boolean(contact.addedDate || contact.lastSeen || contact.createdAt);
            const shouldLookupVersion = canLookupVersion && (!trackedHasAuthor || !trackedHasDating || !hasMeaningfulContactDate || !trackedSnapshot);

            if (shouldLookupVersion) {
              const cacheKey = `${projectKey}|${sheetNameCandidate.toLowerCase()}|${emailNormalized}`;
              if (versionLookupCache.has(cacheKey)) {
                versionLookupInfo = versionLookupCache.get(cacheKey);
              } else {
                try {
                  versionLookupInfo = await getContactAddedInfoFromVersionHistory(sheetNameCandidate, emailValue, projectFileInfo, token);
                } catch (error) {
                  console.warn('[TrackedEntries] Version history lookup failed', projectKey, emailValue, error);
                  versionLookupInfo = null;
                }
                versionLookupCache.set(cacheKey, versionLookupInfo);
              }
            }

            const primaryDate = pickFirstValidDate(
              contact.addedDate,
              contact.detectedAt,
              contact.lastSeen,
              contact.createdAt,
              trackedSnapshot?.detectedAt,
              trackedSnapshot?.lastModifiedDateTime,
              trackedSnapshot?.trackedAt,
              trackedSnapshotInfo?.timestamp,
              versionLookupInfo?.addedDate
            );

            let resolvedDate = primaryDate ? new Date(primaryDate) : null;
            if ((!resolvedDate || Number.isNaN(resolvedDate.getTime())) && simplifiedVersionInfo?.addedDate) {
              const simplifiedCandidate = new Date(simplifiedVersionInfo.addedDate);
              if (!Number.isNaN(simplifiedCandidate.getTime())) {
                const ageMs = Math.abs(Date.now() - simplifiedCandidate.getTime());
                if (ageMs > 6 * 60 * 60 * 1000) {
                  resolvedDate = simplifiedCandidate;
                }
              }
            }

            const validDate = resolvedDate && !Number.isNaN(resolvedDate.getTime()) ? resolvedDate : null;
            const dateAddedLabel = validDate ? validDate.toLocaleDateString('en-US') : 'Unknown';
            const timeAddedLabel = validDate ? validDate.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }) : 'Unknown';
            const timestampValue = validDate ? validDate.getTime() : 0;

            const sharePointUserRaw = pickFirstNonEmpty(
              contact.addedBy,
              contact.lastModifiedBy,
              trackedSnapshot?.lastModifiedBy,
              trackedSnapshot?.addedBy,
              trackedSnapshot?.sharePointUser,
              trackedSnapshot?.detectedBy,
              versionLookupInfo?.addedBy,
              simplifiedVersionInfo?.addedBy,
              projectLastModifiedBy
            );

            const sharePointDisplayCandidate = pickFirstNonEmpty(
              trackedSnapshot?.sharePointUserDisplay,
              trackedSnapshot?.addedByDisplayName,
              versionLookupInfo?.addedByDisplayName,
              simplifiedVersionInfo?.addedByDisplayName,
              projectLastModifiedByDisplay
            );

            const sharePointEmailCandidate = pickFirstNonEmpty(
              trackedSnapshot?.sharePointUserEmail,
              trackedSnapshot?.sharePointUser,
              trackedSnapshot?.addedBy,
              versionLookupInfo?.addedByEmail,
              versionLookupInfo?.addedBy,
              simplifiedVersionInfo?.addedByEmail,
              simplifiedVersionInfo?.addedBy,
              projectLastModifiedByEmail,
              sharePointUserRaw
            );

            const sharePointUser = resolveSharePointUserName(sharePointUserRaw || 'Unknown', {
              displayName: sharePointDisplayCandidate,
              email: sharePointEmailCandidate
            });

            const fallbackEmailForId = hasEmail ? emailValue : `missing-email-${processedContacts}`;
            const entryId = generateEntryId({
              ...contact,
              email: fallbackEmailForId,
              divisionKey: contact.divisionKey || divisionKey
            });
            const displayEmail = hasEmail ? emailValue : ((contact.email && contact.email.trim()) || 'No Email');

            projectEntries.push({
              id: entryId,
              division: contact.division || contact.sheetName || divisionInfo?.displayName || divisionInfo?.name || divisionKey,
              divisionKey: contact.divisionKey || divisionKey,
              company: contact.company,
              contactName: contact.name,
              phone: contact.phone,
              email: displayEmail,
              sharePointUser,
              sharePointUserDisplay: sharePointUser,
              sharePointUserEmail: sharePointEmailCandidate || sharePointUserRaw || '',
              dateAdded: dateAddedLabel,
              timeAdded: timeAddedLabel,
              timestamp: timestampValue,
              bidAmount: contact.bidAmount
            });

            processedContacts += 1;
            updateProgress(`Scanning ${projectLabel}: ${emailValue}`, projectLabel);
          } catch (error) {
            console.error(`[TrackedEntries] Error processing contact ${contact.email}:`, error);
          }
        }
      }

      trackedEntries[projectKey] = {
        projectName: projectLabel,
        lastScanned: new Date().toISOString(),
        entries: projectEntries
      };

      console.log(`[TrackedEntries] Found ${projectEntries.length} entries for project ${projectLabel}`);
    }

    updateProgress('Scan complete!', '');

    await chrome.storage.local.set({ AB_TRACKED_ENTRIES: trackedEntries });
    console.log('[TrackedEntries] Scan complete. Saved to storage.');

    return {
      success: true,
      data: trackedEntries,
      totalEntries: processedContacts
    };

  } catch (error) {
    console.error('[TrackedEntries] Scan failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * Generate unique ID


function findProjectByFileId(itemId) {
  if (!itemId || !Array.isArray(cachedProjects)) return null;
  return cachedProjects.find(project => project && project.fileId === itemId) || null;
}

function getProjectDisplayName(project) {
  if (!project || typeof project !== 'object') return '';
  const candidates = [
    project.displayName,
    project.projectFolderName,
    project.name,
    project.projectName,
    project.fileName
  ];
  for (const value of candidates) {
    if (!value) continue;
    const trimmed = String(value).trim();
    if (trimmed) return trimmed;
  }
  return '';
}

async function refreshWorkbookById(itemId, metadata = {}) {
  const result = {
    success: false,
    itemId,
    projectKey: null,
    projectName: null,
    trackedEntriesUpdated: false
  };

  if (!itemId) {
    result.error = 'missing_item_id';
    return result;
  }

  let project = null;
  let sessionId = null;
  let sessionDriveId = null;
  const previousProjectFile = currentProjectFile;

  try {
    project = findProjectByFileId(itemId);
    if (!project) {
      result.error = 'project_not_found';
      return result;
    }

    const projectKey = project.projectKey || getProjectCacheKey(project.projectFolderName || project.name || project.projectName || '');
    result.projectKey = projectKey;
    result.projectName = getProjectDisplayName(project) || projectKey;

    if (!project.driveId || !project.fileId) {
      const fallbackName = project.projectFolderName || project.name || project.projectName || result.projectName;
      const ensured = await ensureProjectInCache(fallbackName);
      if (ensured && ensured.project) {
        project = ensured.project;
      }
    }

    if (!project.driveId || !project.fileId) {
      result.error = 'missing_drive_or_file';
      return result;
    }

    sessionId = await createWorkbookSession(project.driveId, project.fileId, { persistChanges: false });
    sessionDriveId = project.driveId;
    currentWorkbookSession = { id: sessionId, driveId: project.driveId, itemId: project.fileId };
    currentProjectFile = {
      id: project.fileId,
      name: project.fileName || result.projectName,
      driveId: project.driveId,
      projectName: project.name || project.projectFolderName || result.projectName,
      projectFolderId: project.projectFolderId,
      stateFolderId: project.stateFolderId,
      site: project.site,
      state: project.state
    };

    const runtime = getOrCreateProjectRuntime(result.projectName);
    runtime.project = project;

    const hasMapping = project.divisionMapping && Object.keys(project.divisionMapping).length > 0;
    if (hasMapping) {
      divisionMapping = project.divisionMapping;
      runtime.divisionMapping = project.divisionMapping;
      runtime.divisions = project.divisions;
    } else {
      await buildDivisionMapping({ project, runtime, forceRefresh: true });
    }

    await loadDivisionContactsInternal({ project, runtime, forceRefresh: true });

    result.success = true;

    let fileMetadata = metadata && typeof metadata === 'object' ? { ...metadata } : {};
    if (!fileMetadata.eTag || !fileMetadata.lastModifiedDateTime || !fileMetadata.name) {
      try {
        const graphMetadata = await callGraphAPI(`/drives/${project.driveId}/items/${project.fileId}`);
        fileMetadata = { ...graphMetadata, ...fileMetadata };
      } catch (error) {
        console.warn('[Delta] Unable to load file metadata after refresh:', error);
      }
    }

    const resolvedEtag = fileMetadata.eTag || fileMetadata['@odata.etag'] || metadata.eTag || null;
    const resolvedModified = fileMetadata.lastModifiedDateTime || metadata.lastModifiedDateTime || new Date().toISOString();
    const resolvedName = fileMetadata.name || metadata.name || project.fileName || result.projectName;
    const resolvedPath = (fileMetadata.parentReference && fileMetadata.parentReference.path)
      || (metadata.parentReference && metadata.parentReference.path)
      || null;
    const resolvedWebUrl = fileMetadata.webUrl || metadata.webUrl || project.webUrl || null;

    const indexSnapshot = await deltaStore.getFileIndexSnapshot();
    indexSnapshot[itemId] = {
      id: itemId,
      name: resolvedName,
      driveId: project.driveId,
      folderId: project.projectFolderId || null,
      eTag: resolvedEtag,
      lastModifiedDateTime: resolvedModified,
      path: resolvedPath,
      webUrl: resolvedWebUrl,
      siteId: metadata.siteId || project.siteId || null,
      siteUrl: metadata.siteUrl || project.site || null,
      projectKey: result.projectKey
    };
    await deltaStore.saveFileIndexSnapshot(indexSnapshot);

    project.fileName = resolvedName;
    project.eTag = resolvedEtag;
    project.lastModifiedDateTime = resolvedModified;
    project.lastUpdated = new Date().toISOString();
    if (resolvedWebUrl) {
      project.webUrl = resolvedWebUrl;
    }

    try {
      await DB.saveProjects(cachedProjects);
    } catch (error) {
      console.error('[Delta] Failed to persist project cache after refresh:', error);
    }

  } catch (error) {
    result.error = error?.message || String(error);
    console.error('[Delta] Failed to refresh workbook', itemId, error);
  } finally {
    if (sessionId && sessionDriveId) {
      await closeWorkbookSession(sessionDriveId, itemId, sessionId);
    }
    currentWorkbookSession = null;
    currentProjectFile = previousProjectFile;
  }

  if (result.success && result.projectKey) {
    try {
      const trackedResult = await scanSelectedProjectsForTrackedEntries([result.projectKey]);
      result.trackedEntriesUpdated = !!(trackedResult && trackedResult.success !== false);
    } catch (error) {
      console.warn('[Delta] Failed to rebuild tracked entries for project', result.projectKey, error);
    }
  }

  return result;
}

async function handleDeletedWorkbook(itemId) {
  const result = {
    success: false,
    itemId,
    projectKey: null,
    projectName: null
  };

  try {
    const projectIndex = Array.isArray(cachedProjects)
      ? cachedProjects.findIndex(project => project && project.fileId === itemId)
      : -1;

    await deltaStore.removeFileEntry(itemId);

    if (projectIndex === -1) {
      result.error = 'project_not_found';
      return result;
    }

    const project = cachedProjects[projectIndex];
    const projectKey = project.projectKey || getProjectCacheKey(project.projectFolderName || project.name || project.projectName || '');
    result.projectKey = projectKey;
    result.projectName = getProjectDisplayName(project);

    cachedProjects.splice(projectIndex, 1);

    if (Array.isArray(cachedContacts) && cachedContacts.length) {
      cachedContacts = cachedContacts.filter(contact => {
        if (!contact) return false;
        const contactKey = contact.projectKey || getProjectCacheKey(contact.projectFolderName || contact.project || contact.projectName || '');
        return contactKey !== projectKey;
      });
    }

    if (Array.isArray(trackedProjectsConfig) && trackedProjectsConfig.length) {
      trackedProjectsConfig = trackedProjectsConfig.filter(entry => entry && entry.projectKey !== projectKey);
    }

    metadataCache = metadataCache || {};
    metadataCache.trackedProjects = trackedProjectsConfig;

    try {
      await DB.saveProjects(cachedProjects);
      await DB.saveContacts(cachedContacts);
      await DB.saveMetadata(metadataCache);
    } catch (error) {
      console.error('[Delta] Failed to persist cache after deletion:', error);
    }

    try {
      const storage = await chrome.storage.local.get('AB_TRACKED_ENTRIES');
      const trackedEntries = storage.AB_TRACKED_ENTRIES || {};
      if (projectKey && trackedEntries[projectKey]) {
        delete trackedEntries[projectKey];
        await chrome.storage.local.set({ AB_TRACKED_ENTRIES: trackedEntries });
      }
    } catch (error) {
      console.warn('[Delta] Failed to prune tracked entries cache:', error);
    }

    result.success = true;
    return result;
  } catch (error) {
    result.error = error?.message || String(error);
    console.error('[Delta] Error handling deleted workbook', itemId, error);
    return result;
  }
}

async function scanDelta(folderRef = {}) {
  if (!folderRef.driveId || !folderRef.folderId) {
    throw new Error('Missing driveId or folderId for delta scan');
  }

  const changedMap = new Map();
  const deletedSet = new Set();
  const indexSnapshot = await deltaStore.getFileIndexSnapshot();
  let indexDirty = false;

  let storedLink = await deltaStore.getDeltaLink(folderRef);
  let url = storedLink || `${GRAPH_API}/drives/${folderRef.driveId}/items/${folderRef.folderId}/delta`;

  while (url) {
    let response;
    try {
      response = await callGraphAPI(url);
    } catch (error) {
      if (storedLink && (error.status === 410 || error.status === 404)) {
        console.warn('[Delta] Stored delta link invalid. Restarting delta for folder', folderRef.folderId);
        await deltaStore.clearDeltaLink(folderRef);
        storedLink = null;
        url = `${GRAPH_API}/drives/${folderRef.driveId}/items/${folderRef.folderId}/delta`;
        continue;
      }
      throw error;
    }

    const items = Array.isArray(response.value) ? response.value : [];

    for (const item of items) {
      const itemId = item?.id;
      if (!itemId) continue;

      if (item.deleted) {
        deletedSet.add(itemId);
        if (indexSnapshot[itemId]) {
          delete indexSnapshot[itemId];
          indexDirty = true;
        }
        continue;
      }

      if (!item.file) continue;

      const nameLower = String(item.name || '').toLowerCase();
      if (!nameLower.endsWith('.xlsx') && !nameLower.endsWith('.xlsm')) continue;

      const existing = indexSnapshot[itemId];
      const eTag = item.eTag || item['@odata.etag'] || item['@microsoft.graph.fileSystemInfo']?.etag || null;
      const lastModified = item.lastModifiedDateTime || existing?.lastModifiedDateTime || null;

      if (existing) {
        let metadataUpdated = false;
        if (item.name && item.name !== existing.name) {
          existing.name = item.name;
          metadataUpdated = true;
        }
        const itemPath = item.parentReference && item.parentReference.path;
        if (itemPath && itemPath !== existing.path) {
          existing.path = itemPath;
          metadataUpdated = true;
        }
        if (item.webUrl && item.webUrl !== existing.webUrl) {
          existing.webUrl = item.webUrl;
          metadataUpdated = true;
        }
        if (lastModified && lastModified !== existing.lastModifiedDateTime) {
          existing.lastModifiedDateTime = lastModified;
          metadataUpdated = true;
        }
        if (metadataUpdated) {
          indexSnapshot[itemId] = { ...existing };
          indexDirty = true;
        }
      }

      const hasChanged = !existing || existing.eTag !== eTag;
      if (hasChanged) {
        changedMap.set(itemId, {
          id: itemId,
          driveId: folderRef.driveId,
          folderId: folderRef.folderId,
          siteId: folderRef.siteId || null,
          siteUrl: folderRef.siteUrl || null,
          name: item.name,
          eTag,
          lastModifiedDateTime: lastModified,
          parentReference: item.parentReference || null,
          webUrl: item.webUrl || null
        });
      }
    }

    if (response['@odata.nextLink']) {
      url = response['@odata.nextLink'];
    } else {
      const deltaLink = response['@odata.deltaLink'];
      if (deltaLink) {
        await deltaStore.setDeltaLink(folderRef, deltaLink);
      }
      url = null;
    }
  }

  if (indexDirty) {
    await deltaStore.saveFileIndexSnapshot(indexSnapshot);
  }

  return {
    changedItems: Array.from(changedMap.values()),
    deletedIds: Array.from(deletedSet)
  };
}

async function runDeltaScan(projectKeys = []) {
  const keysInput = Array.isArray(projectKeys) && projectKeys.length
    ? projectKeys
    : getTrackedProjectKeys();

  const normalizedKeys = Array.from(new Set((keysInput || []).map(key => getProjectCacheKey(key || '')))).filter(Boolean);

  if (!normalizedKeys.length) {
    return {
      success: false,
      error: 'no_projects_selected',
      changed: 0,
      refreshed: 0,
      deleted: 0,
      refreshedDetails: [],
      deletedDetails: [],
      errors: []
    };
  }

  const folderMap = new Map();

  for (const projectKey of normalizedKeys) {
    let project = resolveTrackedProjectContext(projectKey);
    if (!project) {
      project = Array.isArray(cachedProjects)
        ? cachedProjects.find(entry => getProjectCacheKey(entry.projectKey || entry.projectFolderName || entry.name || '') === projectKey)
        : null;
    }

    if (!project) {
      console.warn('[Delta] Project not found for key', projectKey);
      continue;
    }

    if (!project.driveId || !project.projectFolderId || !project.fileId) {
      const fallbackName = project.projectFolderName || project.name || project.projectName || projectKey;
      const ensured = await ensureProjectInCache(fallbackName);
      if (ensured && ensured.project) {
        project = ensured.project;
      }
    }

    if (!project.driveId || !project.projectFolderId || !project.fileId) {
      console.warn('[Delta] Missing folder reference for project', projectKey);
      continue;
    }

    const folderRef = {
      siteId: project.siteId || null,
      siteUrl: project.site || null,
      driveId: project.driveId,
      folderId: project.projectFolderId
    };
    const storageKey = buildDeltaStorageKey(folderRef);
    if (!folderMap.has(storageKey)) {
      folderMap.set(storageKey, { folderRef, projects: [] });
    }
    folderMap.get(storageKey).projects.push(project);
  }

  const refreshed = [];
  const deleted = [];
  const errors = [];
  let changedCount = 0;

  for (const { folderRef, projects } of folderMap.values()) {
    try {
      const deltaResult = await scanDelta(folderRef);
      const projectByFileId = new Map(projects.filter(Boolean).map(project => [project.fileId, project]));

      const filteredChanges = deltaResult.changedItems.filter(item => projectByFileId.has(item.id));
      const filteredDeletions = deltaResult.deletedIds.filter(id => projectByFileId.has(id));

      changedCount += filteredChanges.length;

      for (const change of filteredChanges) {
        const refreshResult = await refreshWorkbookById(change.id, { ...change, siteId: folderRef.siteId, siteUrl: folderRef.siteUrl });
        if (refreshResult.success) {
          refreshed.push({
            projectKey: refreshResult.projectKey,
            projectName: refreshResult.projectName,
            itemId: change.id,
            trackedEntriesUpdated: refreshResult.trackedEntriesUpdated
          });
        } else {
          errors.push({
            projectKey: refreshResult.projectKey || (projectByFileId.get(change.id)?.projectKey),
            itemId: change.id,
            reason: refreshResult.error || 'refresh_failed'
          });
        }
      }

      for (const deletedId of filteredDeletions) {
        const deleteResult = await handleDeletedWorkbook(deletedId);
        if (deleteResult.success) {
          deleted.push({
            projectKey: deleteResult.projectKey,
            projectName: deleteResult.projectName,
            itemId: deletedId
          });
        } else {
          errors.push({
            projectKey: deleteResult.projectKey,
            itemId: deletedId,
            reason: deleteResult.error || 'delete_failed'
          });
        }
      }

    } catch (error) {
      console.error('[Delta] Folder scan failed:', error);
      errors.push({
        folderId: folderRef.folderId,
        reason: error?.message || String(error)
      });
    }
  }

  return {
    success: true,
    changed: changedCount,
    refreshed: refreshed.length,
    deleted: deleted.length,
    refreshedDetails: refreshed,
    deletedDetails: deleted,
    errors
  };
}
const RESET_PROJECT_CACHE_TIMEOUT_MS = 10000;

function withTimeout(promise, ms, label = 'operation') {
  let timerId;
  const timeoutPromise = new Promise((_, reject) => {
    timerId = setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms);
  });
  return Promise.race([
    promise.finally(() => clearTimeout(timerId)),
    timeoutPromise
  ]);
}
async function resetTrackedEntries(projectKeys = []) {
  console.log('[Reset] resetTrackedEntries invoked with project keys:', projectKeys);
  const keysInput = Array.isArray(projectKeys) && projectKeys.length
    ? projectKeys
    : getTrackedProjectKeys();

  const normalizedMap = new Map();
  (keysInput || []).forEach(rawKey => {
    const normalized = getProjectCacheKey(rawKey || '');
    if (!normalized) return;
    if (!normalizedMap.has(normalized)) {
      normalizedMap.set(normalized, rawKey);
    }
  });

  const normalizedKeys = Array.from(normalizedMap.keys());
  console.log('[Reset] Normalized project keys:', normalizedKeys);

  if (!normalizedKeys.length) {
    return { success: false, error: 'no_projects_selected' };
  }

  const normalizedSet = new Set(normalizedKeys);
  const folderMap = new Map();
  const projectMap = new Map();
  const missingProjects = [];

  for (const projectKey of normalizedKeys) {
    console.log('[Reset] Processing project key:', projectKey);
    let project = resolveTrackedProjectContext(projectKey);
    if (!project) {
      project = Array.isArray(cachedProjects)
        ? cachedProjects.find(entry => getProjectCacheKey(entry.projectKey || entry.projectFolderName || entry.name || '') === projectKey)
        : null;
    }

    if (!project) {
      const missingKey = normalizedMap.get(projectKey) || projectKey;
      console.warn('[Reset] Project missing from cache, skipping:', missingKey);
      missingProjects.push(missingKey);
      continue;
    }

    if (!project.driveId || !project.projectFolderId || !project.fileId) {
      const fallbackName = project.projectFolderName || project.name || project.projectName || projectKey;
      try {
        console.log('[Reset] Ensuring project cache for:', fallbackName);
        const ensured = await withTimeout(
          ensureProjectInCache(fallbackName),
          RESET_PROJECT_CACHE_TIMEOUT_MS,
          '[Reset] ensureProjectInCache(' + fallbackName + ')'
        );
        if (ensured && ensured.project) {
          project = ensured.project;
          console.log('[Reset] Project cache refreshed for:', projectKey);
        }
      } catch (error) {
        console.warn('[Reset] ensureProjectInCache failed during reset for', projectKey, error);
      }
    }

    if (!project.driveId || !project.projectFolderId || !project.fileId) {
      const missingKey = normalizedMap.get(projectKey) || projectKey;
      console.warn('[Reset] Project missing drive metadata after ensure, skipping:', missingKey);
      missingProjects.push(missingKey);
      continue;
    }

    projectMap.set(projectKey, project);

    const folderRef = {
      siteId: project.siteId || null,
      siteUrl: project.site || null,
      driveId: project.driveId,
      folderId: project.projectFolderId
    };
    const storageKey = buildDeltaStorageKey(folderRef);
    if (!folderMap.has(storageKey)) {
      folderMap.set(storageKey, folderRef);
    }
  }

  if (!projectMap.size) {
    return { success: false, error: 'projects_not_found', missingProjects };
  }

  const storage = await chrome.storage.local.get('AB_TRACKED_ENTRIES');
  const trackedEntries = storage.AB_TRACKED_ENTRIES || {};
  let trackedEntriesRemoved = 0;
  for (const key of Object.keys(trackedEntries)) {
    const normalizedKey = getProjectCacheKey(key || '');
    if (normalizedSet.has(normalizedKey)) {
      delete trackedEntries[key];
      trackedEntriesRemoved += 1;
    }
  }
  await chrome.storage.local.set({ AB_TRACKED_ENTRIES: trackedEntries });

  const removedAuditEntries = await DB.deleteTrackedAuditForProjects(Array.from(projectMap.keys()));

  const fileIndexSnapshot = await deltaStore.getFileIndexSnapshot();
  let fileIndexRemoved = 0;
  for (const [itemId, metadata] of Object.entries(fileIndexSnapshot)) {
    const metaKey = getProjectCacheKey(metadata?.projectKey || '');
    if ((metaKey && normalizedSet.has(metaKey)) || (projectMap.has(getProjectCacheKey(metadata?.projectKey || '')))) {
      delete fileIndexSnapshot[itemId];
      fileIndexRemoved += 1;
    }
  }
  if (fileIndexRemoved > 0) {
    await deltaStore.saveFileIndexSnapshot(fileIndexSnapshot);
  }

  for (const folderRef of folderMap.values()) {
    await deltaStore.clearDeltaLink(folderRef);
  }

  for (const project of projectMap.values()) {
    await deltaStore.removeFileEntry(project.fileId);
  }

  let contactsRemoved = 0;
  if (Array.isArray(cachedContacts) && cachedContacts.length) {
    const filtered = cachedContacts.filter(contact => {
      const key = getProjectCacheKey(contact?.projectKey || contact?.projectFolderName || contact?.project || contact?.projectName || '');
      if (normalizedSet.has(key)) {
        contactsRemoved += 1;
        return false;
      }
      return true;
    });
    if (filtered.length !== cachedContacts.length) {
      cachedContacts = filtered;
      try {
        await DB.saveContacts(cachedContacts);
      } catch (error) {
        console.error('[Reset] Failed to persist contact cache during reset:', error);
      }
    }
  }

  let projectsTouched = 0;
  if (Array.isArray(cachedProjects) && cachedProjects.length) {
    cachedProjects = cachedProjects.map(project => {
      const key = getProjectCacheKey(project.projectKey || project.projectFolderName || project.name || project.projectName || '');
      if (normalizedSet.has(key)) {
        projectsTouched += 1;
        project.contactsByDivision = {};
        project.divisionMapping = project.divisionMapping || {};
        project.divisions = project.divisions || [];
        project.lastContactsSync = null;
        project.lastDivisionSync = null;
      }
      return project;
    });
    try {
      await DB.saveProjects(cachedProjects);
    } catch (error) {
      console.error('[Reset] Failed to persist project cache during reset:', error);
    }
  }

  Array.from(projectRuntimeCache.keys()).forEach(runtimeKey => {
    const runtime = projectRuntimeCache.get(runtimeKey);
    if (!runtime || !runtime.project) return;
    const runtimeProjectKey = getProjectCacheKey(runtime.project.projectKey || runtime.project.projectFolderName || runtime.project.name || '');
    if (normalizedSet.has(runtimeProjectKey)) {
      projectRuntimeCache.delete(runtimeKey);
    }
  });

  metadataCache = metadataCache || {};
  metadataCache.totalProjects = Array.isArray(cachedProjects) ? cachedProjects.length : 0;
  metadataCache.totalContacts = Array.isArray(cachedContacts) ? cachedContacts.length : 0;
  try {
    await DB.saveMetadata(metadataCache);
  } catch (error) {
    console.error('[Reset] Failed to persist metadata during reset:', error);
  }

  const resetSummary = {
    projectsProcessed: projectMap.size,
    projectsTouched,
    trackedEntriesRemoved,
    removedAuditEntries,
    contactsRemoved,
    fileIndexRemoved,
    deltaLinksCleared: folderMap.size,
    missingProjects
  };
  console.log('[Reset] Reset tracked entries summary:', resetSummary);

  return {
    success: true,
    ...resetSummary
  };
}
// Export the function to service worker global scope
if (typeof self !== 'undefined') {
  self.resetTrackedEntries = resetTrackedEntries;
  console.log('[Reset] resetTrackedEntries exported to self:', typeof self.resetTrackedEntries === 'function');
}
try { 
  if (typeof globalThis !== 'undefined') {
    globalThis.resetTrackedEntries = resetTrackedEntries;
  }
} catch (_) {}


/**
 * Generate unique ID for an entry based on contact info
 */
function generateEntryId(contact) {
  const str = `${contact.company}|${contact.email}|${contact.divisionKey}`;
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32-bit integer
  }
  return `entry_${Math.abs(hash)}`;
}

/**
 * Send progress update to UI
 */
function broadcastScanProgress(progressData) {
  try {
    chrome.runtime.sendMessage({
      type: 'TRACKED_SCAN_PROGRESS',
      data: progressData
    }).catch(err => {
      console.log('[TrackedEntries] Could not send progress (no receiver):', err.message);
    });
  } catch (error) {
    console.error('[TrackedEntries] Error broadcasting progress:', error);
  }
}

/**
 * Query SharePoint version history to find when a contact email was first added
 */
async function getContactAddedInfoSimplified(projectFile) {
  try {
    if (!projectFile || typeof projectFile !== 'object') {
      return null;
    }

    const addedDate = pickFirstValidDate(
      projectFile.lastModifiedDateTime,
      projectFile.fileModified,
      projectFile.modifiedTime,
      projectFile.lastUpdated,
      projectFile.lastRefreshed,
      projectFile.detectedAt
    );

    if (!addedDate) {
      return null;
    }

    let addedByEmail = '';
    let addedByDisplayName = '';

    if (projectFile.lastModifiedBy && typeof projectFile.lastModifiedBy === 'object') {
      addedByEmail = normalizeEmail(projectFile.lastModifiedBy.user?.email);
      addedByDisplayName = pickFirstNonEmpty(
        projectFile.lastModifiedBy.user?.displayName,
        projectFile.lastModifiedBy.application?.displayName
      ) || '';
    } else {
      const fallback = pickFirstNonEmpty(projectFile.lastModifiedBy, projectFile.modifiedBy);
      if (fallback) {
        if (fallback.includes('@')) {
          addedByEmail = normalizeEmail(fallback);
        } else {
          addedByDisplayName = fallback;
        }
      }
    }

    const resolvedName = resolveSharePointUserName(addedByEmail || addedByDisplayName || 'Unknown', {
      displayName: addedByDisplayName,
      email: addedByEmail
    });

    const dateObject = addedDate instanceof Date ? addedDate : new Date(addedDate);

    return {
      addedDate: dateObject.toISOString(),
      addedBy: addedByEmail || resolvedName,
      addedByDisplayName: resolvedName,
      addedByEmail: addedByEmail || null
    };
  } catch (error) {
    console.warn('[ContactAudit] Error getting simplified info:', error.message);
    return null;
  }
}



async function getContactAddedInfoFromVersionHistory(sheetName, email, projectFile, accessToken) {
  try {
    if (!email || !projectFile || !accessToken) {
      console.warn('[ContactAudit] Missing parameters for version history lookup');
      return null;
    }

    const versionsEndpoint = `/drives/${projectFile.driveId}/items/${projectFile.id}/versions`;
    let versionsData;
    try {
      versionsData = await callGraphAPI(versionsEndpoint);
    } catch (error) {
      console.warn('[ContactAudit] Failed to fetch versions:', error?.status || error);
      return null;
    }

    const versions = (versionsData?.value || []).sort((a, b) =>
      new Date(a.lastModifiedDateTime) - new Date(b.lastModifiedDateTime)
    );

    if (versions.length === 0) {
      return null;
    }

    console.log(`[ContactAudit] Checking ${versions.length} versions for email: ${email}`);

    let previousHadEmail = false;

    for (let i = 0; i < versions.length; i++) {
      const version = versions[i];
      const versionEmail = normalizeEmail(version?.lastModifiedBy?.user?.email);
      const versionDisplayName = pickFirstNonEmpty(
        version?.lastModifiedBy?.user?.displayName,
        version?.lastModifiedBy?.application?.displayName
      );

      let content;
      try {
        const contentEndpoint = `/drives/${projectFile.driveId}/items/${projectFile.id}/versions/${version.id}/content`;
        content = await callGraphAPI(contentEndpoint);
      } catch (error) {
        console.warn(`[ContactAudit] Failed to read version ${version.id} content:`, error?.status || error);
        continue;
      }

      const buffer = content instanceof ArrayBuffer ? content : null;
      if (!buffer) {
        console.warn('[ContactAudit] Version content not an ArrayBuffer for', version.id);
        continue;
      }

      let emailExists = false;
      try {
        const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        const sheetExists = workbook.SheetNames.includes(sheetName);
        if (!sheetExists) {
          previousHadEmail = false;
          continue;
        }

        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        for (const row of data) {
          for (const cell of row) {
            if (cell && String(cell).toLowerCase().trim() === email.toLowerCase().trim()) {
              emailExists = true;
              break;
            }
          }
          if (emailExists) break;
        }
      } catch (error) {
        console.warn('[ContactAudit] Error parsing version content', error);
        continue;
      }

      if (!previousHadEmail && emailExists) {
        const addedDate = version.lastModifiedDateTime;
        const resolvedName = resolveSharePointUserName(versionEmail || versionDisplayName || 'Unknown', {
          displayName: versionDisplayName,
          email: versionEmail
        });
        return {
          addedDate,
          addedBy: versionEmail || resolvedName,
          addedByDisplayName: resolvedName,
          addedByEmail: versionEmail || null
        };
      }

      previousHadEmail = emailExists;
    }

    if (previousHadEmail && versions.length > 0) {
      const oldestVersion = versions[0];
      const versionEmail = normalizeEmail(oldestVersion?.lastModifiedBy?.user?.email);
      const versionDisplayName = pickFirstNonEmpty(
        oldestVersion?.lastModifiedBy?.user?.displayName,
        oldestVersion?.lastModifiedBy?.application?.displayName
      );
      const resolvedName = resolveSharePointUserName(versionEmail || versionDisplayName || 'Unknown', {
        displayName: versionDisplayName,
        email: versionEmail
      });
      return {
        addedDate: oldestVersion.lastModifiedDateTime,
        addedBy: versionEmail || resolvedName,
        addedByDisplayName: resolvedName,
        addedByEmail: versionEmail || null
      };
    }

    return null;

  } catch (error) {
    console.error('[ContactAudit] Error querying version history:', error);
    return null;
  }
}


async function extractContactsWithVersionHistory(division, projectFile, options = {}) {
  const { accessToken, skipVersionLookup = false } = options;
  
  try {
    const contacts = await readDivisionContactsForMapping(division);
    
    if (contacts.length === 0 || !accessToken || skipVersionLookup) {
      return contacts;
    }

    console.log(`[ContactAudit] Enriching ${contacts.length} contacts with version history...`);
    
    const enrichedContacts = [];
    
    for (let i = 0; i < contacts.length; i++) {
      const contact = contacts[i];
      
      if (!contact.email) {
        enrichedContacts.push(contact);
        continue;
      }

      const versionInfo = await getContactAddedInfoFromVersionHistory(
        division.sheetName,
        contact.email,
        projectFile,
        accessToken
      );

      enrichedContacts.push({
        ...contact,
        addedDate: versionInfo?.addedDate || contact.lastSeen,
        addedBy: versionInfo?.addedBy || 'Unknown',
        versionHistoryChecked: !!versionInfo
      });

      await new Promise(resolve => setTimeout(resolve, 100));
    }

    console.log(`[ContactAudit] Enriched ${enrichedContacts.length} contacts`);
    return enrichedContacts;

  } catch (error) {
    console.error('[ContactAudit] Error extracting contacts with version history:', error);
    return await readDivisionContactsForMapping(division);
  }
}

/**
 * Get division mapping for a specific site
 */
function getAuditDivisionMappingForSite(site = 'AutoBuilders') {
  if (typeof getDivisionMappingForWorksheet === 'function') {
    return getDivisionMappingForWorksheet(site);
  }
  return {};
}

/**
 * Get all contact audit entries for display
 */
async function getContactAuditReport(filters = {}) {
  const { site = 'AutoBuilders', status = null } = filters;

  try {
    const db = await DB.init();
    const tx = db.transaction(['contactAudit'], 'readonly');
    const store = tx.objectStore('contactAudit');

    return new Promise((resolve, reject) => {
      const request = store.getAll();
      
      request.onsuccess = () => {
        let entries = request.result || [];

        if (site) {
          entries = entries.filter(e => e.site === site);
        }

        if (status) {
          entries = entries.filter(e => e.status === status);
        }

        entries.sort((a, b) => {
          if (a.divisionKey !== b.divisionKey) {
            return a.divisionKey.localeCompare(b.divisionKey);
          }
          return (a.company || '').localeCompare(b.company || '');
        });

        resolve(entries);
      };

      request.onerror = () => reject(request.error);
    });

  } catch (error) {
    console.error('[ContactAudit] Error retrieving audit report:', error);
    return [];
  }
}

console.log('[ContactAudit] Module loaded');



function flattenContactsForPanel(contactsByDivision = {}) {
  const flattened = [];
  if (!contactsByDivision) return flattened;

  for (const [divisionKey, contacts] of Object.entries(contactsByDivision)) {
    if (!Array.isArray(contacts)) continue;

    contacts.forEach((contact, index) => {
      const divisionName = contact.division || contact.sheetName || divisionKey;
      flattened.push({
        divisionKey,
        division: divisionName,
        company: contact.company || '',
        name: contact.name || contact.contactName || '',
        contactName: contact.contactName || contact.name || '',
        email: contact.email || '',
        phone: contact.phone || '',
        bidAmount: contact.bidAmount || '',
        columnLetter: contact.columnLetter,
        emailDomain: contact.emailDomain,
        phoneDigits: contact.phoneDigits,
        sheetName: contact.sheetName || divisionName,
        index
      });
    });
  }

  return flattened;
}

async function computeContactConflicts(projectName, contactData, divisionKeys = [], options = {}) {
  const { includeOtherDivisions = true, forceRefresh = false } = options;

  const workbook = await loadProjectWorkbook(projectName, {
    includeContacts: true,
    forceRefresh: forceRefresh && options.scope === 'divisions',
    forceContactRefresh: forceRefresh
  });

  if (!workbook || workbook.success === false || !workbook.project) {
    const errMsg = workbook?.error || 'Project not loaded';
    throw new Error(errMsg);
  }

  const runtime = getOrCreateProjectRuntime(workbook.project.name || projectName);
  const contactsByDivision = runtime.contactsByDivision || workbook.contactsByDivision || {};

  const targetKeys = new Set(divisionKeys);
  const companyNormalized = normalizeMatchValue(contactData.company);
  const phoneDigits = normalizePhoneDigits(contactData.phone);
  const emailLower = (contactData.email || '').toLowerCase();
  const emailDomain = extractEmailDomain(contactData.email);

  const divisionResults = {};
  const otherMatches = [];

  for (const [divisionKey, contacts] of Object.entries(contactsByDivision)) {
    const matches = [];
    for (const existing of contacts) {
      const reasons = [];
      if (emailLower && existing.email && existing.email.toLowerCase() === emailLower) {
        reasons.push('email');
      }
      if (phoneDigits && existing.phoneDigits && existing.phoneDigits === phoneDigits) {
        reasons.push('phone');
      }
      if (emailDomain && existing.emailDomain && existing.emailDomain === emailDomain) {
        reasons.push('domain');
      }
      if (companyNormalized && existing.companyNormalized) {
        if (existing.companyNormalized === companyNormalized
          || existing.companyNormalized.includes(companyNormalized)
          || companyNormalized.includes(existing.companyNormalized)) {
          reasons.push('company');
        }
      }

      if (reasons.length === 0) continue;

      const match = {
        divisionKey,
        divisionName: existing.sheetName,
        reasons,
        existingContact: existing
      };

      if (targetKeys.has(divisionKey)) {
        matches.push(match);
      } else if (includeOtherDivisions) {
        otherMatches.push(match);
      }
    }

    if (matches.length > 0) {
      divisionResults[divisionKey] = {
        conflicts: matches,
        existingContact: matches[0]?.existingContact,
        hasConflicts: true,
        conflictReasons: matches.flatMap(m => m.reasons)
      };
    }
  }

  const hasBlockingConflicts = Object.keys(divisionResults).length > 0;

  return {
    hasBlockingConflicts,
    divisionResults,
    otherMatches
  };
}

async function addBidderToProject(projectName, contactData, selectedDivisions, options = {}) {
  try {
    const divisionKeys = Array.isArray(selectedDivisions) ? selectedDivisions : [];

    if (divisionKeys.length === 0) {
      throw new Error('No divisions selected');
    }

    const conflictReport = await computeContactConflicts(projectName, contactData, divisionKeys, {
      includeOtherDivisions: true,
      forceRefresh: options.forceRefresh
    });

    if (conflictReport.hasBlockingConflicts && !options.overrideConflicts) {
      return {
        success: false,
        requiresOverride: true,
        conflicts: conflictReport
      };
    }

    await loadProjectWorkbook(projectName, { includeContacts: true });

    const results = [];

    for (const divisionKey of divisionKeys) {
      const division = divisionMapping[divisionKey];
      if (!division) {
        results.push({ division: divisionKey, success: false, error: 'Division not found' });
        continue;
      }

      const divisionConflict = conflictReport.divisionResults[divisionKey];
      const columnLetter = divisionConflict?.existingContact?.columnLetter;

      try {
        const result = await addContactToSheet(divisionKey, {
          ...contactData,
          bidAmount: contactData.bidAmount || ''
        }, { columnLetter });

        results.push({ division: divisionKey, success: true, column: result.column });
      } catch (error) {
        console.error('[Excel API] Failed to add bidder to', divisionKey, error);
        results.push({ division: divisionKey, success: false, error: error.message });
      }
    }

    const addedCount = results.filter(r => r.success).length;
    const failedCount = results.filter(r => !r.success).length;

    return {
      success: failedCount === 0,
      results,
      addedCount,
      failedCount,
      conflicts: conflictReport
    };

  } catch (error) {
    console.error('[Excel API] Error in addBidderToProject:', error);
    return { success: false, error: error.message };
  }
}

function sanitizePathComponent(value, fallback = 'Untitled') {
  if (!value || typeof value !== 'string') return fallback;
  const cleaned = value.replace(/[\/:*?"<>|#%]+/g, ' ').trim();
  return cleaned || fallback;
}

async function ensureFolderExists(driveId, parentId, folderName) {
  const safeName = sanitizePathComponent(folderName);
  const parent = parentId || 'root';
  const siblings = await getFolders(driveId, parent === 'root' ? '' : parent);
  const existing = siblings.find(folder => (folder.name || '').toLowerCase() === safeName.toLowerCase());
  if (existing) return existing;

  const createUrl = parent === 'root'
    ? `${GRAPH_API}/drives/${driveId}/root/children`
    : `${GRAPH_API}/drives/${driveId}/items/${parent}:/children`;
  const response = await fetch(createUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      name: safeName,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename'
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to create folder ${safeName}: ${errorText}`);
  }

  return await response.json();
}

async function ensureFolderPath(driveId, rootFolderId, segments = [], cacheContainer) {
  let currentId = rootFolderId || 'root';
  let lastFolder = null;

  for (const segment of segments) {
    if (!segment) continue;
    const safeName = sanitizePathComponent(segment);
    const cacheKey = `${currentId}:${safeName}`;

    if (cacheContainer && cacheContainer[cacheKey]) {
      lastFolder = cacheContainer[cacheKey];
      currentId = lastFolder.id;
      continue;
    }

    const folder = await ensureFolderExists(driveId, currentId === 'root' ? 'root' : currentId, safeName);
    lastFolder = folder;
    currentId = folder.id;

    if (cacheContainer) {
      cacheContainer[cacheKey] = { id: folder.id, name: folder.name };
    }
  }

  return lastFolder;
}

function base64ToUint8Array(base64) {
  const stripped = base64.replace(/^data:.*;base64,/, '');
  const binary = atob(stripped);
  const length = binary.length;
  const bytes = new Uint8Array(length);
  for (let i = 0; i < length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

async function uploadFileToFolder(driveId, parentId, fileName, contentBytes, contentType = 'application/octet-stream') {
  const safeName = sanitizePathComponent(fileName, 'Attachment');
  const uploadUrl = `${GRAPH_API}/drives/${driveId}/items/${parentId}:/${encodeURIComponent(safeName)}:/content?@microsoft.graph.conflictBehavior=replace`;

  const response = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': contentType
    },
    body: contentBytes
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to upload ${safeName}: ${errorText}`);
  }

  return await response.json();
}

function createEmailArchiveFile(emailMeta = {}, contactData = {}, divisionName = '') {
  const subject = sanitizePathComponent(emailMeta.subject || 'Email');
  const sender = emailMeta.from || contactData.email || '';
  const sentAt = emailMeta.sentAt ? new Date(emailMeta.sentAt) : new Date();
  const sentDisplay = sentAt.toLocaleString();
  const bodyHtml = emailMeta.bodyHtml || `<pre>${(emailMeta.bodyText || '').replace(/[<&>]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]))}</pre>`;

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${subject}</title></head><body>` +
    `<h2>${subject}</h2>` +
    `<p><strong>From:</strong> ${sender}</p>` +
    `<p><strong>Sent:</strong> ${sentDisplay}</p>` +
    (divisionName ? `<p><strong>Division:</strong> ${divisionName}</p>` : '') +
    `<hr/>${bodyHtml}</body></html>`;

  const encoder = new TextEncoder();
  const bytes = encoder.encode(html);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = `${subject}-${timestamp}.html`;

  return {
    name: fileName,
    bytes,
    contentType: 'text/html'
  };
}

async function ensureBidArtifacts(project, divisionKey, divisionName, companyName, attachments = [], emailMeta = {}) {
  if (!project || !project.driveId || !project.projectFolderId) {
    throw new Error('Project folder details unavailable for artifact upload');
  }

  project.folderCache = project.folderCache || {};

  const proposalsFolder = await ensureFolderPath(
    project.driveId,
    project.projectFolderId,
    ['Proposals'],
    project.folderCache
  );

  const divisionFolder = await ensureFolderPath(
    project.driveId,
    proposalsFolder?.id || project.projectFolderId,
    [divisionName || divisionKey],
    project.folderCache
  );

  const companyFolder = await ensureFolderPath(
    project.driveId,
    divisionFolder?.id,
    [companyName],
    project.folderCache
  );

  if (!companyFolder) {
    throw new Error('Unable to ensure company folder');
  }

  const uploaded = [];

  for (const attachment of attachments) {
    if (!attachment || attachment.isInline) {
      continue;
    }

    if (attachment.size && attachment.size < 20480 && attachment.contentType && attachment.contentType.startsWith('image/')) {
      continue;
    }

    if (!attachment.content) continue;

    const bytes = base64ToUint8Array(attachment.content);
    const contentType = attachment.contentType || 'application/octet-stream';
    const fileName = attachment.name || `Attachment-${Date.now()}`;

    const uploadedFile = await uploadFileToFolder(project.driveId, companyFolder.id, fileName, bytes, contentType);
    uploaded.push({ name: uploadedFile.name, id: uploadedFile.id });
  }

  if (emailMeta && (emailMeta.bodyHtml || emailMeta.bodyText)) {
    const emailFile = createEmailArchiveFile(emailMeta, {}, divisionName);
    const uploadedEmail = await uploadFileToFolder(project.driveId, companyFolder.id, emailFile.name, emailFile.bytes, emailFile.contentType);
    uploaded.push({ name: uploadedEmail.name, id: uploadedEmail.id });
  }

  try {
    await DB.saveProjects(cachedProjects);
  } catch (error) {
    console.error('[Bid] Failed to persist folder cache updates:', error);
  }

  return {
    success: true,
    folderId: companyFolder.id,
    files: uploaded
  };
}

async function submitBidWorkflow(message) {
  try {
    const { projectName, contactData = {}, divisions = [], attachments = [], emailMeta = {}, overrideConflicts = false } = message;

    if (!projectName) {
      throw new Error('Project name is required.');
    }

    if (!divisions.length) {
      throw new Error('Select at least one division before submitting a bid.');
    }

    const divisionKeys = divisions.map(d => d.key);
    const conflictReport = await computeContactConflicts(projectName, contactData, divisionKeys, {
      includeOtherDivisions: true,
      forceRefresh: message.forceRefresh
    });

    if (conflictReport.hasBlockingConflicts && !overrideConflicts) {
      return { success: false, requiresOverride: true, conflicts: conflictReport };
    }

    const workbook = await loadProjectWorkbook(projectName, {
      includeContacts: true,
      forceRefresh: message.forceRefresh,
      forceContactRefresh: message.forceContactRefresh
    });

    const results = [];

    for (const divisionSelection of divisions) {
      const divisionKey = divisionSelection.key;
      const division = divisionMapping[divisionKey];
      if (!division) {
        results.push({ divisionKey, success: false, error: 'Division not found in workbook' });
        continue;
      }

      const divisionConflict = conflictReport.divisionResults[divisionKey];
      const payload = {
        ...contactData,
        bidAmount: divisionSelection.bidAmount || '',
        projectName
      };

      try {
        const writeResult = await addContactToSheet(divisionKey, payload, {
          columnLetter: divisionConflict?.existingContact?.columnLetter
        });
        results.push({ divisionKey, success: true, column: writeResult.column });
      } catch (error) {
        console.error('[Bid] Failed to write contact for division', divisionKey, error);
        results.push({ divisionKey, success: false, error: error.message });
      }
    }

    const failed = results.filter(r => !r.success);
    const artifactResults = [];

    if (attachments.length > 0 || emailMeta?.bodyHtml || emailMeta?.bodyText) {
      for (const divisionSelection of divisions) {
        const divisionKey = divisionSelection.key;
        const divisionResult = results.find(r => r.divisionKey === divisionKey);
        if (!divisionResult || !divisionResult.success) continue;

        try {
          const artifact = await ensureBidArtifacts(
            workbook.project,
            divisionKey,
            divisionSelection.displayName || divisionSelection.label || divisionKey,
            contactData.company || 'Untitled Company',
            attachments,
            emailMeta
          );
          artifactResults.push({ divisionKey, ...artifact });
        } catch (error) {
          console.error('[Bid] Artifact upload failed for division', divisionKey, error);
          artifactResults.push({ divisionKey, success: false, error: error.message });
        }
      }
    }

    return {
      success: failed.length === 0,
      results,
      artifactResults,
      addedCount: results.filter(r => r.success).length,
      failedCount: failed.length,
      conflicts: conflictReport
    };

  } catch (error) {
    console.error('[Bid] Error submitting bid:', error);
    return { success: false, error: error.message };
  }
}

async function lookupZipLocation(zip) {
  if (!zip) {
    throw new Error('Zip code is required');
  }

  const normalized = String(zip).trim().slice(0, 5).padStart(5, '0');
  if (zipLocationCache[normalized]) {
    return { success: true, zip: normalized, location: zipLocationCache[normalized], cached: true };
  }

  const response = await fetch(`https://api.zippopotam.us/us/${normalized}`);
  if (!response.ok) {
    throw new Error(`Zip lookup failed (${response.status})`);
  }

  const data = await response.json();
  const place = (data.places || [])[0];
  if (!place) {
    throw new Error('No location data for provided zip');
  }

  const location = {
    lat: parseFloat(place.latitude),
    lon: parseFloat(place.longitude),
    city: place['place name'],
    state: data['state abbreviation'] || data.state
  };

  zipLocationCache[normalized] = location;
  metadataCache.zipCache = zipLocationCache;

  try {
    await DB.saveMetadata(metadataCache);
  } catch (error) {
    console.error('[Zip] Failed to persist zip cache:', error);
  }

  return { success: true, zip: normalized, location, cached: false };
}

const cacheDriveContext = {
  siteUrl: SHAREPOINT_SITES[0],
  siteId: null,
  driveId: null,
  folderId: null,
  folderPathSegments: [...CACHE_FOLDER_SEGMENTS],
  folderCache: {}
};

function buildDrivePathFromSegments(segments = []) {
  return segments.map(segment => encodeURIComponent(segment)).join('/');
}

async function ensureCacheFolderContext() {
  if (cacheDriveContext.driveId && cacheDriveContext.folderId) {
    return cacheDriveContext;
  }

  const site = await getSiteByUrl(cacheDriveContext.siteUrl);
  if (!site) {
    throw new Error('Cache site not found');
  }

  const drivesResponse = await getDrives(site.id);
  const drives = drivesResponse?.value || [];
  const sharedDocs = drives.find(d => {
    const name = (d.name || '').toLowerCase();
    return name === 'documents' || name === 'shared documents' || name.includes('shared');
  });

  if (!sharedDocs) {
    throw new Error('Shared Documents library not found for cache');
  }

  const folder = await ensureFolderPath(sharedDocs.id, 'root', CACHE_FOLDER_SEGMENTS, cacheDriveContext.folderCache);
  if (!folder) {
    throw new Error('Failed to ensure cache folder path');
  }

  cacheDriveContext.siteId = site.id;
  cacheDriveContext.driveId = sharedDocs.id;
  cacheDriveContext.folderId = folder.id;
  cacheDriveContext.folderPathSegments = [...CACHE_FOLDER_SEGMENTS];
  return cacheDriveContext;
}

async function uploadDriveFile(driveId, folderSegments, fileName, contentBytes, contentType = 'application/octet-stream') {
  const bytes = contentBytes instanceof Uint8Array ? contentBytes : new Uint8Array(contentBytes);
  const pathSegments = [...folderSegments, fileName];
  const encodedPath = buildDrivePathFromSegments(pathSegments);
  const totalSize = bytes.byteLength;

  if (totalSize <= 4 * 1024 * 1024) {
    const uploadUrl = `${GRAPH_API}/drives/${driveId}/root:/${encodedPath}:/content?@microsoft.graph.conflictBehavior=replace`;
    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': contentType
      },
      body: bytes
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to upload ${fileName}: ${errorText}`);
    }
    return await response.json();
  }

  const sessionUrl = `${GRAPH_API}/drives/${driveId}/root:/${encodedPath}:/createUploadSession`;
  const sessionResponse = await fetch(sessionUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      item: {
        '@microsoft.graph.conflictBehavior': 'replace',
        name: fileName
      }
    })
  });

  if (!sessionResponse.ok) {
    const errorText = await sessionResponse.text();
    throw new Error(`Failed to create upload session for ${fileName}: ${errorText}`);
  }

  const session = await sessionResponse.json();
  const uploadUrl = session.uploadUrl;
  const chunkSize = 5 * 1024 * 1024;
  const totalLength = bytes.byteLength;
  let position = 0;

  let finalMetadata = null;

  while (position < totalLength) {
    const nextPosition = Math.min(position + chunkSize, totalLength);
    const chunk = bytes.slice(position, nextPosition);
    const contentRange = `bytes ${position}-${nextPosition - 1}/${totalLength}`;

    const chunkResponse = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': `${chunk.byteLength}`,
        'Content-Range': contentRange
      },
      body: chunk
    });

    if (!chunkResponse.ok && chunkResponse.status !== 202) {
      const errorText = await chunkResponse.text();
      throw new Error(`Failed chunk upload (${contentRange}): ${errorText}`);
    }

    if (chunkResponse.status === 201 || chunkResponse.status === 200) {
      finalMetadata = await chunkResponse.json();
    }

    position = nextPosition;
  }

  return finalMetadata || { success: true };
}

async function downloadDriveFile(driveId, folderSegments, fileName) {
  const pathSegments = [...folderSegments, fileName];
  const encodedPath = buildDrivePathFromSegments(pathSegments);
  const url = `${GRAPH_API}/drives/${driveId}/root:/${encodedPath}:/content`;

  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`
    }
  });

  if (response.status === 404) {
    return null;
  }

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to download ${fileName}: ${errorText}`);
  }

  return await response.text();
}

async function ensureRemoteCacheSync(options = {}) {
  try {
    const context = await ensureCacheFolderContext();
    const manifestText = await downloadDriveFile(context.driveId, context.folderPathSegments, CACHE_MANIFEST_NAME);
    if (!manifestText) {
      console.log('[Cache] No remote manifest found');
      return { success: false, reason: 'missing' };
    }

    const manifest = JSON.parse(manifestText);
    const localUpdated = metadataCache.remoteManifest?.updatedAt;
    const remoteUpdated = manifest.updatedAt || null;

    const needsDownload = options.forceDownload || !localUpdated || !remoteUpdated ||
      new Date(remoteUpdated) > new Date(localUpdated);

    if (!needsDownload) {
      return { success: true, updated: false, manifest };
    }

    const downloadResult = await downloadCacheFromSharePoint(manifest);
    return { success: true, updated: true, manifest: downloadResult.manifest || manifest, projectsLoaded: downloadResult.projectsLoaded, contactsLoaded: downloadResult.contactsLoaded };
  } catch (error) {
    console.error('[Cache] Failed to sync remote cache:', error);
    return { success: false, error: error.message };
  }
}

async function uploadCacheToSharePoint() {
  const context = await ensureCacheFolderContext();
  const encoder = new TextEncoder();
  const timestamp = new Date().toISOString();

  const projectsContent = encoder.encode(JSON.stringify(cachedProjects));
  const contactsContent = encoder.encode(JSON.stringify(cachedContacts));
  const zipCacheContent = encoder.encode(JSON.stringify(zipLocationCache || {}));
  const trackedAuditEntries = await DB.getAllTrackedAuditEntries();
  const trackedProjectsList = getTrackedProjectsConfigSnapshot();
  const trackedAuditContent = encoder.encode(JSON.stringify(trackedAuditEntries));
  const trackedProjectsContent = encoder.encode(JSON.stringify(trackedProjectsList));

  await uploadDriveFile(context.driveId, context.folderPathSegments, 'projects.json', projectsContent, 'application/json');
  await uploadDriveFile(context.driveId, context.folderPathSegments, 'contacts.json', contactsContent, 'application/json');
  await uploadDriveFile(context.driveId, context.folderPathSegments, 'zip-cache.json', zipCacheContent, 'application/json');
  await uploadDriveFile(context.driveId, context.folderPathSegments, 'tracked-audit.json', trackedAuditContent, 'application/json');
  await uploadDriveFile(context.driveId, context.folderPathSegments, 'tracked-projects.json', trackedProjectsContent, 'application/json');

  const manifest = {
    version: '1.1',
    updatedAt: timestamp,
    totals: {
      projects: cachedProjects.length,
      contacts: cachedContacts.length,
      trackedAudit: trackedAuditEntries.length,
      trackedProjects: trackedProjectsList.length
    },
    files: [
      { name: 'projects.json', type: 'projects', size: projectsContent.length, updatedAt: timestamp },
      { name: 'contacts.json', type: 'contacts', size: contactsContent.length, updatedAt: timestamp },
      { name: 'zip-cache.json', type: 'zipCache', size: zipCacheContent.length, updatedAt: timestamp },
      { name: 'tracked-audit.json', type: 'trackedAudit', size: trackedAuditContent.length, updatedAt: timestamp },
      { name: 'tracked-projects.json', type: 'trackedProjects', size: trackedProjectsContent.length, updatedAt: timestamp }
    ]
  };

  const manifestContent = encoder.encode(JSON.stringify(manifest, null, 2));
  await uploadDriveFile(context.driveId, context.folderPathSegments, CACHE_MANIFEST_NAME, manifestContent, 'application/json');

  metadataCache.remoteManifest = manifest;
  metadataCache.zipCache = zipLocationCache;
  metadataCache.trackedProjects = trackedProjectsConfig;
  try {
    await DB.saveMetadata(metadataCache);
  } catch (error) {
    console.error('[Cache] Failed to persist remote manifest metadata:', error);
  }

  return { success: true, manifest };
}

async function downloadCacheFromSharePoint(existingManifest = null) {
  const context = await ensureCacheFolderContext();
  let manifest = existingManifest;

  if (!manifest) {
    const manifestText = await downloadDriveFile(context.driveId, context.folderPathSegments, CACHE_MANIFEST_NAME);
    if (!manifestText) {
      throw new Error('SharePoint cache manifest not found.');
    }
    manifest = JSON.parse(manifestText);
  }

  const projectsFile = manifest.files?.find(f => f.type === 'projects')?.name || 'projects.json';
  const contactsFile = manifest.files?.find(f => f.type === 'contacts')?.name || 'contacts.json';
  const zipCacheFile = manifest.files?.find(f => f.type === 'zipCache')?.name || 'zip-cache.json';
  const trackedAuditFile = manifest.files?.find(f => f.type === 'trackedAudit')?.name || 'tracked-audit.json';
  const trackedProjectsFile = manifest.files?.find(f => f.type === 'trackedProjects')?.name || 'tracked-projects.json';

  const projectsText = await downloadDriveFile(context.driveId, context.folderPathSegments, projectsFile);
  const contactsText = await downloadDriveFile(context.driveId, context.folderPathSegments, contactsFile);
  const zipCacheText = zipCacheFile ? await downloadDriveFile(context.driveId, context.folderPathSegments, zipCacheFile) : null;

  let trackedAuditText = null;
  let trackedProjectsText = null;

  try {
    trackedAuditText = await downloadDriveFile(context.driveId, context.folderPathSegments, trackedAuditFile);
  } catch (error) {
    console.warn('[Cache] Failed to download tracked audit history:', error);
    trackedAuditText = null;
  }

  try {
    trackedProjectsText = await downloadDriveFile(context.driveId, context.folderPathSegments, trackedProjectsFile);
  } catch (error) {
    console.warn('[Cache] Failed to download tracked project list:', error);
    trackedProjectsText = null;
  }

  if (!projectsText || !contactsText) {
    throw new Error('Failed to download cache files from SharePoint');
  }

  cachedProjects = JSON.parse(projectsText).map(normalizeProjectRecord);
  cachedContacts = JSON.parse(contactsText);

  enforceProjectRetentionLimit();

  if (zipCacheText) {
    try {
      zipLocationCache = JSON.parse(zipCacheText) || {};
    } catch (error) {
      console.warn('[Cache] Failed to parse remote zip cache:', error);
      zipLocationCache = {};
    }

  } else {
    zipLocationCache = {};
  }
  zipLocationCache = {
    ...staticZipLocationCache,
    ...zipLocationCache
  };



  let trackedAuditEntries = [];
  if (trackedAuditText) {
    try {
      trackedAuditEntries = JSON.parse(trackedAuditText) || [];
    } catch (error) {
      console.warn('[Cache] Failed to parse tracked audit history:', error);
      trackedAuditEntries = [];
    }
  }

  try {
    await DB.replaceTrackedAuditEntries(Array.isArray(trackedAuditEntries) ? trackedAuditEntries : []);
  } catch (error) {
    console.error('[Cache] Failed to persist tracked audit history:', error);
  }

  let remoteTrackedProjects = trackedProjectsText ? (() => {
    try {
      const parsed = JSON.parse(trackedProjectsText);
      return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      console.warn('[Cache] Failed to parse tracked project list:', error);
      return [];
    }
  })() : trackedProjectsConfig;

  if (!Array.isArray(remoteTrackedProjects)) {
    remoteTrackedProjects = [];
  }

  trackedProjectsConfig = remoteTrackedProjects;
  metadataCache.trackedProjects = trackedProjectsConfig;
  metadataCache.zipCache = zipLocationCache;
  metadataCache.remoteManifest = manifest;

  await saveCacheToIndexedDB();

  return {
    success: true,
    manifest,
    projectsLoaded: cachedProjects.length,
    contactsLoaded: cachedContacts.length,
    trackedAuditLoaded: Array.isArray(trackedAuditEntries) ? trackedAuditEntries.length : 0,
    trackedProjectsLoaded: trackedProjectsConfig.length
  };
}

function buildContactAuditSummary(entries = []) {
  const userMap = new Map();
  const projectMap = new Map();

  entries.forEach(entry => {
    if (!entry) return;
    const userKey = (entry.lastModifiedBy || 'Unknown').trim() || 'Unknown';
    const projectKey = entry.projectKey || entry.projectName || 'unknown-project';
    const projectName = entry.projectName || entry.projectKey || 'Unknown Project';
    const divisionName = entry.divisionName || entry.divisionKey || 'Unknown Division';

    if (!userMap.has(userKey)) {
      userMap.set(userKey, { user: userKey, total: 0 });

    }
  });

  const totalsByUser = Array.from(userMap.values()).sort((a, b) => b.total - a.total);
  const projects = Array.from(projectMap.values()).map(projectData => ({
    user: projectData.user,
    total: projectData.total
  })).sort((a, b) => b.total - a.total);

  return {
    totalEntries: entries.length,
    totalsByUser,
    projects
  };
}

function buildTrackedAuditSummary(entries = []) {
  const dailyTotals = { today: 0, yesterday: 0, twoDaysAgo: 0 };
  const projectMap = new Map();
  const divisionMap = new Map();

  const now = new Date();
  const startToday = new Date(now);
  startToday.setHours(0, 0, 0, 0);
  const startYesterday = new Date(startToday);
  startYesterday.setDate(startYesterday.getDate() - 1);
  const startTwoDaysAgo = new Date(startYesterday);
  startTwoDaysAgo.setDate(startTwoDaysAgo.getDate() - 1);

  entries.forEach(entry => {
    if (!entry) return;

    let detectedAt = entry.detectedAt ? new Date(entry.detectedAt) : null;
    if (!detectedAt || Number.isNaN(detectedAt.getTime())) {
      detectedAt = new Date();
    }

    if (detectedAt >= startToday) {
      dailyTotals.today += 1;
    } else if (detectedAt >= startYesterday) {
      dailyTotals.yesterday += 1;
    } else if (detectedAt >= startTwoDaysAgo) {
      dailyTotals.twoDaysAgo += 1;
    }

    const projectKey = entry.projectKey || entry.projectName || 'unknown-project';
    const projectName = entry.projectName || entry.projectDisplayName || entry.projectKey || 'Unknown Project';
    const divisionName = entry.divisionName || entry.divisionKey || 'Unknown Division';
    const companyName = entry.company || entry.companyName || 'Unknown Company';

    if (!projectMap.has(projectKey)) {
      projectMap.set(projectKey, {
        project: projectName,
        projectKey,
        total: 0,
        divisions: new Map(),
        companies: new Map()
      });
    }

    const projectData = projectMap.get(projectKey);
    projectData.total += 1;

    if (!projectData.divisions.has(divisionName)) {
      projectData.divisions.set(divisionName, { division: divisionName, total: 0, companies: new Map() });
    }
    const divisionData = projectData.divisions.get(divisionName);
    divisionData.total += 1;

    divisionData.companies.set(companyName, (divisionData.companies.get(companyName) || 0) + 1);
    projectData.companies.set(companyName, (projectData.companies.get(companyName) || 0) + 1);

    if (!divisionMap.has(divisionName)) {
      divisionMap.set(divisionName, { division: divisionName, total: 0 });
    }
    divisionMap.get(divisionName).total += 1;
  });

  const projects = Array.from(projectMap.values()).map(projectData => ({
    project: projectData.project,
    projectKey: projectData.projectKey,
    total: projectData.total,
    companies: Array.from(projectData.companies.entries()).map(([company, total]) => ({ company, total })).sort((a, b) => b.total - a.total),
    divisions: Array.from(projectData.divisions.values()).map(divisionData => ({
      division: divisionData.division,
      total: divisionData.total,
      companies: Array.from(divisionData.companies.entries()).map(([company, total]) => ({ company, total })).sort((a, b) => b.total - a.total)
    })).sort((a, b) => b.total - a.total)
  })).sort((a, b) => b.total - a.total);

  const divisions = Array.from(divisionMap.values()).sort((a, b) => b.total - a.total);

  return {
    totalEntries: entries.length,
    dailyTotals,
    projects,
    divisions
  };
}

async function getTrackedAuditReport(options = {}) {
  const filters = { ...options };
  if (!filters.limit || !Number.isFinite(filters.limit)) {
    filters.limit = 2000;
  } else {
    filters.limit = Math.min(filters.limit, 5000);
  }

  const entries = await DB.getTrackedAuditEntries(filters);
  const summary = buildTrackedAuditSummary(entries);
  summary.range = { start: filters.start || null, end: filters.end || null };
  summary.generatedAt = new Date().toISOString();
  summary.projectKeys = getTrackedProjectKeys();
  return { entries, summary };
}

globalThis.getTrackedAuditReport = getTrackedAuditReport;
self.getTrackedAuditReport = getTrackedAuditReport;
console.log('[Audit] getTrackedAuditReport helper ready');
console.log('[Audit] typeof getTrackedAuditReport:', typeof globalThis.getTrackedAuditReport);


function getDateRangeForDate(date = new Date()) {
  const start = new Date(date);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 1);
  return { start: start.toISOString(), end: end.toISOString() };
}

function computeNextAuditRunTime() {
  const now = new Date();
  const target = new Date(now);
  target.setHours(AUDIT_DAILY_RUN_HOUR, AUDIT_DAILY_RUN_MINUTE, 0, 0);
  if (target <= now) {
    target.setDate(target.getDate() + 1);
  }
  return target.getTime();
}

function scheduleDailyAuditRun() {
  try {
    const when = computeNextAuditRunTime();
    chrome.alarms.create(AUDIT_DAILY_ALARM_NAME, { when, periodInMinutes: 24 * 60 });
    console.log('[Audit] Daily audit alarm scheduled for', new Date(when).toISOString());
  } catch (error) {
    console.error('[Audit] Failed to schedule daily audit alarm:', error);
  }
}

async function runDailyAuditSnapshot() {
  try {
    const { start, end } = getDateRangeForDate(new Date());
    const report = (await getContactAuditReport({ start, end, limit: MAX_AUDIT_FETCH_LIMIT })) || {};
    const summary = report.summary || {};
    const normalizedSummary = {
      generatedAt: summary.generatedAt || new Date().toISOString(),
      totalEntries: typeof summary.totalEntries === 'number' ? summary.totalEntries : 0,
      totalsByUser: summary.totalsByUser || {},
      projects: Array.isArray(summary.projects) ? summary.projects : []
    };
    metadataCache.lastAuditSummary = {
      start,
      end,
      generatedAt: normalizedSummary.generatedAt,
      totalEntries: normalizedSummary.totalEntries,
      totalsByUser: normalizedSummary.totalsByUser,
      projects: normalizedSummary.projects
    };
    if (!report.summary) {
      report.summary = normalizedSummary;
    }
    try {
      await DB.saveMetadata(metadataCache);
    } catch (error) {
      console.warn('[Audit] Failed to persist audit summary metadata:', error);
    }
    return report;
  } catch (error) {
    console.error('[Audit] Daily audit snapshot failed:', error);
    return { success: false, summary: null, error: error?.message || String(error) };
  }
}


async function rebuildTrackedAuditEntriesForProjects(projectConfigs = []) {
  const configs = Array.isArray(projectConfigs) ? projectConfigs : [];
  if (!configs.length) {
    await DB.replaceTrackedAuditEntries([]);
    return { success: true, entries: 0, projects: 0 };
  }

  const rebuildTimestamp = new Date().toISOString();
  const entries = [];
  const seenProjectKeys = new Set();

  const resolveCachedProject = (projectKey) => {
    if (!projectKey) return null;
    const normalizedKey = projectKey;
    if (!Array.isArray(cachedProjects)) return null;
    return cachedProjects.find(project => {
      if (!project) return false;
      const keyCandidates = [
        project.projectKey,
        project.projectFolderName,
        project.displayName,
        project.name,
        project.projectName
      ];
      return keyCandidates.some(candidate => {
        if (!candidate) return false;
        const key = getProjectCacheKey(candidate);
        return key && key === normalizedKey;
      });
    }) || null;
  };

  const buildWorkbookKey = ({ projectKey, workbookId, workbookPath, workbookName, emailNormalized }) => {
    const candidates = [
      workbookId,
      workbookPath,
      workbookName && projectKey ? `${projectKey}::${workbookName}` : null,
      projectKey,
      emailNormalized ? `${projectKey}::${emailNormalized}` : null
    ];
    for (const candidate of candidates) {
      if (candidate === null || candidate === undefined) continue;
      const trimmed = String(candidate).trim();
      if (trimmed) return trimmed.toLowerCase();
    }
    return projectKey || emailNormalized || '';
  };

  configs.forEach(config => {
    if (!config || typeof config !== 'object') return;
    const projectKey = config.projectKey;
    if (!projectKey || seenProjectKeys.has(projectKey)) return;
    seenProjectKeys.add(projectKey);

    const cachedProject = resolveCachedProject(projectKey);
    if (!cachedProject) {
      console.warn('[Audit] Tracked project not found in cache for rebuild:', projectKey);
      return;
    }

    const contactsByDivision = cachedProject.contactsByDivision || {};
    if (!contactsByDivision || typeof contactsByDivision !== 'object') {
      console.warn('[Audit] Cached project missing contacts for rebuild:', projectKey);
      return;
    }

    const projectName = cachedProject.displayName || cachedProject.projectFolderName || cachedProject.projectName || config.projectName || projectKey;
    const projectNumber = cachedProject.projectNumber || config.projectNumber || null;
    const driveIdDefault = config.driveId || cachedProject.driveId || null;
    const siteUrl = config.siteUrl || cachedProject.site || cachedProject.siteUrl || cachedProject.siteURL || null;
    const stateName = config.stateName || cachedProject.stateName || cachedProject.state || null;

    const latestFile = cachedProject.latestFile || {};
    const workbookIdDefault = latestFile.id || config.workbookId || cachedProject.workbookId || null;
    const workbookNameDefault = latestFile.name || config.workbookName || cachedProject.fileName || cachedProject.projectFolderName || projectName;
    const workbookPathDefault = (latestFile.parentReference && (latestFile.parentReference.path || (latestFile.parentReference.sharepointIds && latestFile.parentReference.sharepointIds.siteUrl)))
      || latestFile.webUrl
      || config.workbookPath
      || cachedProject.workbookPath
      || null;
    const driveId = driveIdDefault || (latestFile.parentReference && latestFile.parentReference.driveId) || null;

    const resolveLastModifiedBy = () => {
      const fileUser = latestFile && latestFile.lastModifiedBy && latestFile.lastModifiedBy.user && (latestFile.lastModifiedBy.user.displayName || latestFile.lastModifiedBy.user.email);
      const appName = latestFile && latestFile.lastModifiedBy && latestFile.lastModifiedBy.application && latestFile.lastModifiedBy.application.displayName;
      return config.lastModifiedBy || fileUser || appName || cachedProject.lastModifiedBy || 'Unknown';
    };

    const resolveLastModifiedDateTime = () => {
      return (latestFile && latestFile.lastModifiedDateTime)
        || cachedProject.fileModified
        || cachedProject.lastUpdated
        || config.lastModifiedDateTime
        || null;
    };

    const lastModifiedByFallback = resolveLastModifiedBy();
    const lastModifiedDateTimeFallback = resolveLastModifiedDateTime();

    Object.entries(contactsByDivision).forEach(([divisionKey, contactList]) => {
      if (!Array.isArray(contactList) || contactList.length === 0) return;
      const deduped = dedupeContacts(contactList);
      deduped.forEach(contact => {
        if (!contact) return;
        const email = String(contact.email || '').trim();
        if (!email) return;
        const emailNormalized = email.toLowerCase();
        const divisionName = contact.division || contact.divisionName || divisionKey || 'Unknown Division';
        const contactState = contact.state || stateName || null;
        const contactSource = contact.source || 'rebuild';
        const detectedAt = (() => {
          const candidate = contact.detectedAt || contact.lastModifiedDateTime || lastModifiedDateTimeFallback;
          if (!candidate) return rebuildTimestamp;
          try {
            return new Date(candidate).toISOString();
          } catch (_) {
            return rebuildTimestamp;
          }
        })();
        const lastModifiedBy = contact.lastModifiedBy || lastModifiedByFallback || 'Unknown';
        let lastModifiedDateTime = contact.lastModifiedDateTime || lastModifiedDateTimeFallback || null;
        if (lastModifiedDateTime) {
          try {
            lastModifiedDateTime = new Date(lastModifiedDateTime).toISOString();
          } catch (_) {
            lastModifiedDateTime = String(lastModifiedDateTime).trim() || null;
          }
        }

        const workbookId = contact.workbookId || workbookIdDefault;
        const workbookName = contact.workbookName || contactSource || workbookNameDefault;
        const workbookPath = contact.workbookPath || workbookPathDefault;

        const workbookKey = buildWorkbookKey({
          projectKey,
          workbookId,
          workbookPath,
          workbookName,
          emailNormalized
        });

        entries.push({
          projectKey,
          projectName,
          projectNumber,
          workbookId,
          workbookName,
          workbookPath,
          workbookKey,
          driveId,
          siteUrl,
          state: contactState,
          divisionKey,
          divisionName,
          email,
          emailNormalized,
          company: contact.company || '',
          contactName: contact.name || contact.contactName || '',
          phone: contact.phone || '',
          detectedAt,
          lastModifiedBy,
          lastModifiedDateTime,
          source: contactSource,
          trackedAt: rebuildTimestamp,
          projectState: contactState
        });
      });
    });
  });

  if (!entries.length) {
    await DB.replaceTrackedAuditEntries([]);
    return { success: true, entries: 0, projects: seenProjectKeys.size };
  }

  try {
    await DB.replaceTrackedAuditEntries(entries);
    console.log('[Audit] Rebuilt tracked audit history for', seenProjectKeys.size, 'projects with', entries.length, 'entries');
    return { success: true, entries: entries.length, projects: seenProjectKeys.size };
  } catch (error) {
    console.error('[Audit] Failed to rebuild tracked audit history:', error);
    return { success: false, error: error?.message || String(error), entries: entries.length, projects: seenProjectKeys.size };
  }
}

async function rescanTrackedProjects(options = {}) {
  const { upload = true, triggeredBy = 'manual' } = options || {};
  const trackedProjects = getTrackedProjectsConfigSnapshot();

  if (!trackedProjects.length) {
    return { success: false, reason: 'no_tracked_projects' };
  }

  if (globalScanInProgress) {
    return { success: false, reason: 'in_progress' };
  }

  const groupMap = new Map();
  trackedProjects.forEach(project => {
    if (!project || !project.siteUrl || !project.stateName) {
      return;
    }
    const siteUrl = String(project.siteUrl).trim();
    const stateName = String(project.stateName).trim();
    if (!siteUrl || !stateName) return;
    const key = `${siteUrl}::${stateName}`;
    if (!groupMap.has(key)) {
      groupMap.set(key, { siteUrl, stateName, projects: [] });
    }
    groupMap.get(key).projects.push(project);
  });

  if (groupMap.size === 0) {
    return { success: false, reason: 'incomplete_project_context' };
  }

  if (!ensureAccessTokenValidNoPrompt()) {
    return { success: false, reason: 'no_token' };
  }

  globalScanInProgress = true;
  const summary = {
    success: true,
    triggeredBy,
    states: [],
    failures: [],
    projectsTracked: trackedProjects.length
  };

  try {
    for (const group of groupMap.values()) {
      try {
        const result = await scanSingleState(group.siteUrl, group.stateName, null, group.projects);
        summary.states.push({ group, result });
        if (!result || result.success === false) {
          summary.failures.push({ siteUrl: group.siteUrl, stateName: group.stateName, error: result?.error || 'unknown' });
        }
      } catch (error) {
        summary.success = false;
        summary.failures.push({ siteUrl: group.siteUrl, stateName: group.stateName, error: error?.message || String(error) });
      }
    }
  } finally {
    globalScanInProgress = false;
  }

  if (upload) {
    try {
      const uploadResult = await uploadCacheToSharePoint();
      summary.cacheUpload = uploadResult;
    } catch (error) {
      summary.cacheUpload = { success: false, error: error?.message || String(error) };
    }
  }

  try {
    const rebuildResult = await rebuildTrackedAuditEntriesForProjects(trackedProjects);
    summary.trackedAuditRebuild = rebuildResult;
  } catch (error) {
    console.error('[Audit] Post-rescan tracked audit rebuild failed:', error);
    summary.trackedAuditRebuild = { success: false, error: error?.message || String(error) };
  }

  summary.success = summary.failures.length === 0;
  summary.completedAt = new Date().toISOString();

  return summary;
}

globalThis.rescanTrackedProjects = rescanTrackedProjects;
self.rescanTrackedProjects = rescanTrackedProjects;
console.log('[Audit] rescanTrackedProjects helper ready');
console.log('[Audit] typeof rescanTrackedProjects:', typeof globalThis.rescanTrackedProjects);

function ensureAccessTokenValidNoPrompt() {
  if (!accessToken || !tokenExpiry) return true;
  // tokenExpiry is stored as a timestamp (number in milliseconds), not a Date object
  return tokenExpiry > Date.now() + 60 * 1000;
}
initializeStaticProjectCache();

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('Background received:', message.type);
  if (typeof message?.type === 'string' && (message.type.startsWith('AB_CAMPAIGN') || message.type.startsWith('campaign:') || message.type.startsWith('AB_CMP'))) {
    console.log('[Background] Delegating campaign message to campaigns worker');
    return false;
  }

  switch (message.type) {
    case 'AB_EXCEL_SUMMARY': {
      // Handled by excel-summary-bg.js listener
      return false;
    }
    case 'AB_PANEL_DETACH_REQUEST': {
      const tabId = sender.tab?.id || message.tabId || detachedPanelTabId;
      if (!tabId) {
        if (typeof sendResponse === 'function') { sendResponse({ success: false, error: 'No host tab available' }); }
        return true;
      }
      detachedPanelTabId = tabId;
      const focusPopup = () => {
        if (detachedPanelWindowId) {
          chrome.windows.update(detachedPanelWindowId, { focused: true }).catch(() => {});
        }
      };
      const createPopup = () => {
        const popupUrl = chrome.runtime.getURL('panel.html?mode=popup');
        chrome.windows.create({ url: popupUrl, type: 'popup', width: 520, height: 760, focused: true }, windowInfo => {
          detachedPanelWindowId = windowInfo?.id || null;
          chrome.tabs.sendMessage(tabId, { type: 'AB_PANEL_DETACH_CONFIRMED' }).catch(() => {});
        });
      };
      if (detachedPanelWindowId) {
        focusPopup();
        chrome.tabs.sendMessage(tabId, { type: 'AB_PANEL_DETACH_CONFIRMED' }).catch(() => {});
      } else {
        createPopup();
      }
      if (typeof sendResponse === 'function') { sendResponse({ success: true }); }
      return true;
    }
    case 'AB_PANEL_ATTACH_REQUEST': {
      const targetTabId = message.tabId || sender.tab?.id || detachedPanelTabId;
      if (detachedPanelWindowId) {
        suppressNextAttach = true;
        chrome.windows.remove(detachedPanelWindowId).catch(() => {});
      }
      if (targetTabId) {
        detachedPanelTabId = targetTabId;
        chrome.tabs.sendMessage(targetTabId, { type: 'AB_PANEL_ATTACH' }).catch(() => {});
      }
      if (typeof sendResponse === 'function') { sendResponse({ success: true }); }
      return true;
    }
    case 'AB_PANEL_FOCUS_POPUP': {
      if (detachedPanelWindowId) {
        chrome.windows.update(detachedPanelWindowId, { focused: true }).catch(() => {});
        if (typeof sendResponse === 'function') { sendResponse({ success: true }); }
      } else if (detachedPanelTabId) {
        chrome.tabs.sendMessage(detachedPanelTabId, { type: 'AB_PANEL_ATTACH' }).catch(() => {});
        if (typeof sendResponse === 'function') { sendResponse({ success: false, error: 'Popup not available' }); }
      } else {
        if (typeof sendResponse === 'function') { sendResponse({ success: false, error: 'Popup not available' }); }
      }
      return true;
    }
    case 'AB_AUTHENTICATE':
      authenticateWithMicrosoft()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_SCAN_STATE':
      const progressCallback = (progress) => {
        if (sender.tab) {
          chrome.tabs.sendMessage(sender.tab.id, {
            type: 'AB_SCAN_PROGRESS',
            progress: progress
          }).catch(() => {});
        } else {
          chrome.runtime.sendMessage({
            type: 'AB_SCAN_PROGRESS',
            progress: progress
          }).catch(() => {});
        }
      };
      
      scanSingleState(message.siteUrl, message.stateName, progressCallback)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_GET_STATES':
      getAvailableStates()
        .then(states => sendResponse({ success: true, states }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_GET_CACHED_DATA': {
      (async () => {
        let snapshot = buildBundledCacheSnapshot();
        if (!snapshot.projects.length && !snapshot.contacts.length) {
          const hydrateResult = await hydrateCacheFromRemote({ triggeredBy: 'ab_get_cached_data', silent: true });
          if (hydrateResult && hydrateResult.success) {
            snapshot = buildBundledCacheSnapshot();
          }
        }
        sendResponse({
          success: true,
          projects: snapshot.projects,
          contacts: snapshot.contacts,
          lastScanTime: lastScanTime,
          zipCache: snapshot.zipCache,
          trackedProjects: getTrackedProjectsConfigSnapshot(),
          staticProjectsCount: snapshot.staticProjectsCount,
          dynamicProjectsCount: snapshot.dynamicProjectsCount,
          staticContactsCount: snapshot.staticContactsCount,
          dynamicContactsCount: snapshot.dynamicContactsCount,
          activeProjectKeys: snapshot.activeProjectKeys
        });
      })().catch(error => {
        console.warn('[Cache] AB_GET_CACHED_DATA fallback failed', error);
        try { sendResponse({ success: false, error: String(error?.message || error) }); } catch (_) {}
      });
      return true;
    }

    case 'AB_GET_STATISTICS':

      sendResponse({
        success: true,
        stats: getScanStatistics()
      });
      return false;

    case 'AB_GET_TRACKED_PROJECTS':
      sendResponse({ success: true, projects: getTrackedProjectsConfigSnapshot() });
      return false;

    case 'AB_SET_TRACKED_PROJECTS':
      setTrackedProjectsConfig(Array.isArray(message.projects) ? message.projects : [])
        .then(projects => sendResponse({ success: true, projects }))
        .catch(error => sendResponse({ success: false, error: error?.message || String(error) }));
      return true;

    case 'AB_GET_TRACKED_AUDIT': {
      (async () => {
        try {
          if (typeof self.getTrackedAuditReport !== 'function') {
            self.getTrackedAuditReport = getTrackedAuditReport;
          }
          const reportFn = self.getTrackedAuditReport || globalThis.getTrackedAuditReport || getTrackedAuditReport;
          if (typeof reportFn !== 'function') {
            throw new Error('Tracked audit helper unavailable');
          }
          const { entries = [], summary = null } = await reportFn(message.filters || {});
          sendResponse({ success: true, entries, summary });
        } catch (error) {
          sendResponse({ success: false, error: error?.message || String(error) });
        }
      })();
      return true;
    }


    case 'AB_GET_CONTACT_AUDIT': {
      (async () => {
        try {
          const reportFn = typeof getContactAuditReport === 'function' ? getContactAuditReport : null;
          if (typeof reportFn !== 'function') {
            throw new Error('Contact audit helper unavailable');
          }
          const entries = await reportFn(message.filters || {});
          sendResponse({ success: true, entries });
        } catch (error) {
          sendResponse({ success: false, error: error?.message || String(error) });
        }
      })();
      return true;
    }

    case 'AB_EXTRACT_CONTACTS_FOR_AUDIT': {
      (async () => {
        try {
          if (!currentProjectFile) {
            throw new Error('No project file loaded');
          }
          if (!accessToken) {
            throw new Error('No access token available');
          }
          
          const result = await extractAllContactsForAudit(message.projectKeys || [], {
            accessToken,
            site: message.site || 'AutoBuilders'
          });
          
          sendResponse({ success: result.success, result });
        } catch (error) {
          sendResponse({ success: false, error: error?.message || String(error) });
        }
      })();
      return true;
    }

    case 'AB_AUDIT_RESCAN': {
      (async () => {
        try {
          if (typeof self.rescanTrackedProjects !== 'function') {
            self.rescanTrackedProjects = rescanTrackedProjects;
          }
          const rescanFn = self.rescanTrackedProjects || globalThis.rescanTrackedProjects || rescanTrackedProjects;
          if (typeof rescanFn !== 'function') {
            throw new Error('Tracked rescan helper unavailable');
          }
          const result = await rescanFn({
            triggeredBy: message.triggeredBy || 'manual',
            upload: message.upload !== false
          });
          sendResponse(result);
        } catch (error) {
          sendResponse({ success: false, error: error?.message || String(error) });
        }
      })();
      return true;
    }

    case 'AB_GET_USER_PROFILE':
      callGraphAPI('/me')
        .then(profile => sendResponse({ success: true, profile }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    
    case 'AB_AUTH_RESET': {
      try {
        if (chrome.identity && chrome.identity.clearAllCachedAuthTokens) {
          chrome.identity.clearAllCachedAuthTokens(() => {});
        }
      } catch (_) {}
      try { accessToken = null; tokenExpiry = 0; } catch (_) {}
      try { chrome.storage.local.remove(['accessToken','tokenExpiry']); } catch (_) {}
      if (typeof sendResponse === 'function') sendResponse({ success: true });
      return true;
    }

    case 'AB_TEST_MAIL': {
      const to = String(message.to || '').trim();
      const subject = message.subject || 'AB Test';
      const html = message.html || '<b>Test</b>';
      if (!to) { sendResponse?.({ success:false, error:'no_to' }); return true; }
      (async () => {
        try {
          const payload = {
            message: {
              subject,
              body: { contentType: 'HTML', content: html },
              toRecipients: [{ emailAddress: { address: to } }]
            },
            saveToSentItems: true
          };
          await callGraphAPI('/me/sendMail', { method:'POST', body: payload });
          sendResponse?.({ success: true });
        } catch (e) {
          sendResponse?.({ success:false, error: e?.message || String(e) });
        }
      })();
      return true;
    }


    case 'AB_EXPORT_EXCEL':
      exportContactsToExcel(message.contacts || cachedContacts)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_CLEAR_CACHE':
      clearCache()
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_DISCONNECT':
      Promise.all([
        chrome.storage.local.clear(),
        clearCache()
      ])
        .then(() => {
          accessToken = null;
          tokenExpiry = null;
          sendResponse({ success: true });
        })
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_LOAD_PROJECT':
      loadProjectWorkbook(message.projectName, {
        includeContacts: !!message.includeContacts,
        forceRefresh: !!message.forceRefresh,
        forceContactRefresh: !!message.forceContactRefresh
      })
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_LOAD_PROJECT_WORKBOOK':
      loadProjectWorkbook(message.projectName, {
        includeContacts: true,
        forceRefresh: !!message.forceRefresh,
        forceContactRefresh: !!message.forceContactRefresh
      })
        .then(result => {
          if (!result || !result.success) {
            sendResponse(result || { success: false, error: 'Failed to load project workbook' });
            return;
          }

          const contactsByDivision = result.contactsByDivision || {};
          const contacts = flattenContactsForPanel(contactsByDivision);

          sendResponse({
            success: true,
            project: result.project,
            fileName: result.fileName,
            source: result.source,
            divisionSource: result.divisionSource,
            divisions: result.divisions,
            contactsByDivision,
            contacts
          });
        })
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_PRECHECK_CONTACT':
      computeContactConflicts(message.projectName, message.contactData, message.divisionKeys || [], {
        includeOtherDivisions: message.includeOtherDivisions !== false,
        forceRefresh: !!message.forceRefresh
      })
        .then(report => sendResponse({ success: true, report }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_ADD_BIDDER':
      addBidderToProject(message.projectName, message.contactData, message.selectedDivisions, {
        overrideConflicts: !!message.overrideConflicts,
        forceRefresh: !!message.forceRefresh
      })
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_SUBMIT_BID':
      submitBidWorkflow(message)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_LOOKUP_ZIP':
      lookupZipLocation(message.zip)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_UPLOAD_CACHE_REMOTE':
      uploadCacheToSharePoint()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_SCAN_RECENT':
      runAutomaticScan({
        baseline: message.baseline === true,
        triggeredBy: message.triggeredBy || 'manual',
        upload: message.upload !== false
      })
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_DOWNLOAD_CACHE_REMOTE':
      downloadCacheFromSharePoint()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_SAVE_CACHE_BACKUP':
      exportCacheToFile()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_LOAD_CACHE_FROM_FILE':
      importCacheFromFile(message.fileContent)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'AB_GET_AUDIT_LOG':
      {
        const requested = Number(message.limit);
        const limit = Number.isFinite(requested) && requested > 0
          ? Math.min(requested, MAX_AUDIT_FETCH_LIMIT)
          : Math.min(AUDIT_LOG_DEFAULT_LIMIT, MAX_AUDIT_FETCH_LIMIT);
        DB.getContactAuditEntries({ limit })
          .then(entries => sendResponse({ success: true, entries }))
          .catch(error => {
            console.error('[Audit] Failed to load audit log:', error);
            sendResponse({ success: false, error: error?.message || String(error || 'Audit log failed') });
          });
      }
      return true;

    case 'AB_GET_CONTACT_AUDIT':
      getContactAuditReport(message.filters || {})
        .then(({ entries = [], summary = null }) => sendResponse({ success: true, entries, summary }))
        .catch(error => {
          console.error('[Audit] Failed to load audit report:', error);
          sendResponse({ success: false, error: error.message || String(error || 'Audit request failed') });
        });
      return true;

    case 'START_TRACKED_ENTRIES_SCAN':
      if (trackedEntriesScanInProgress) {
        sendResponse({ success: false, error: 'A scan is already in progress. Please wait for it to complete.' });
        return true;
      }
      
      trackedEntriesScanInProgress = true;
      (async () => {
        try {
          // Check if we have a valid token
          let token = accessToken;
          if (!token) {
            const stored = await chrome.storage.local.get(['accessToken']);
            token = stored.accessToken;
          }
          
          // If no token, authenticate first
          if (!token) {
            console.log('[TrackedEntries] No token found, starting authentication...');
            const authResult = await authenticateWithMicrosoft();
            if (!authResult.success) {
              throw new Error(`Authentication failed: ${authResult.error}`);
            }
            console.log('[TrackedEntries] Authentication successful, starting scan...');
          }
          
          // Now run the scan
          const result = await scanSelectedProjectsForTrackedEntries(message.projectKeys);
          sendResponse(result);
          // Broadcast final progress
          broadcastScanProgress({
            status: result.success ? 'Scan complete!' : 'Scan failed',
            percent: result.success ? 100 : 0,
            totalEntriesFound: result.totalEntries || 0
          });
        } catch (error) {
          console.error('[TrackedEntries] Error:', error);
          sendResponse({ success: false, error: error.message });
          broadcastScanProgress({
            status: `Error: ${error.message}`,
            percent: 0,
            totalEntriesFound: 0
          });
        } finally {
          trackedEntriesScanInProgress = false;
        }
      })();
      return true;

    case 'AB_SCAN_DELTA':
      if (deltaScanInProgress) {
        sendResponse({ success: false, error: 'A delta scan is already in progress. Please wait for it to complete.' });
        return true;
      }

      deltaScanInProgress = true;
      (async () => {
        try {
          const result = await runDeltaScan(Array.isArray(message.projectKeys) ? message.projectKeys : []);
          sendResponse(result);
        } catch (error) {
          console.error('[Delta] Error:', error);
          sendResponse({ success: false, error: error?.message || String(error) });
        } finally {
          deltaScanInProgress = false;
        }
      })();
      return true;


    case 'AB_RESET_TRACKED_ENTRIES':
      if (trackedEntriesScanInProgress || deltaScanInProgress) {
        sendResponse({ success: false, error: 'A scan is already in progress. Please wait for it to complete.' });
        return true;
      }

      trackedEntriesScanInProgress = true;
      let responded = false;

      // Set a timeout - if we don't respond in 120 seconds, fail
      const timeoutId = setTimeout(() => {
        if (!responded) {
          responded = true;
          console.error('[Reset] Reset handler timeout after 120 seconds');
          sendResponse({ success: false, error: 'Reset operation timed out (120s)' });
          trackedEntriesScanInProgress = false;
        }
      }, 120000);

      const projectKeys = Array.isArray(message.projectKeys) ? message.projectKeys : [];
      (async () => {
        try {
          console.log('[Reset] Reset starting with projects:', projectKeys);
          const result = await resetTrackedEntries(projectKeys);
          if (!responded) {
            responded = true;
            clearTimeout(timeoutId);
            sendResponse(result);
          }
        } catch (error) {
          if (!responded) {
            responded = true;
            clearTimeout(timeoutId);
            console.error('[Reset] Error during resetTrackedEntries:', error);
            sendResponse({ success: false, error: error?.message || String(error) });
          }
        } finally {
          trackedEntriesScanInProgress = false;
        }
      })();

      return true;

    default:
      sendResponse({ success: false, error: 'Unknown message type' });
      return false;
  }
});

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (!alarm) {
    return;
  }

  if (alarm.name === AUDIT_DAILY_ALARM_NAME) {
    try {
      await runDailyAuditSnapshot();
    } catch (error) {
      console.error('[Audit] Scheduled audit snapshot failed:', error);
    }
    return;
  }

  if (alarm.name === SCAN_ALARM_NAME) {
    // Handle scan alarm
    if (!AUTO_SCAN_ENABLED) {
      console.log('[Scan] Auto scan disabled; skipping scheduled scan alarm');
      return;
    }

    if (!needsIncrementalScan()) {
      console.log('[Scan] Incremental scan skipped (recently updated)');
      return;
    }

    await runAutomaticScan({ baseline: false, triggeredBy: 'alarm', upload: true });
  }
  
  // Let other alarms (like campaign alarms) pass through to other listeners
});

chrome.runtime.onInstalled.addListener(async () => {
  console.log('AutoBuilders Extension installed');
  try {
    await bootstrapCache({ triggeredBy: 'onInstalled', hydrateRemote: true });
    if (AUTO_SCAN_ENABLED && needsBaselineScan()) {
      runAutomaticScan({ baseline: true, triggeredBy: 'onInstalled', upload: true })
        .then(result => {
          if (result && result.success) {
            console.log('[Scan] Baseline scan completed on install');
          } else {
            console.warn('[Scan] Baseline scan finished with issues on install', result);
          }
        })
        .catch(error => console.error('[Scan] Baseline scan failed on install:', error));
    } else if (!AUTO_SCAN_ENABLED) {
      console.log('[Scan] Auto scan disabled; skipping baseline scan on install');
    } else {
      console.log('[Scan] Baseline scan not required during install');
    }
  } catch (error) {
    console.error('Error during install bootstrap:', error);
  }
});

chrome.runtime.onStartup.addListener(async () => {
  console.log('AutoBuilders Extension starting up');
  try {
    // Just load the cache from IndexedDB - don't auto-scan
    // User can manually trigger scans via UI
    await loadCacheFromIndexedDB();
    console.log('[Scan] Startup: cache loaded from IndexedDB. Manual scan available via UI.');
  } catch (error) {
    console.error('Error during startup bootstrap:', error);
  }
});

bootstrapCache({ triggeredBy: 'initial-load', hydrateRemote: false }).catch(error => {
  console.error('Initial cache bootstrap failed:', error);
});

















