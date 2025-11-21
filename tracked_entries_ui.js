/**
 * Tracked Entries UI - Scan, Display, Filter, and Sort
 */

// Global state for tracked entries
let trackedEntriesData = {};
let trackedEntriesScanInProgress = false;
let lastDeltaSummary = { message: '', variant: 'info' };
let trackedEntriesCurrentFilter = { startDate: null, endDate: null };

const TRACKED_DB_NAME = 'AutoBuildersDB';
const TRACKED_DB_VERSION = 5;
let trackedEntriesMessageListenerRegistered = false;

/**
 * Format a project label with code-first naming.
 */
function formatTrackedProjectLabel(project = {}) {
  const toString = (value) => {
    if (value === null || value === undefined) return '';
    return typeof value === 'string' ? value.trim() : String(value).trim();
  };

  const codeCandidates = [
    toString(project.projectNumber),
    toString(project.projectFolderName),
    toString(project.displayName),
    toString(project.name),
    toString(project.projectDisplayName),
    toString(project.projectName)
  ];

  let code = '';
  for (const candidate of codeCandidates) {
    const extracted = extractProjectCodeFromText(candidate);
    if (extracted && extracted.nearEnd && extracted.index > 0) {
      code = extracted.code;
      break;
    }
  }

  const descriptorCandidates = [
    toString(project.displayName),
    toString(project.name),
    toString(project.projectDisplayName),
    toString(project.projectFolderName),
    toString(project.projectName),
    toString(project.fileName)
  ];

  let descriptor = descriptorCandidates.find(value => value && value.trim()) || '';

  if (code && descriptor) {
    const flexibleCodePattern = code.replace('-', '[-\s_]?');
    const codeRegex = new RegExp(`\b${flexibleCodePattern}\b`, 'i');
    descriptor = descriptor.replace(codeRegex, '').replace(/[\s-]+$/, '').trim();
  }

  const fallback = toString(project.projectKey) || toString(project.key) || 'Unknown Project';
  const descriptorLabel = descriptor || fallback;
  const label = code ? `${code} ${descriptorLabel}`.trim() : descriptorLabel;

  return {
    label,
    code,
    descriptor: descriptorLabel
  };
}


function extractProjectCodeFromText(text = '') {
  if (!text) return null;
  const str = String(text);
  const matches = Array.from(str.matchAll(/(\d{2})[-\s_]?(\d{2,3})/g));
  if (!matches.length) return null;
  const last = matches[matches.length - 1];
  const index = last.index != null ? last.index : str.lastIndexOf(last[0]);
  if (index < 0) return null;
  const tokenLength = last[0].length;
  const nearEnd = index + tokenLength >= str.length - 2;
  return { code: `${last[1]}-${last[2]}`, index, nearEnd };
}

function projectCodeToSortValue(code = '') {
  if (!code) return null;
  const parts = code.split('-');
  if (parts.length !== 2) return null;
  const major = parseInt(parts[0], 10);
  const minor = parseInt(parts[1], 10);
  if (!Number.isFinite(major) || !Number.isFinite(minor)) return null;
  return major * 1000 + minor;
}


/**
 * Build a lowercase search index for a project record.
 */
function buildTrackedProjectSearchIndex(project = {}, label = '') {
  const fields = [
    label,
    project.projectFolderName,
    project.displayName,
    project.name,
    project.projectNumber,
    project.projectKey,
    project.key,
    project.state,
    project.stateName,
    project.city,
    project.zip
  ];

  return fields
    .map(value => {
      if (value === null || value === undefined) return '';
      return String(value).toLowerCase();
    })
    .filter(Boolean)
    .join(' ');
}

/**
 * Safely parse a date value and return a timestamp.
 */
function parseTrackedProjectDate(value) {
  if (!value) return null;
  let date;
  if (value instanceof Date) {
    date = value;
  } else if (typeof value === 'number') {
    date = new Date(value);
  } else {
    const trimmed = String(value).trim();
    if (!trimmed) return null;
    date = new Date(trimmed);
  }
  const time = date.getTime();
  return Number.isNaN(time) ? null : time;
}

/**
 * Initialize tracked entries module
 */
async function initTrackedEntriesUI() {
  console.log('[TrackedUI] Initializing tracked entries UI');

  if (chrome.runtime && chrome.runtime.onMessage && !trackedEntriesMessageListenerRegistered) {
    chrome.runtime.onMessage.addListener((request) => {
      if (request.type === 'TRACKED_SCAN_PROGRESS') {
        updateScanProgressUI(request.data);
      }
    });
    trackedEntriesMessageListenerRegistered = true;
  }

  await loadTrackedEntries();
  renderTrackedEntriesUI();
}

/**
 * Load tracked entries from Chrome storage
 */
async function loadTrackedEntries() {
  try {
    const result = await chrome.storage.local.get('AB_TRACKED_ENTRIES');
    trackedEntriesData = result.AB_TRACKED_ENTRIES || {};
    console.log('[TrackedUI] Loaded tracked entries:', Object.keys(trackedEntriesData).length, 'projects');
    return trackedEntriesData;
  } catch (error) {
    console.error('[TrackedUI] Error loading entries:', error);
    return {};
  }
}

/**
 * Setup searchable, multi-select projects dropdown
 */
async function setupTrackedProjectsSearchDropdown() {
  const searchInput = $('trackedProjectSearchInput');
  const selector = $('trackedProjectSelectorDropdown');
  
  if (!searchInput || !selector) {
    console.error('[TrackedUI] Search/selector elements not found');
    return;
  }

  // Initialize selected projects array
  window.selectedProjectsForScan = window.selectedProjectsForScan || [];

  // Get all available projects from IndexedDB
  let allProjects = [];
  
  try {
    allProjects = await loadProjectsFromIndexedDB();
    console.log('[TrackedUI] Loaded', allProjects.length, 'projects from IndexedDB');
  } catch (error) {
    console.error('[TrackedUI] Error loading projects from IndexedDB:', error);
  }

  if (allProjects.length === 0) {
    selector.innerHTML = '<div style="padding: 12px; color: #d32f2f; text-align: center; font-size: 12px;">No projects found. Load projects in the Projects tab first, then return to Audit.</div>';
    return;
  }

  // Build and sort projects list (most recent first)
  let allProjectsList = [];

  const buildProjectsList = () => {
    const projects = [];
    const seenKeys = new Set();
    allProjects.forEach(p => {
      if (!p) return;

      const projectKey = p.key || p.projectKey || p.id || (p.projectFolderName ? `${p.projectFolderName}_${p.displayName || p.name || ''}` : null);
      if (!projectKey) return;
      const normalizedKey = String(projectKey).trim().toLowerCase();
      if (!normalizedKey || seenKeys.has(normalizedKey)) return;
      seenKeys.add(normalizedKey);

      const labelInfo = formatTrackedProjectLabel(p);
      const label = labelInfo.label;
      const stateName = (p.state || p.stateName || '').trim();
      const lastModifiedRaw = p.lastModifiedDateTime || p.fileModified || p.modifiedTime || p.lastUpdated || p.lastRefreshed || p.timeCreated || '';
      const lastModifiedTs = parseTrackedProjectDate(lastModifiedRaw);
      const codeSortValue = projectCodeToSortValue(labelInfo.code);
      const searchIndex = `${buildTrackedProjectSearchIndex(p, label)} ${normalizedKey}`.trim();

      projects.push({
        key: projectKey,
        name: label,
        state: stateName,
        lastModifiedRaw,
        sortTimestamp: Number.isFinite(lastModifiedTs) ? lastModifiedTs : null,
        searchIndex,
        code: labelInfo.code,
        descriptor: labelInfo.descriptor,
        codeSortValue
      });
    });

    projects.sort((a, b) => {
      const hasCodeA = Number.isFinite(a.codeSortValue);
      const hasCodeB = Number.isFinite(b.codeSortValue);
      if (hasCodeA && hasCodeB) {
        if (a.codeSortValue !== b.codeSortValue) {
          return b.codeSortValue - a.codeSortValue;
        }
        return a.name.localeCompare(b.name);
      }
      if (hasCodeA) return -1;
      if (hasCodeB) return 1;

      const hasTimeA = Number.isFinite(a.sortTimestamp);
      const hasTimeB = Number.isFinite(b.sortTimestamp);
      if (hasTimeA && hasTimeB) {
        return b.sortTimestamp - a.sortTimestamp;
      }
      if (hasTimeA) return -1;
      if (hasTimeB) return 1;

      return a.name.localeCompare(b.name);
    });

    return projects;
  };

  allProjectsList = buildProjectsList();

  console.log('[TrackedUI] Built project list with', allProjectsList.length, 'projects');

  // Render dropdown items with scroll
  const renderItems = (items) => {
    selector.innerHTML = '';

    if (items.length === 0) {
      selector.innerHTML = '<div style="padding: 12px; color: #999; text-align: center; font-size: 12px;">No projects match your search</div>';
      return;
    }

    const fragment = document.createDocumentFragment();
    items.forEach((item, index) => {
      const isSelected = window.selectedProjectsForScan.find(p => p.key === item.key);

      const label = document.createElement('label');
      label.style.cssText = `
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 8px 12px;
        border-bottom: 1px solid #eee;
        cursor: pointer;
        background: ${isSelected ? '#e8f5e9' : index % 2 === 0 ? '#fafafa' : 'white'};
        transition: background 0.2s;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      `;
      label.title = item.name;
      label.onmouseover = () => { label.style.background = '#f0f0f0'; };
      label.onmouseout = () => {
        label.style.background = isSelected ? '#e8f5e9' : (index % 2 === 0 ? '#fafafa' : 'white');
      };
      label.dataset.projectKey = item.key;

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.dataset.projectKey = item.key;
      checkbox.value = item.key;
      checkbox.checked = !!isSelected;
      checkbox.style.cursor = 'pointer';
      checkbox.style.width = '18px';
      checkbox.style.height = '18px';

      label.appendChild(checkbox);

      const textWrapper = document.createElement('span');
      textWrapper.style.cssText = 'display: flex; align-items: center; gap: 8px; flex: 1; overflow: hidden;';

      if (item.code) {
        const codeSpan = document.createElement('span');
        codeSpan.textContent = item.code;
        codeSpan.style.cssText = 'font-size: 13px; font-weight: 600; color: #1b5e20; white-space: nowrap;';
        textWrapper.appendChild(codeSpan);
      }

      const descriptorSpan = document.createElement('span');
      descriptorSpan.textContent = item.descriptor || item.name;
      descriptorSpan.style.cssText = 'flex: 1; font-size: 13px; font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;';
      textWrapper.appendChild(descriptorSpan);

      label.appendChild(textWrapper);

      if (item.state) {
        const stateSpan = document.createElement('span');
        stateSpan.style.cssText = 'font-size: 11px; color: #666; margin-left: 4px; padding: 2px 6px; background: #e0e0e0; border-radius: 3px; white-space: nowrap;';
        stateSpan.textContent = item.state;
        label.appendChild(stateSpan);
      }

      checkbox.addEventListener('change', (e) => {
        e.stopPropagation();
        if (checkbox.checked) {
          if (!window.selectedProjectsForScan.find(p => p.key === item.key)) {
            window.selectedProjectsForScan.push({ key: item.key, name: item.name });
            console.log('[TrackedUI] Added project:', item.name);
          }
        } else {
          window.selectedProjectsForScan = window.selectedProjectsForScan.filter(p => p.key !== item.key);
          console.log('[TrackedUI] Removed project:', item.name);
        }
        updateSelectedProjectsDisplay();
      });

      label.addEventListener('click', (e) => {
        if (e.target !== checkbox) {
          checkbox.checked = !checkbox.checked;
          checkbox.dispatchEvent(new Event('change'));
        }
      });

      fragment.appendChild(label);
    });

    selector.appendChild(fragment);

  };

  // Initial render - show entire list
  renderItems(allProjectsList);

  // Store full list for searching
  let currentList = allProjectsList;

  // Search input handler
  searchInput.addEventListener('input', (e) => {
    const searchTerm = e.target.value.toLowerCase().trim();
    
    let filtered = allProjectsList;
    
    if (searchTerm) {
      const tokens = searchTerm.split(/\s+/).filter(Boolean);
      filtered = allProjectsList.filter(project => {
        if (!tokens.length) return true;
        return tokens.every(token => project.searchIndex.includes(token));
      });
    }

    currentList = filtered;
    renderItems(filtered);
    console.log('[TrackedUI] Search applied:', searchTerm, '- found', filtered.length, 'projects');
  });

  console.log('[TrackedUI] Dropdown initialized with', allProjectsList.length, 'projects (sorted by project number)');
}

/**
 * Load projects from IndexedDB
 */
function loadProjectsFromIndexedDB() {
  return new Promise((resolve, reject) => {
    const openDatabase = (useVersion) => {
      const request = useVersion
        ? indexedDB.open(TRACKED_DB_NAME, TRACKED_DB_VERSION)
        : indexedDB.open(TRACKED_DB_NAME);

      request.onerror = () => {
        if (useVersion && request.error && request.error.name === 'VersionError') {
          console.warn(`[IndexedDB] Version mismatch opening ${TRACKED_DB_NAME}, retrying without explicit version`);
          openDatabase(false);
          return;
        }
        console.error('[IndexedDB] Failed to open database', request.error);
        reject(new Error('Failed to open IndexedDB'));
      };

      request.onupgradeneeded = (event) => {
        const upgradeDb = event.target.result;
        if (!upgradeDb.objectStoreNames.contains('projects')) {
          upgradeDb.createObjectStore('projects', { keyPath: 'id', autoIncrement: true });
        }
      };

      request.onsuccess = (event) => {
        const db = event.target.result;

        if (!db.objectStoreNames.contains('projects')) {
          console.warn(`[IndexedDB] projects store not found in ${TRACKED_DB_NAME}`);
          db.close();
          resolve([]);
          return;
        }

        const tx = db.transaction(['projects'], 'readonly');
        const store = tx.objectStore('projects');
        const getAllRequest = store.getAll();

        tx.oncomplete = () => db.close();
        tx.onabort = () => {
          console.error('[IndexedDB] Projects transaction aborted', tx.error);
          db.close();
        };

        getAllRequest.onsuccess = () => {
          const results = getAllRequest.result || [];
          console.log('[IndexedDB] Loaded', results.length, 'projects');
          resolve(results);
        };

        getAllRequest.onerror = () => {
          console.error('[IndexedDB] Failed to get projects', getAllRequest.error);
          db.close();
          reject(new Error('Failed to get projects from IndexedDB'));
        };
      };
    };

    openDatabase(true);
  });
}

/**
 * Update the selected projects display
 */
function updateSelectedProjectsDisplay() {
  const container = $('selectedProjectsList');
  const countSpan = $('selectedCount');
  
  if (!container) return;

  if (!window.selectedProjectsForScan || window.selectedProjectsForScan.length === 0) {
    container.innerHTML = '<span style="color: #999; font-size: 12px;">None selected</span>';
    if (countSpan) countSpan.textContent = '0';
    return;
  }

  if (countSpan) countSpan.textContent = window.selectedProjectsForScan.length;

  container.innerHTML = window.selectedProjectsForScan
    .map(p => `
      <div style="
        background: #4CAF50;
        color: white;
        padding: 4px 8px;
        border-radius: 3px;
        font-size: 12px;
        display: flex;
        align-items: center;
        gap: 6px;
        white-space: nowrap;
      ">
        <span>${p.name}</span>
        <span style="cursor: pointer; font-weight: bold; margin-left: 4px;" onclick="removeProjectFromSelection('${p.key}')">×</span>
      </div>
    `)
    .join('');
}

/**
 * Remove a project from selection
 */
function removeProjectFromSelection(projectKey) {
  window.selectedProjectsForScan = (window.selectedProjectsForScan || [])
    .filter(p => p.key !== projectKey);
  updateSelectedProjectsDisplay();
  setupTrackedProjectsSearchDropdown(); // Re-render dropdown to show unchecked
  console.log('[TrackedUI] Project removed from selection:', projectKey);
}

/**
 * Start initial scan of selected projects only
 */
async function startTrackedEntriesScan() {
  try {
    if (trackedEntriesScanInProgress) {
      console.log('[TrackedUI] Scan already in progress');
      alert('Scan already in progress. Please wait for it to complete.');
      return;
    }

    // Get selected projects from global state
    const selectedProjectKeys = (window.selectedProjectsForScan || []).map(p => p.key);
    
    if (selectedProjectKeys.length === 0) {
      alert('Please select at least one project before scanning.');
      return;
    }

    console.log(`[TrackedUI] Selected ${selectedProjectKeys.length} project(s) for scan:`, selectedProjectKeys);

    setDeltaScanSummary('');
    trackedEntriesScanInProgress = true;
    showScanProgressUI({ title: 'Scanning Tracked Entries...', status: 'Preparing scan...' });
    console.log('[TrackedUI] Sending scan request to background');

    // Send message to background with selected projects
    const response = await new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({
        type: 'START_TRACKED_ENTRIES_SCAN',
        projectKeys: selectedProjectKeys
      }, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(response);
        }
      });
    });

    console.log('[TrackedUI] Scan response:', response);

    if (!response || !response.success) {
      throw new Error(response?.error || 'Scan failed');
    }

    // Load results
    console.log('[TrackedUI] Loading scan results');
    await loadTrackedEntries();
    renderTrackedEntriesUI();
    hideScanProgressUI();
    trackedEntriesScanInProgress = false;
    console.log('[TrackedUI] Scan complete and UI updated');

  } catch (error) {
    console.error('[TrackedUI] Scan error:', error);
    alert(`Scan failed: ${error.message}`);
    hideScanProgressUI();
    trackedEntriesScanInProgress = false;
  }
}

async function startDeltaScan() {
  try {
    if (trackedEntriesScanInProgress) {
      console.log('[TrackedUI] Delta scan already in progress');
      alert('Another scan is in progress. Please wait for it to complete.');
      return;
    }

    const selectedProjectKeys = (window.selectedProjectsForScan || []).map(p => p.key);

    if (selectedProjectKeys.length === 0) {
      alert('Please select at least one project before running a delta scan.');
      return;
    }

    setDeltaScanSummary('');
    trackedEntriesScanInProgress = true;
    showScanProgressUI({ title: 'Running Delta Scan...', status: 'Preparing delta query...' });

    const response = await new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({
        type: 'AB_SCAN_DELTA',
        projectKeys: selectedProjectKeys
      }, (reply) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(reply);
        }
      });
    });

    console.log('[TrackedUI] Delta scan response:', response);

    if (!response || response.success === false) {
      throw new Error(response?.error || 'Delta scan failed');
    }

    const changedCount = Number(response.changed || 0);
    const refreshedCount = Number(response.refreshed || 0);
    const deletedCount = Number(response.deleted || 0);
    const errorCount = Array.isArray(response.errors) ? response.errors.length : 0;

    updateScanProgressUI({
      percent: 100,
      status: 'Delta scan complete!',
      currentProject: 'Summary',
      totalEntriesFound: changedCount,
      eta: `Refreshed: ${refreshedCount} | Deleted: ${deletedCount}`
    });

    hideScanProgressUI();

    const parts = [
      `changed ${changedCount}`,
      `refreshed ${refreshedCount}`,
      `deleted ${deletedCount}`
    ];
    if (errorCount) {
      parts.push(`errors ${errorCount}`);
    }
    const summaryMessage = `Delta scan complete - ${parts.join(', ')}.`;
    const summaryVariant = errorCount ? 'error' : 'success';
    setDeltaScanSummary(summaryMessage, summaryVariant);

    await loadTrackedEntries();
    renderTrackedEntriesUI();
    setDeltaScanSummary(summaryMessage, summaryVariant);
  } catch (error) {
    console.error('[TrackedUI] Delta scan error:', error);
    alert(`Delta scan failed: ${error.message}`);
    setDeltaScanSummary(`Delta scan failed: ${error.message}`, 'error');
    hideScanProgressUI();
  } finally {
    trackedEntriesScanInProgress = false;
  }
}
async function resetAndFullScan() {
  try {
    if (trackedEntriesScanInProgress) {
      console.log('[TrackedUI] Reset requested while another scan is running');
      alert('Another scan is in progress. Please wait for it to finish.');
      return;
    }

    const selectedProjectKeys = (window.selectedProjectsForScan || []).map(p => p.key);
    if (selectedProjectKeys.length === 0) {
      alert('Please select at least one project before resetting.');
      return;
    }

    setDeltaScanSummary('');
    trackedEntriesScanInProgress = true;
    showScanProgressUI({ title: 'Resetting tracked entries...', status: 'Clearing cached data...' });

    const resetResponse = await new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({
        type: 'AB_RESET_TRACKED_ENTRIES',
        projectKeys: selectedProjectKeys
      }, (reply) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(reply);
        }
      });
    });

    if (!resetResponse || resetResponse.success === false) {
      throw new Error(resetResponse?.error || 'Reset failed');
    }

    hideScanProgressUI();

    const missingList = Array.isArray(resetResponse.missingProjects) ? resetResponse.missingProjects : [];
    const summaryPieces = [
      `projects processed: ${resetResponse.projectsProcessed || 0}`,
      `projects touched: ${resetResponse.projectsTouched || 0}`,
      `cached entries removed: ${resetResponse.trackedEntriesRemoved || 0}`,
      `audit rows removed: ${resetResponse.removedAuditEntries || 0}`,
      `contacts removed: ${resetResponse.contactsRemoved || 0}`,
      `file index trimmed: ${resetResponse.fileIndexRemoved || 0}`,
      `delta links cleared: ${resetResponse.deltaLinksCleared || 0}`
    ];
    const summaryMessage = `Reset complete - ${summaryPieces.join(', ')}.`;
    const summaryVariant = missingList.length ? 'error' : 'success';
    const missingMessage = missingList.length ? ` Missing: ${missingList.join(', ')}` : '';

    trackedEntriesScanInProgress = false;
    await startTrackedEntriesScan();
    setDeltaScanSummary(summaryMessage + missingMessage, summaryVariant);
  } catch (error) {
    console.error('[TrackedUI] Reset and full scan error:', error);
    alert(`Reset failed: ${error.message}`);
    hideScanProgressUI();
    trackedEntriesScanInProgress = false;
    setDeltaScanSummary(`Reset failed: ${error.message}`, 'error');
  }
}


/**
 * Show progress bar and status UI
 */
function showScanProgressUI(options = {}) {
  const container = $('trackedScanProgressContainer');
  if (!container) return;

  const {
    title = 'Scanning Tracked Entries...',
    status = 'Starting scan...',
    percent = 0
  } = options;

  const percentLabel = Math.max(0, Math.min(100, Math.round(percent)));

  container.innerHTML = `
    <div id="trackedScanOverlay" style="
      background: #f5f5f5;
      border: 1px solid #ddd;
      border-radius: 4px;
      padding: 16px;
      margin-bottom: 16px;
    ">
      <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
        <span style="font-weight: bold;">${title}</span>
        <span id="trackedScanProgress" style="font-size: 12px; color: #666;">${percentLabel}%</span>
      </div>
      <div style="width: 100%; height: 20px; background: #e0e0e0; border-radius: 2px; overflow: hidden;">
        <div id="trackedScanProgressBar" style="
          height: 100%;
          background: linear-gradient(90deg, #4CAF50, #45a049);
          width: ${percentLabel}%;
          transition: width 0.3s ease;
        "></div>
      </div>
      <div id="trackedScanStatus" style="margin-top: 8px; font-size: 12px; color: #666;">${status}</div>
      <div id="trackedScanStats" style="margin-top: 8px; font-size: 12px; color: #333; font-weight: bold;"></div>
    </div>
  `;

  container.style.display = 'block';
}

/**
 * Hide progress UI
 */
function hideScanProgressUI() {
  const container = $('trackedScanProgressContainer');
  if (container) {
    container.style.display = 'none';
  }
}

function setDeltaScanSummary(message, variant = 'info') {
  lastDeltaSummary = { message, variant };
  const summaryEl = $('deltaScanSummary');
  if (!summaryEl) return;

  if (!message) {
    summaryEl.textContent = '';
    summaryEl.style.display = 'none';
    return;
  }

  summaryEl.textContent = message;
  summaryEl.style.display = 'block';

  if (variant === 'error') {
    summaryEl.style.color = '#d32f2f';
  } else if (variant === 'success') {
    summaryEl.style.color = '#2e7d32';
  } else {
    summaryEl.style.color = '#333';
  }
}

/**
 * Update progress UI with scan data
 */
function updateScanProgressUI(progressData) {
  const progressBar = $('trackedScanProgressBar');
  const progressText = $('trackedScanProgress');
  const statusText = $('trackedScanStatus');
  const statsText = $('trackedScanStats');

  if (progressBar) {
    const percent = Math.round(progressData.percent || 0);
    progressBar.style.width = percent + '%';
  }

  if (progressText) {
    progressText.textContent = `${Math.round(progressData.percent || 0)}%`;
  }

  if (statusText) {
    statusText.textContent = progressData.status || '';
  }

  if (statsText) {
    statsText.innerHTML = `
      <div>Current Project: ${progressData.currentProject || 'N/A'}</div>
      <div>Total Entries Found: ${progressData.totalEntriesFound || 0}</div>
      <div>ETA: ${progressData.eta || 'Calculating...'}</div>
    `;
  }
}

/**
 * Render the main tracked entries UI
 */
function renderTrackedEntriesUI() {
  const container = $('trackedEntriesContainer');
  if (!container) {
    console.error('[TrackedUI] trackedEntriesContainer not found');
    return;
  }

  // Build project cards from loaded data
  const projectCards = Object.entries(trackedEntriesData)
    .map(([ projectKey, projectData ], index) => buildProjectCard(projectKey, projectData))
    .join('');

  container.innerHTML = `
    <div style="margin-bottom: 16px; padding: 12px; background: #f9f9f9; border: 1px solid #ddd; border-radius: 4px;">
      <div style="margin-bottom: 12px; font-weight: bold; color: #333; font-size: 14px;">Select Projects to Scan:</div>
      
      <div style="margin-bottom: 8px;">
        <input type="text" 
          id="trackedProjectSearchInput" 
          placeholder="Search projects..." 
          style="
            width: 100%;
            padding: 8px;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 13px;
            box-sizing: border-box;
          ">
      </div>

      <div id="trackedProjectSelectorDropdown" style="
        max-height: 250px; 
        overflow-y: auto; 
        border: 1px solid #ccc; 
        border-radius: 4px; 
        background: white;
        margin-bottom: 12px;
      ">
        <div style="padding: 8px; color: #999; text-align: center; font-size: 12px;">Loading projects...</div>
      </div>

      <div style="margin-bottom: 12px; min-height: 30px; padding: 8px; background: #e8f5e9; border: 1px solid #4CAF50; border-radius: 4px;">
        <div style="font-size: 12px; color: #666; margin-bottom: 4px;">Selected Projects: <span id="selectedCount" style="font-weight: bold; color: #333;">0</span></div>
        <div id="selectedProjectsList" style="display: flex; flex-wrap: wrap; gap: 6px; min-height: 24px;">
          <span style="color: #999; font-size: 12px;">None selected</span>
        </div>
      </div>

      <div style="display: flex; flex-wrap: wrap; gap: 8px; align-items: center;">
        <button id="clearSelectedProjectsBtn" style="
          padding: 6px 12px;
          background: #f44336;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 12px;
        ">Clear Selection</button>
        <div style="flex: 1 1 auto;"></div>
        <button id="resetAndFullScanBtn" style="
          padding: 8px 16px;
          background: #d32f2f;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-weight: bold;
          font-size: 14px;
        ">
          Reset & Full Scan
        </button>
        <button id="startDeltaScanBtn" style="
          padding: 8px 16px;
          background: #1976D2;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-weight: bold;
          font-size: 14px;
        ">
          Delta Scan
        </button>
        <button id="startTrackedScanBtn" style="
          padding: 8px 16px;
          background: #4CAF50;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-weight: bold;
          font-size: 14px;
        ">
          Full Scan
        </button>
      </div>
      <div id="deltaScanSummary" style="margin-top: 8px; font-size: 12px; color: #333; display: none;"></div>
    </div>

    <div id="trackedScanProgressContainer" style="display: none;"></div>

    <div id="trackedProjectsList">
      ${projectCards || '<p style="color: #666; padding: 16px;">No entries scanned yet. Select projects above and click "Full Scan" to begin.</p>'}
    </div>
  `;

  // Attach all event listeners after DOM is populated
  attachTrackedEntriesEventListeners();
  // Call async function to setup dropdown
  setupTrackedProjectsSearchDropdown();
  setDeltaScanSummary(lastDeltaSummary.message, lastDeltaSummary.variant);
}

/**
 * Attach event listeners to rendered elements
 */
function attachTrackedEntriesEventListeners() {
  // Initialize selected projects array
  window.selectedProjectsForScan = window.selectedProjectsForScan || [];

  // Clear selection button
  const clearBtn = $('clearSelectedProjectsBtn');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      window.selectedProjectsForScan = [];
      $('trackedProjectSearchInput').value = '';
      updateSelectedProjectsDisplay();
      setupTrackedProjectsSearchDropdown();
      setDeltaScanSummary('');
      console.log('[TrackedUI] All selections cleared');
    });
  }

  // Attach main scan button click handler
  const scanBtn = $('startTrackedScanBtn');
  if (scanBtn) {
    scanBtn.addEventListener('click', (e) => {
      e.preventDefault();
      console.log('[TrackedUI] Full scan button clicked');
      startTrackedEntriesScan();
    });
    console.log('[TrackedUI] Attached click handler to full scan button');
  } else {
    console.error('[TrackedUI] startTrackedScanBtn not found after render');
  }

  const deltaBtn = $('startDeltaScanBtn');
  if (deltaBtn) {
    deltaBtn.addEventListener('click', (e) => {
      e.preventDefault();
      console.log('[TrackedUI] Delta scan button clicked');
      startDeltaScan();
    });
    console.log('[TrackedUI] Attached click handler to delta scan button');
  } else {
    console.error('[TrackedUI] startDeltaScanBtn not found after render');
  }

  const resetBtn = $('resetAndFullScanBtn');
  if (resetBtn) {
    resetBtn.addEventListener('click', (e) => {
      e.preventDefault();
      console.log('[TrackedUI] Reset and full scan button clicked');
      resetAndFullScan();
    });
    console.log('[TrackedUI] Attached click handler to reset/full scan button');
  } else {
    console.error('[TrackedUI] resetAndFullScanBtn not found after render');
  }

  // Attach project card header click handlers
  Object.keys(trackedEntriesData).forEach(projectKey => {
    const header = $(`trackedCard_${projectKey}_header`);
    if (header) {
      header.addEventListener('click', () => {
        toggleProjectCard(`trackedCard_${projectKey}`);
      });
    }
  });

  // Attach filter button handlers
  Object.keys(trackedEntriesData).forEach(projectKey => {
    const applyBtn = $(`applyFilter_${projectKey}`);
    const clearBtn = $(`clearFilter_${projectKey}`);
    
    if (applyBtn) {
      applyBtn.addEventListener('click', () => applyDateFilter(projectKey));
    }
    if (clearBtn) {
      clearBtn.addEventListener('click', () => clearDateFilter(projectKey));
    }

    // Attach table header sort handlers
    const fields = ['division', 'company', 'contactName', 'phone', 'email', 'sharePointUser', 'dateAdded'];
    fields.forEach(field => {
      const header = $(`sortHeader_${projectKey}_${field}`);
      if (header) {
        header.addEventListener('click', () => {
          console.log(`[TrackedUI] Sort clicked: ${projectKey} by ${field}`);
          sortEntriesBy(projectKey, field);
        });
      }
    });
  });
}

/**
 * Build individual project card (collapsible)
 */
function buildProjectCard(projectKey, projectData) {
  const entries = projectData?.entries || [];
  const cardId = `trackedCard_${projectKey}`;
  const projectName = projectData?.projectName || projectKey;

  return `
    <div style="
      border: 1px solid #ddd;
      border-radius: 4px;
      margin-bottom: 12px;
      overflow: hidden;
    ">
      <div style="
        background: #f9f9f9;
        padding: 12px;
        cursor: pointer;
        display: flex;
        justify-content: space-between;
        align-items: center;
      " id="${cardId}_header">
        <div style="display: flex; align-items: center; gap: 8px; flex: 1;">
          <span id="${cardId}_arrow" style="display: inline-block; transition: transform 0.2s;">▶</span>
          <span style="font-weight: bold;">${projectName}</span>
          <span style="font-size: 12px; color: #666;">(${entries.length} entries)</span>
        </div>
        <span style="font-size: 12px; color: #999;">
          ${projectData?.lastScanned ? 'Last scanned: ' + new Date(projectData.lastScanned).toLocaleDateString() : 'Not scanned'}
        </span>
      </div>

      <div id="${cardId}_content" style="display: none; padding: 16px; background: white; border-top: 1px solid #ddd;">
        ${buildProjectContent(projectKey, entries)}
      </div>
    </div>
  `;
}

/**
 * Build content for expanded project card
 */
function buildProjectContent(projectKey, entries) {
  if (entries.length === 0) {
    return '<p style="color: #999;">No entries found for this project.</p>';
  }

  const filteredEntries = filterEntriesByDateRange(entries);
  const divisionSummary = calculateDivisionSummary(filteredEntries);
  const dailySummary = calculateDailySummary(filteredEntries);

  return `
    <div style="margin-bottom: 16px;">
      <h4 style="margin: 0 0 8px 0;">Division Summary</h4>
      ${buildDivisionTable(divisionSummary)}
    </div>

    <div style="margin-bottom: 16px;">
      <h4 style="margin: 0 0 8px 0;">Daily Summary</h4>
      ${buildDailySummaryHTML(dailySummary)}
    </div>

    <div style="margin-bottom: 16px;">
      <h4 style="margin: 0 0 8px 0;">Date Range Filter</h4>
      <div style="display: flex; gap: 8px; align-items: center;">
        <input type="date" id="${projectKey}_startDate" style="padding: 6px; border: 1px solid #ddd; border-radius: 3px;">
        <span>to</span>
        <input type="date" id="${projectKey}_endDate" style="padding: 6px; border: 1px solid #ddd; border-radius: 3px;">
        <button id="applyFilter_${projectKey}" style="
          padding: 6px 12px;
          background: #2196F3;
          color: white;
          border: none;
          border-radius: 3px;
          cursor: pointer;
        ">Apply Filter</button>
        <button id="clearFilter_${projectKey}" style="
          padding: 6px 12px;
          background: #f44336;
          color: white;
          border: none;
          border-radius: 3px;
          cursor: pointer;
        ">Clear</button>
      </div>
    </div>

    <div style="margin-bottom: 16px;">
      <h4 style="margin: 0 0 8px 0;">Entries (${filteredEntries.length} total)</h4>
      ${buildEntriesTable(projectKey, filteredEntries)}
    </div>
  `;
}

/**
 * Calculate division summary
 */
function calculateDivisionSummary(entries) {
  const summary = {};
  entries.forEach(entry => {
    const div = entry.division;
    summary[div] = (summary[div] || 0) + 1;
  });
  return summary;
}

/**
 * Build division summary table
 */
function buildDivisionTable(summary) {
  const divisions = Object.keys(summary).sort();
  if (!divisions.length) {
    return '<p style="color: #999; font-size: 12px;">No divisions recorded for this range.</p>';
  }

  const coverageNeeded = divisions.filter(div => summary[div] <= 2);
  const covered = divisions.filter(div => summary[div] > 2);

  const buildRow = (div, isLow) => `
    <div style="display: flex; justify-content: space-between; align-items: center; border: 1px solid ${isLow ? '#ef9a9a' : '#ddd'}; padding: 6px 8px; background: ${isLow ? '#ffebee' : '#fafafa'}; border-radius: 3px;">
      <span style="font-size: 12px; font-weight: ${isLow ? '700' : '600'}; color: ${isLow ? '#c62828' : '#333'}; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${div}</span>
      <span style="font-size: 12px; font-weight: ${isLow ? '700' : 'bold'}; margin-left: 12px; color: ${isLow ? '#c62828' : '#1b5e20'}; white-space: nowrap;">${summary[div]}</span>
    </div>
  `;

  let html = '<div style="display: flex; flex-direction: column; gap: 10px;">';

  if (coverageNeeded.length) {
    html += `
      <div>
        <div style="font-size: 12px; font-weight: 700; color: #b71c1c; margin-bottom: 4px;">Coverage Needed (<= 2)</div>
        <div style="display: flex; flex-direction: column; gap: 6px;">
          ${coverageNeeded.map(div => buildRow(div, true)).join('')}
        </div>
      </div>`;
  }

  if (covered.length) {
    html += `
      <div>
        ${coverageNeeded.length ? '<div style="font-size: 12px; font-weight: 600; color: #555; margin-bottom: 4px;">Covered</div>' : ''}
        <div style="display: flex; flex-direction: column; gap: 6px;">
          ${covered.map(div => buildRow(div, false)).join('')}
        </div>
      </div>`;
  }

  html += '</div>';
  return html;
}



/**
 * Calculate daily summary
 */
function calculateDailySummary(entries) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);

  const threeDaysAgo = new Date(today);
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);

  const todayCount = entries.filter(e => {
    const entryDate = new Date(e.dateAdded);
    entryDate.setHours(0, 0, 0, 0);
    return entryDate.getTime() === today.getTime();
  }).length;

  const yesterdayCount = entries.filter(e => {
    const entryDate = new Date(e.dateAdded);
    entryDate.setHours(0, 0, 0, 0);
    return entryDate.getTime() === yesterday.getTime();
  }).length;

  const threeDaysCount = entries.filter(e => {
    const entryDate = new Date(e.dateAdded);
    entryDate.setHours(0, 0, 0, 0);
    return entryDate.getTime() === threeDaysAgo.getTime();
  }).length;

  return { todayCount, yesterdayCount, threeDaysCount };
}

/**
 * Build daily summary HTML
 */
function buildDailySummaryHTML(summary) {
  return `
    <div style="display: flex; flex-direction: column; gap: 8px;">
      <div style="display: flex; justify-content: space-between; align-items: center; background: #e8f5e9; padding: 10px 12px; border-radius: 4px; white-space: nowrap;">
        <span style="font-size: 12px; color: #2e7d32; font-weight: 600;">Today</span>
        <span style="font-size: 18px; font-weight: bold; color: #2e7d32;">${summary.todayCount}</span>
      </div>
      <div style="display: flex; justify-content: space-between; align-items: center; background: #f3e5f5; padding: 10px 12px; border-radius: 4px; white-space: nowrap;">
        <span style="font-size: 12px; color: #6a1b9a; font-weight: 600;">Yesterday</span>
        <span style="font-size: 18px; font-weight: bold; color: #6a1b9a;">${summary.yesterdayCount}</span>
      </div>
      <div style="display: flex; justify-content: space-between; align-items: center; background: #e3f2fd; padding: 10px 12px; border-radius: 4px; white-space: nowrap;">
        <span style="font-size: 12px; color: #1e88e5; font-weight: 600;">3 Days Ago</span>
        <span style="font-size: 18px; font-weight: bold; color: #1e88e5;">${summary.threeDaysCount}</span>
      </div>
    </div>
  `;
}


/**
 * Build entries table with sorting
 */
function buildEntriesTable(projectKey, entries) {
  const sortedEntries = [...entries].sort((a, b) => b.timestamp - a.timestamp);
  const tableStyle = 'width: 100%; border-collapse: collapse; font-size: 12px; table-layout: fixed;';
  const headerBaseStyle = 'border: 1px solid #ddd; padding: 8px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; text-align: left; cursor: pointer; font-weight: bold;';
  const cellBaseStyle = 'border: 1px solid #ddd; padding: 8px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;';

  return `
    <table style="${tableStyle}">
      <colgroup>
        <col style="width: 140px;">
        <col style="width: 180px;">
        <col style="width: 160px;">
        <col style="width: 120px;">
        <col style="width: 220px;">
        <col style="width: 180px;">
        <col style="width: 160px;">
      </colgroup>
      <thead style="background: #f5f5f5;">
        <tr>
          <th id="sortHeader_${projectKey}_division" style="${headerBaseStyle}">Division</th>
          <th id="sortHeader_${projectKey}_company" style="${headerBaseStyle}">Company</th>
          <th id="sortHeader_${projectKey}_contactName" style="${headerBaseStyle}">Contact</th>
          <th id="sortHeader_${projectKey}_phone" style="${headerBaseStyle}">Phone</th>
          <th id="sortHeader_${projectKey}_email" style="${headerBaseStyle}">Email</th>
          <th id="sortHeader_${projectKey}_sharePointUser" style="${headerBaseStyle}">SharePoint User</th>
          <th id="sortHeader_${projectKey}_dateAdded" style="${headerBaseStyle}">Date - Time</th>
        </tr>
      </thead>
      <tbody>
        ${sortedEntries.map((entry, idx) => `
          <tr style="background: ${idx % 2 === 0 ? '#fff' : '#f9f9f9'};">
            <td style="${cellBaseStyle}">${entry.division}</td>
            <td style="${cellBaseStyle}">${entry.company}</td>
            <td style="${cellBaseStyle}">${entry.contactName}</td>
            <td style="${cellBaseStyle}">${entry.phone}</td>
            <td style="${cellBaseStyle}">${entry.email}</td>
            <td style="${cellBaseStyle}">${entry.sharePointUser}</td>
            <td style="${cellBaseStyle}">${entry.dateAdded} - ${entry.timeAdded}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}


/**
 * Toggle project card expansion
 */
function toggleProjectCard(cardId) {
  const content = $(cardId + '_content');
  const arrow = $(cardId + '_arrow');
  
  if (content) {
    const isHidden = content.style.display === 'none';
    content.style.display = isHidden ? 'block' : 'none';
    if (arrow) {
      arrow.style.transform = isHidden ? 'rotate(90deg)' : 'rotate(0deg)';
    }
  }
}

/**
 * Filter entries by date range
 */
function filterEntriesByDateRange(entries) {
  if (!trackedEntriesCurrentFilter.startDate && !trackedEntriesCurrentFilter.endDate) {
    return entries;
  }

  return entries.filter(entry => {
    const entryDate = new Date(entry.dateAdded);
    
    if (trackedEntriesCurrentFilter.startDate) {
      const startDate = new Date(trackedEntriesCurrentFilter.startDate);
      if (entryDate < startDate) return false;
    }
    
    if (trackedEntriesCurrentFilter.endDate) {
      const endDate = new Date(trackedEntriesCurrentFilter.endDate);
      endDate.setHours(23, 59, 59, 999);
      if (entryDate > endDate) return false;
    }
    
    return true;
  });
}

/**
 * Apply date filter
 */
function applyDateFilter(projectKey) {
  const startInput = $(projectKey + '_startDate');
  const endInput = $(projectKey + '_endDate');
  
  trackedEntriesCurrentFilter = {
    startDate: startInput?.value || null,
    endDate: endInput?.value || null
  };
  
  renderTrackedEntriesUI();
}

/**
 * Clear date filter
 */
function clearDateFilter(projectKey) {
  trackedEntriesCurrentFilter = { startDate: null, endDate: null };
  renderTrackedEntriesUI();
}

/**
 * Sort entries by column
 */
function sortEntriesBy(projectKey, field) {
  console.log(`Sorting by ${field} for project ${projectKey}`);
  // Implementation for sorting
}
