
// coverage-bg.js — v2: robust summary/detail for Coverage tab
(function(){
  console.log('[CoverageBG v2] loaded');

  // ===== persistence for hidden divisions =====
  async function setCoverageHidden(itemId, hiddenKeys) {
    try {
      const db = await DB.init();
      const tx = db.transaction(['metadata'], 'readwrite');
      const store = tx.objectStore('metadata');
      await store.put({ key: `coverageHidden:${itemId}`, value: Array.from(new Set(hiddenKeys || [])) });
      return true;
    } catch (e) {
      console.warn('[CoverageBG] setCoverageHidden failed', e);
      return false;
    }
  }
  async function getCoverageHidden(itemId) {
    try {
      const db = await DB.init();
      const tx = db.transaction(['metadata'], 'readonly');
      const store = tx.objectStore('metadata');
      return await new Promise((resolve, reject) => {
        const req = store.get(`coverageHidden:${itemId}`);
        req.onsuccess = () => resolve(req.result?.value || []);
        req.onerror = () => reject(req.error);
      });
    } catch (e) {
      console.warn('[CoverageBG] getCoverageHidden failed', e);
      return [];
    }
  }

  // ===== helpers =====
  function keyOfDivision(div) {
    if (!div) return '';
    if (typeof div === 'string') return div.trim();
    return (div.sheetName || div.name || div.key || '').trim();
  }

  function displayOfDivision(div) {
    if (!div) return '';
    if (typeof div === 'string') return div.trim();
    return (div.displayName || div.name || div.sheetName || div.key || '').trim();
  }

  function sortRows(rows) {
    rows.sort((a, b) => {
      const ac = (a.displayName.match(/^\\d{2}-\\d{3}/) || [])[0] || '99-999';
      const bc = (b.displayName.match(/^\\d{2}-\\d{3}/) || [])[0] || '99-999';
      return ac.localeCompare(bc) || a.displayName.localeCompare(b.displayName);
    });
    return rows;
  }
  async function ensureProjectLoadedByName(name) {
    // if the worker restarted, currentProjectFile may be null
    if (currentProjectFile && currentProjectFile.name === name) return true;
    if (!name || typeof loadProjectWorkbook !== 'function') return false;
    try {
      await loadProjectWorkbook(name, { includeContacts: true, forceRefresh: false, forceContactRefresh: false });
      return true;
    } catch (e) {
      console.warn('[CoverageBG] loadProjectWorkbook failed', e);
      return false;
    }
  }

  async function ensureDivisionMapBuilt() {
    if (!divisionMapping || Object.keys(divisionMapping).length === 0) {
      try {
        await buildDivisionMapping({ project: activeProjectRef || null, runtime: activeRuntimeRef || null, forceRefresh: false, includeContacts: false });
      } catch (e) {
        console.warn('[CoverageBG] buildDivisionMapping failed', e);
      }
    }
  }

  async function ensureContactsLoaded() {
    if (!activeRuntimeRef) activeRuntimeRef = {};
    if (!activeRuntimeRef.contactsByDivision) {
      try {
        const res = await loadDivisionContactsInternal({ project: activeProjectRef || null, runtime: activeRuntimeRef, forceRefresh: false });
        activeRuntimeRef.contactsByDivision = res;
      } catch (e) {
        console.warn('[CoverageBG] loadDivisionContactsInternal failed', e);
      }
    }
    return activeRuntimeRef.contactsByDivision || {};
  }

  function makeSummary(divisions, contactsByDivision, hidden) {
    const hiddenSet = new Set(hidden || []);
    const rows = [];
    for (const div of divisions) {
      const key = keyOfDivision(div);
      if (!key) continue;
      const list = contactsByDivision[key] || [];
      rows.push({
        key,
        displayName: displayOfDivision(div),
        count: Array.isArray(list) ? list.length : 0,
        hidden: hiddenSet.has(key)
      });
    }
    return sortRows(rows);
  }

  // ===== wire messages =====
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || !message.type) return false;

    switch (message.type) {

      case 'AB_GET_COVERAGE_SUMMARY':
        (async () => {
          try {
            // hydrate if panel passed project name
            if (!currentProjectFile && message.projectName) {
              await ensureProjectLoadedByName(message.projectName);
            }
            if (!currentProjectFile) {
              sendResponse({ success: true, rows: [], hidden: [], note: 'NO_ACTIVE_PROJECT' });
              return;
            }

            await ensureDivisionMapBuilt();
            const hidden = await getCoverageHidden(currentProjectFile.id);
            const divisions = (activeRuntimeRef?.divisions) ||
                              (activeProjectRef?.divisions) ||
                              Object.values(divisionMapping || {});
            const contactsByDivision = await ensureContactsLoaded();
            const rows = makeSummary(divisions, contactsByDivision, hidden);
            sendResponse({ success: true, rows, hidden });
          } catch (e) {
            console.error('[CoverageBG] summary error', e);
            sendResponse({ success: false, error: e.message });
          }
        })();
        return true;

      case 'AB_GET_COVERAGE_DIVISION_DETAIL':
        (async () => {
          try {
            if (!currentProjectFile && message.projectName) {
              await ensureProjectLoadedByName(message.projectName);
            }
            const key = message.divisionKey;
            const contactsByDivision = await ensureContactsLoaded();
            const list = (contactsByDivision[key] || []).map(row => {
              if (Array.isArray(row)) {
                return { company: String(row[0]||'), name: String(row[1]||'), phone: String(row[2]||'), email: String(row[3]||') };
              } else {
                return { company: String(row.company||'), name: String(row.name||'), phone: String(row.phone||'), email: String(row.email||') };
              }
            });
            sendResponse({ success: true, rows: list });
          } catch (e) {
            console.error('[CoverageBG] detail error', e);
            sendResponse({ success: false, error: e.message });
          }
        })();
        return true;

      case 'AB_SET_COVERAGE_HIDDEN':
        (async () => {
          try {
            if (!currentProjectFile && message.projectName) {
              await ensureProjectLoadedByName(message.projectName);
            }
            const targetId = currentProjectFile?.id || message.projectName || 'none';
            const ok = await setCoverageHidden(targetId, message.hiddenKeys || []);
            sendResponse({ success: ok });
          } catch (e) {
            sendResponse({ success: false, error: e.message });
          }
        })();
        return true;

      case 'AB_GET_COVERAGE_HIDDEN':
        (async () => {
          try {
            if (!currentProjectFile && message.projectName) {
              await ensureProjectLoadedByName(message.projectName);
            }
            const targetId = currentProjectFile?.id || message.projectName || 'none';
            const hidden = await getCoverageHidden(targetId);
            sendResponse({ success: true, hidden });
          } catch (e) {
            sendResponse({ success: false, error: e.message });
          }
        })();
        return true;
    }

    return false;
  });

})();





