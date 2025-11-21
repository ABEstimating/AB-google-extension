/**
 * Excel Division Summary - Background module
 * Requirements: background.js must importScripts('excel-summary-bg.js')
 * Uses Graph API helpers already defined in background.js: callGraphAPI, getSiteByUrl
 */

(function(){
  if (self.AB_EXCEL_SUMMARY_BG_LOADED) return;
  self.AB_EXCEL_SUMMARY_BG_LOADED = true;
  console.log('[ExcelSummaryBG] loaded');

  // Message handler
  try {
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
      if (!request || typeof request !== 'object') return;

      if (request.type === 'AB_EXCEL_SUMMARY') {
        (async () => {
          try {
            const result = await buildExcelDivisionSummary(request.url || '', request.sheetName || '');
            sendResponse({ success: true, result });
          } catch (err) {
            console.error('[ExcelSummaryBG] summary failed', err);
            sendResponse({ success: false, error: String(err?.message || err) });
          }
        })();
        return true; // keep port open for async
      }
    });
  } catch(e) {
    console.warn('[ExcelSummaryBG] message hook failed', e);
  }

  // === Core ===
  async function buildExcelDivisionSummary(pageUrl, sheetName) {
    if (!pageUrl) throw new Error('Missing pageUrl');
    const ctx = await resolveDriveItemFromExcelUrl(pageUrl);
    if (!ctx || !ctx.driveId || !ctx.itemId || !ctx.siteId) {
      throw new Error('Could not resolve workbook from URL');
    }

    // Create a non-persistent workbook session
    const sessionId = await createSession(ctx.driveId, ctx.itemId);
    try {
      const headers = { 'workbook-session-id': sessionId };

      const targetSheet = await resolveWorksheetName(sheetName, ctx, headers);
      const sheet = encodeURIComponent(targetSheet);

      // Pull a generous rectangle to parse client-side
      const rangeAddress = "A1:ZZ500";
      const range = await callGraphAPI(
        `/drives/${ctx.driveId}/items/${ctx.itemId}/workbook/worksheets('${sheet}')/range(address='${rangeAddress}')`,
        { headers }
      );

      const values = (range && range.values) || [];
      const data = normalizeTable(values);
      logPreviewTable(data);

      const summary = computeSummary(data);
      summary.context = {
        sheetName: targetSheet,
        workbookName: ctx.name || '',
        siteUrl: ctx.siteUrl || ''
      };
      return summary;
    } finally {
      try { await closeSession(ctx.driveId, ctx.itemId, sessionId); } catch(_) {}
    }
  }

  function logPreviewTable(tbl) {
    try {
      const previewRows = Math.min(25, tbl.rows.length);
      for (let r = 0; r < previewRows; r++) {
        const row = tbl.rows[r] || [];
        console.log('[ExcelSummaryBG] row', r + 1, JSON.stringify(row.slice(0, 16)));
      }
    } catch (e) {
      console.warn('[ExcelSummaryBG] preview log failed', e);
    }
  }

  // Try to resolve driveId + itemId from an Excel Online URL
  async function resolveDriveItemFromExcelUrl(pageUrl) {
    const url = new URL(pageUrl);
    const host = url.hostname.toLowerCase();

    const residParam = url.searchParams.get('resid');
    if (residParam) {
      return await resolveDriveItemFromResid(residParam);
    }

    if (!/sharepoint\.com$/i.test(host)) {
      if (host.includes('excel.office.com') || host.includes('officeapps.live.com')) {
        throw new Error('Unable to locate SharePoint file info from Excel Online URL');
      }
      throw new Error('Only SharePoint-hosted workbooks are supported');
    }
    // site path like "/sites/estimating/"
    const sitePath = (() => {
      const m = url.pathname.match(/\/sites\/[^\/]+/i);
      return m ? m[0] : null;
    })();
    if (!sitePath) throw new Error('Site path not found in URL');
    const siteUrlToken = `${host}:${sitePath}`;
    const site = await getSiteByUrl(siteUrlToken);
    const siteId = site && site.id;
    if (!siteId) throw new Error('siteId not resolved');

    const sourcedocParam = url.searchParams.get('sourcedoc') || url.searchParams.get('SourceDoc');
    if (sourcedocParam) {
      const fileName = url.searchParams.get('file') || url.searchParams.get('File');
      const resolved = await resolveBySourceDoc(siteId, sourcedocParam, fileName, `https://${host}${sitePath}`);
      if (resolved) return resolved;
    }

    // Try id= query param first (drive root relative path)
    let relPath = url.searchParams.get('id');
    if (relPath) {
      // id is usually URL-encoded relative path beginning with /sites/.../Shared Documents/...
      try { relPath = decodeURIComponent(relPath); } catch(_) {}
      const driveItem = await callGraphAPI(`/sites/${siteId}/drive/root:${encodeURIComponent(relPath)}?select=id,name,parentReference`);
      return {
        siteId,
        driveId: driveItem?.parentReference?.driveId,
        itemId: driveItem?.id,
        name: driveItem?.name,
        siteUrl: `https://${host}${sitePath}`
      };
    }

    // Fallback: use the visible path after /Shared%20Documents/ or generic /:x:/r/ paths
    const sharedMatch = url.pathname.match(/\/(Shared%20Documents|Documents)\/(.+)$/i);
    if (sharedMatch) {
      const after = decodeURIComponent(sharedMatch[0]).replace(/^\/+/,'');
      const driveItem = await callGraphAPI(`/sites/${siteId}/drive/root:/${encodeURIComponent(after)}?select=id,name,parentReference`);
      return {
        siteId,
        driveId: driveItem?.parentReference?.driveId,
        itemId: driveItem?.id,
        name: driveItem?.name,
        siteUrl: `https://${host}${sitePath}`
      };
    }

    const modernMatch = url.pathname.match(/\/:[^\/]+\/r\/(.+)$/i);
    if (modernMatch && modernMatch[1]) {
      const relative = '/' + modernMatch[1];
      const decoded = decodeURIComponent(relative).replace(/^\/+/,'');
      const driveItem = await callGraphAPI(`/sites/${siteId}/drive/root:/${encodeURIComponent(decoded)}?select=id,name,parentReference`);
      return {
        siteId,
        driveId: driveItem?.parentReference?.driveId,
        itemId: driveItem?.id,
        name: driveItem?.name,
        siteUrl: `https://${host}${sitePath}`
      };
    }

    throw new Error('Could not derive drive path from URL');
  }

  async function resolveBySourceDoc(siteId, sourceDocGuid, fileName, fallbackSiteUrl) {
    if (!siteId || !sourceDocGuid) return null;
    let guid = sourceDocGuid;
    try { guid = decodeURIComponent(guid); } catch (_) {}
    guid = guid.replace(/[{}]/g, '').toUpperCase();
    const query = (fileName || guid).split('/').pop();
    let searchRes;
    try {
      searchRes = await callGraphAPI(`/sites/${siteId}/drive/root/search(q='${encodeURIComponent(query)}')?$select=id,name,parentReference,webUrl,sharepointIds`);
    } catch (_) {
      return null;
    }
    const candidates = Array.isArray(searchRes?.value) ? searchRes.value : [];
    const match = candidates.find(item => {
      const uniqueId = (item?.sharepointIds?.listItemUniqueId || '').replace(/[{}]/g, '').toUpperCase();
      return uniqueId && uniqueId === guid;
    }) || candidates[0];
    if (!match) return null;
    let siteUrl = fallbackSiteUrl || '';
    if (!siteUrl && match.webUrl) {
      try {
        const parsed = new URL(match.webUrl);
        siteUrl = `${parsed.protocol}//${parsed.hostname}`;
      } catch (_) {}
    }
    return {
      siteId,
      driveId: match.parentReference?.driveId,
      itemId: match.id,
      name: match.name,
      siteUrl
    };
  }

  async function resolveDriveItemFromResid(rawResid) {
    if (!rawResid) throw new Error('Missing resid parameter');
    let resid = rawResid;
    try { resid = decodeURIComponent(resid); } catch (_) {}
    resid = resid.replace(/^\{|\}$/g, '');
    const parts = resid.split('!');
    if (parts.length < 2) throw new Error('Invalid resid parameter');
    const driveId = parts[0];
    const itemId = parts.slice(1).join('!');
    if (!driveId || !itemId) throw new Error('Invalid resid parameter');
    const item = await callGraphAPI(`/drives/${driveId}/items/${itemId}?select=id,name,parentReference,webUrl`);
    if (!item || !item.id) throw new Error('Unable to resolve workbook from resid');
    let siteUrl = '';
    let siteId = item?.parentReference?.siteId || null;
    if (!siteId && item?.parentReference?.driveId) {
      siteId = await resolveSiteIdFromDriveId(item.parentReference.driveId);
    }
    if (siteId) {
      siteUrl = await resolveSiteUrlFromSiteId(siteId);
    }
    if (!siteUrl && item.webUrl) {
      try {
        const parsed = new URL(item.webUrl);
        siteUrl = `${parsed.protocol}//${parsed.hostname}`;
      } catch (_) {}
    }
    return {
      driveId,
      itemId,
      siteId,
      name: item.name,
      siteUrl
    };
  }

  async function resolveSiteUrlFromSiteId(siteId) {
    if (!siteId) return '';
    try {
      const site = await callGraphAPI(`/sites/${siteId}?$select=webUrl`);
      return site?.webUrl || '';
    } catch (_) {
      return '';
    }
  }

  async function resolveSiteIdFromDriveId(driveId) {
    if (!driveId) return '';
    try {
      const drive = await callGraphAPI(`/drives/${driveId}?$select=id,name,webUrl,sharePointIds`);
      return drive?.sharePointIds?.siteId || '';
    } catch (_) {
      return '';
    }
  }

  async function createSession(driveId, itemId) {
    const res = await callGraphAPI(`/drives/${driveId}/items/${itemId}/workbook/createSession`, {
      method: 'POST',
      body: { persistChanges: false }
    });
    if (!res?.id) throw new Error('Failed creating workbook session');
    return res.id;
  }

  async function closeSession(driveId, itemId, sessionId) {
    if (!sessionId) return;
    try {
      await callGraphAPI(`/drives/${driveId}/items/${itemId}/workbook/closeSession`, {
        method: 'POST',
        headers: { 'workbook-session-id': sessionId }
      });
    } catch(_) {}
  }

  async function resolveWorksheetName(requestedName, ctx, headers) {
    const list = await callGraphAPI(`/drives/${ctx.driveId}/items/${ctx.itemId}/workbook/worksheets`, { headers });
    const sheets = Array.isArray(list?.value) ? list.value : [];
    console.log('[ExcelSummaryBG] requestedName:', requestedName, 'available sheets:', sheets.map(s => s?.name));
    const normalized = requestedName ? requestedName.trim().toLowerCase() : '';
    if (normalized) {
      const exact = sheets.find(w => String(w?.name || '').trim().toLowerCase() === normalized);
      if (exact) return exact.name;
      const contains = sheets.find(w => {
        const name = String(w?.name || '').trim().toLowerCase();
        return name.includes(normalized) || normalized.includes(name);
      });
      if (contains) return contains.name;
    }
    const priorityOrder = ['DFH','Hi-Speed Overhead','OH Doors','Storefront','Panels','EIFS or Stucco','Drywall','Acoustic','Plumb','HVAC','Elec'];
    for (const keyword of priorityOrder) {
      const match = sheets.find(w => String(w?.name || '').toLowerCase() === keyword.toLowerCase());
      if (match) return match.name;
    }
    const visible = sheets.find(w => w?.visibility !== 'VeryHidden');
    return (visible && visible.name) || (sheets[0] && sheets[0].name) || 'Sheet1';
  }

  // Convert 2D values array to a helper with row access
  function normalizeTable(values) {
    const rows = values.map((r) => Array.isArray(r) ? r : []);
    return {
      rows,
      get(r,c){ return (rows[r] && rows[r][c]) ?? ''; }
    };
  }

  function asNumber(v){
    if (typeof v === 'number') return v;
    if (typeof v === 'string') {
      const t = v.replace(/[\$,()\s]/g,'').replace(/--/g,'-');
      const n = Number(t);
      return Number.isFinite(n) ? n : null;
    }
    return null;
  }

  // Based on user layout: 
  // - company columns start at G (index 6)
  // - header rows G3:Z6 contain company/title; base bid row is row 7 (index 6)
  // - "TOTAL BASE BID" row in column A must be found by text match
  // - scope rows start ~ row 8; include only rows where column F (index 5) > 0
  function computeSummary(tbl){
    const R = tbl.rows.length;
    const C = R ? tbl.rows[0].length : 0;
    const colScope = 0;
    const colSf = 1;
    const colQty = 2;
    const colUnit = 3;
    const colCost = 4; // Column E
    const colFilter = 5; // Column F (Total Cost / filter column)
    const firstCompanyCol = 6; // Column G onward

    let dataStartRow = 8;
    for (let r = 0; r < R; r++) {
      const cell = String(tbl.get(r, colScope) || '').trim().toUpperCase();
      if (cell === 'SCOPE') {
        dataStartRow = r + 1;
        if (String(tbl.get(dataStartRow, colScope) || '').trim().toUpperCase() === 'SUPPLY') {
          dataStartRow += 1;
        }
        break;
      }
    }
    dataStartRow = Math.max(dataStartRow, 8);
    const headerRow = 2;

    let totalRow = -1;
    for (let r = 0; r < R; r++) {
      const value = String(tbl.get(r, colScope) || '').trim().toUpperCase();
      if (value.includes('TOTAL') && value.includes('BASE') && value.includes('BID')) {
        totalRow = r;
        break;
      }
    }

    const companies = [];
    for (let c = firstCompanyCol; c < Math.min(C, firstCompanyCol + 52); c++) {
      const label = detectCompanyLabel(tbl, c);
      if (label) {
        companies.push({ col: c, name: label });
      }
    }

    const items = [];
    for (let r = dataStartRow; r < R; r++) {
      const name = String(tbl.get(r, colScope) || '').trim();
      if (!name) continue;
      const upperName = name.toUpperCase();
      if (upperName.includes('TOTAL') && upperName.includes('BASE') && upperName.includes('BID')) {
        totalRow = r;
        break;
      }

      const filterVal = asNumber(tbl.get(r, colFilter));
      const costVal = asNumber(tbl.get(r, colCost));
      const hasPositive = (filterVal && filterVal > 0) || (costVal && costVal > 0);
      let hasCompanyBid = false;
      for (let i = 0; i < companies.length; i++) {
        const bidValue = tbl.get(r, companies[i].col);
        if (bidValue != null && String(bidValue).trim() !== '') {
          hasCompanyBid = true;
          break;
        }
      }
      if (!hasPositive && !hasCompanyBid) continue;

      const row = {
        name,
        sf: tbl.get(r, colSf) || '',
        qty: tbl.get(r, colQty) || '',
        unit: tbl.get(r, colUnit) || '',
        cost: costVal != null ? costVal : (tbl.get(r, colCost) || ''),
        bids: {}
      };
      companies.forEach(co => { row.bids[co.name] = tbl.get(r, co.col) || ''; });
      items.push(row);
    }

    const baseBids = {};
    if (totalRow >= 0) {
      companies.forEach(co => {
        baseBids[co.name] = tbl.get(totalRow, co.col) || '';
      });
    } else {
      companies.forEach(co => { baseBids[co.name] = ''; });
    }

    return { companies, items, baseBids };
  }

  function detectCompanyLabel(tbl, col) {
    const R = tbl.rows.length;
    const limit = Math.min(R, 15);
    for (let r = 0; r < limit; r++) {
      const text = String(tbl.get(r, col) || '').trim();
      if (!text) continue;
      const lower = text.toLowerCase();
      if (lower.includes('estimating') || lower.includes('cell')) continue;
      if (/[@]/.test(text)) continue;
      if (/[$]/.test(text)) continue;
      if (/^[0-9()\-\s]+$/.test(text)) continue;
      return text;
    }
    return '';
  }

})();
