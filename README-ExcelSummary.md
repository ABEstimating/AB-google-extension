# Excel Division Summary Popup (Web Excel)

Purpose: when a SharePoint Web Excel workbook is open, the AB floating button opens a compact popup that shows a filtered summary of the active Division sheet:
- Contacts from columns G.. onward become company columns.
- Only scope rows whose **column F** value is a number > 0 are listed.
- The **TOTAL BASE BID** row is detected by text in column A to grab the totals.
- The table mirrors your screenshot headers (Scope Item, SF, QTY, UNIT, Cost, AB Estimate if present, then each company).

## Files
- `excel-summary-bg.js` – background worker using Microsoft Graph to read the sheet.
- `README-ExcelSummary.md` – this document.

## Minimal integration
1) **background.js**: near the other `importScripts(...)` lines at the top, add:
```js
try { importScripts('excel-summary-bg.js'); } catch(e) { console.warn('excel-summary not loaded', e); }
```

2) **content.js**: add the following helpers and click handler override.
Insert after your current `createFloatingButton()` definition, before first usage of `togglePanel`:

```js
// ==== Excel Online detection & Summary popup ====
function isExcelOnlinePage() {
  const host = location.hostname;
  const h = host.toLowerCase();
  if (h.includes('sharepoint.com')) {
    return /\/_layouts\/|\/doc\.aspx|\.xlsx/i.test(location.href);
  }
  if (h.includes('excel.office.com')) return true;
  return false;
}

function getVisibleSheetName() {
  // Excel web shows the selected sheet tab with aria-selected="true"
  const tab = document.querySelector('button[role="tab"][aria-selected="true"], div[role="tab"][aria-selected="true"]');
  if (tab && tab.textContent) return tab.textContent.trim();
  // Fallback from title or URL
  return '';
}

async function openExcelDivisionSummary() { 
  const sheetName = getVisibleSheetName();
  const resp = await chrome.runtime.sendMessage({
    type: 'AB_EXCEL_SUMMARY',
    url: location.href,
    sheetName
  });
  if (!resp || !resp.success) {
    alert('Summary failed: ' + (resp && resp.error ? resp.error : 'Unknown error'));
    return;
  }
  showSummaryOverlay(resp.result);
}

function showSummaryOverlay(summary) {
  // Basic modal with table
  const existed = document.getElementById('ab-excel-summary-overlay');
  if (existed) existed.remove();

  const overlay = document.createElement('div');
  overlay.id = 'ab-excel-summary-overlay';
  Object.assign(overlay.style, {
    position: 'fixed', inset: '0', background: 'rgba(0,0,0,0.35)', zIndex: 2147483646,
    display: 'flex', alignItems: 'center', justifyContent: 'center'
  });

  const panel = document.createElement('div');
  Object.assign(panel.style, {
    background: '#fff', borderRadius: '10px', width: '90%', maxWidth: '1100px',
    maxHeight: '80vh', overflow: 'auto', boxShadow: '0 12px 28px rgba(0,0,0,.25)'
  });

  const header = document.createElement('div');
  header.textContent = (summary.context?.sheetName || 'Division') + ' — Summary';
  Object.assign(header.style, { padding: '10px 14px', fontWeight: '600', background: '#0F6CBD', color: '#fff' });

  const close = document.createElement('button');
  close.textContent = 'Close';
  Object.assign(close.style, { marginLeft: 'auto', background: 'rgba(255,255,255,0.2)', color:'#fff', border:0, padding:'6px 10px', borderRadius:'6px', cursor:'pointer', float:'right' });
  close.onclick = () => overlay.remove();
  header.appendChild(close);

  const tableWrap = document.createElement('div');
  tableWrap.style.padding = '12px';
  tableWrap.style.overflow = 'auto';

  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';
  table.style.fontSize = '12px';
  tableWrap.appendChild(table);

  function th(txt){ const el = document.createElement('th'); el.textContent = txt; el.style.border='1px solid #ccc'; el.style.padding='6px'; el.style.background='#f0f0f0'; el.style.position='sticky'; el.style.top='0'; return el; }
  function td(txt){ const el = document.createElement('td'); el.textContent = txt==null?'':String(txt); el.style.border='1px solid #ddd'; el.style.padding='6px'; return el; }

  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  ['Storefront Scope Item','SF','QTY','UNIT','Cost'].forEach(h => trh.appendChild(th(h)));
  (summary.companies || []).forEach(c => trh.appendChild(th(c.name)));
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  (summary.items || []).forEach(row => {
    const tr = document.createElement('tr');
    tr.appendChild(td(row.name));
    tr.appendChild(td(row.sf));
    tr.appendChild(td(row.qty));
    tr.appendChild(td(row.unit));
    tr.appendChild(td(row.cost));
    (summary.companies || []).forEach(c => tr.appendChild(td(row.bids[c.name] || '')));
    tbody.appendChild(tr);
  });

  // Total Base Bid row
  const trTotal = document.createElement('tr');
  const totalLabel = td('TOTAL BASE BID'); totalLabel.style.fontWeight='600';
  trTotal.appendChild(totalLabel);
  trTotal.appendChild(td('')); trTotal.appendChild(td('')); trTotal.appendChild(td('')); trTotal.appendChild(td(''));
  (summary.companies || []).forEach(c => {
    const cell = td(summary.baseBids?.[c.name] || '');
    cell.style.fontWeight = '600';
    cell.style.background = '#fff8e1';
    trTotal.appendChild(cell);
  });
  tbody.appendChild(trTotal);

  table.appendChild(tbody);

  panel.appendChild(header);
  panel.appendChild(tableWrap);
  overlay.appendChild(panel);
  document.body.appendChild(overlay);
}

// Override floating button behavior: Excel pages open summary; others keep existing togglePanel
const _ab_old_togglePanel = (typeof togglePanel === 'function') ? togglePanel : null;
async function togglePanelOverride() {
  if (isExcelOnlinePage()) {
    try { await openExcelDivisionSummary(); } catch(e){ console.error(e); }
    return;
  }
  if (_ab_old_togglePanel) return _ab_old_togglePanel();
}

// Attach override after floating button is created
const _ab_orig_createFloatingButton = createFloatingButton;
createFloatingButton = function() {
  _ab_orig_createFloatingButton();
  // Swap click handler to the override
  const btn = document.getElementById('ab-floating-button');
  if (btn) {
    const clone = btn.cloneNode(true);
    btn.replaceWith(clone);
    clone.addEventListener('click', togglePanelOverride);
  }
};
```

This keeps your normal panel everywhere else. On Excel pages it opens the summary.
