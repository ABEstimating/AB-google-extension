
/* === Coverage Tab: summary + detail === */
(function(){
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initCoverageTab);
  } else {
    initCoverageTab();
  }

  function initCoverageTab(){
    try {
      const firstTab = document.querySelector('.tab');
      const tabBar = firstTab ? firstTab.parentElement : null;
      const firstContent = document.querySelector('.tab-content');
      const contentWrap = firstContent ? firstContent.parentElement : null;
      if (!tabBar || !contentWrap) return;

      // Make button
      if (!document.querySelector('.tab[data-tab="coverage"]')) {
        const btn = document.createElement('button');
        btn.className = 'tab';
        btn.dataset.tab = 'coverage';
        btn.type = 'button';
        btn.textContent = 'Coverage';
        tabBar.appendChild(btn);
      }

      // Make content
      if (!document.getElementById('coverage-tab')) {
        const div = document.createElement('div');
        div.className = 'tab-content';
        div.id = 'coverage-tab';
        div.innerHTML = `
          <style>
            .coverage-toolbar{display:flex;gap:.5rem;align-items:center;margin:8px 0;}
            .coverage-table, .coverage-detail{border-collapse:collapse;width:100%;}
            .coverage-table th, .coverage-table td,
            .coverage-detail th, .coverage-detail td{border:1px solid #ccc;padding:6px 8px;vertical-align:top;}
            .coverage-row{cursor:pointer;}
            .coverage-row:hover{background:#f5f5f5;}
            .coverage-detail{margin:6px 0 12px 24px;}
            .hidden-flag{opacity:.6;}
          </style>
          <div class="coverage-toolbar">
            <button id="cov-refresh" type="button">Refresh</button>
            <button id="cov-hide" type="button">Hide selected</button>
            <button id="cov-unhide" type="button">Unhide all</button>
          </div>
          <table class="coverage-table" id="coverage-table">
            <thead><tr>
              <th style="width:28px;"><input type="checkbox" id="cov-check-all"></th>
              <th>Division</th>
              <th style="width:120px;">Bidders</th>
            </tr></thead>
            <tbody id="coverage-body"></tbody>
          </table>
        `;
        contentWrap.appendChild(div);
      }

      // Wire tab system already present
      if (typeof setupTabs === 'function') setupTabs();

      // Wire actions
      const body = document.getElementById('coverage-body');
      const btnRefresh = document.getElementById('cov-refresh');
      const btnHide = document.getElementById('cov-hide');
      const btnUnhide = document.getElementById('cov-unhide');
      const checkAll = document.getElementById('cov-check-all');

      let lastSummary = [];
      let expandedKey = null;
      let hiddenSet = new Set();

      function renderSummary(rows){
        body.innerHTML = '';
        rows.forEach(row => {
          const tr = document.createElement('tr');
          tr.className = 'coverage-row' + (row.hidden ? ' hidden-flag' : '');
          tr.dataset.key = row.key;

          tr.innerHTML = `
            <td><input type="checkbox" class="cov-check" value="${escapeHtml(row.key)}"></td>
            <td><span class="cov-name">${escapeHtml(row.displayName)}</span></td>
            <td style="text-align:right;"><span class="cov-count">${row.count}</span></td>
          `;
          body.appendChild(tr);

          tr.addEventListener('click', async (e) => {
            if (e.target && (e.target.tagName === 'INPUT' || e.target.closest('input'))) return;
            const key = tr.dataset.key;
            if (expandedKey === key) {
              collapseDetails(key);
              expandedKey = null;
              return;
            }
            // collapse existing
            if (expandedKey) collapseDetails(expandedKey);
            expandedKey = key;
            const detail = await fetchDivisionDetail(key);
            insertDetailsAfter(tr, detail);
          });
        });
      }

      function collapseDetails(key){
        const row = body.querySelector(`tr[data-key="${CSS.escape(key)}"]`);
        const next = row ? row.nextElementSibling : null;
        if (next && next.classList.contains('coverage-detail-row')) {
          next.remove();
        }
      }

      function insertDetailsAfter(tr, rows){
        collapseDetails(tr.dataset.key);
        const detailRow = document.createElement('tr');
        detailRow.className = 'coverage-detail-row';
        const col = document.createElement('td');
        col.colSpan = 3;
        col.innerHTML = `
          <table class="coverage-detail">
            <thead><tr>
              <th>Company</th><th>Contact</th><th>Phone</th><th>Email</th>
            </tr></thead>
            <tbody>
              ${rows.map(r => `
                <tr>
                  <td>${escapeHtml(r.company||'')}</td>
                  <td>${escapeHtml(r.name||'')}</td>
                  <td>${escapeHtml(r.phone||'')}</td>
                  <td>${escapeHtml(r.email||'')}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        `;
        detailRow.appendChild(col);
        tr.after(detailRow);
      }

      function escapeHtml(s){
        return String(s||'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
      }

      function getCheckedKeys(){
        return Array.from(body.querySelectorAll('.cov-check:checked')).map(i => i.value);
      }

      async function fetchSummary(){
        return await new Promise((resolve) => {
          chrome.runtime.sendMessage({ type: 'ABX_GET_COVERAGE_SUMMARY', projectName: (window.currentProject || null) }, (res) => {
            if (!res || !res.success) { resolve({ rows: [], hidden: [] }); return; }
            resolve(res);
          });
        });
      }

      async function fetchDivisionDetail(key){
        return await new Promise((resolve) => {
          chrome.runtime.sendMessage({ type: 'ABX_GET_COVERAGE_DIVISION_DETAIL', projectName: (window.currentProject || null), divisionKey: key }, (res) => {
            if (!res || !res.success) { resolve([]); return; }
            resolve(res.rows || []);
          });
        });
      }

      async function saveHidden(keys){
        return await new Promise((resolve) => {
          chrome.runtime.sendMessage({ type: 'ABX_SET_COVERAGE_HIDDEN', hiddenKeys: Array.from(new Set(keys)) }, (res) => resolve(!!res?.success));
        });
      }

      btnRefresh.addEventListener('click', async () => {
        const { rows, hidden } = await fetchSummary();
        hiddenSet = new Set(hidden || []);
        lastSummary = rows;
        renderSummary(rows);
      });

      btnHide.addEventListener('click', async () => {
        const keys = getCheckedKeys();
        keys.forEach(k => hiddenSet.add(k));
        await saveHidden(Array.from(hiddenSet));
        const { rows, hidden } = await fetchSummary();
        hiddenSet = new Set(hidden || []);
        lastSummary = rows;
        renderSummary(rows);
      });

      btnUnhide.addEventListener('click', async () => {
        hiddenSet = new Set();
        await saveHidden([]);
        const { rows } = await fetchSummary();
        lastSummary = rows;
        renderSummary(rows);
      });

      checkAll.addEventListener('change', () => {
        const on = checkAll.checked;
        body.querySelectorAll('.cov-check').forEach(ch => ch.checked = on);
      });

      // initial load if the tab is current, else load when clicked
      async function initialIfActive(){
        const covBtn = document.querySelector('.tab[data-tab="coverage"]');
        const covPanel = document.getElementById('coverage-tab');
        if (covBtn && covPanel && covBtn.classList.contains('active')) {
          const { rows } = await fetchSummary();
          lastSummary = rows;
          renderSummary(rows);
        } else if (covBtn) {
          covBtn.addEventListener('click', async () => {
            const { rows } = await fetchSummary();
            lastSummary = rows;
            renderSummary(rows);
          }, { once: true });
        }
      }
      initialIfActive();

    } catch (e) {
      console.warn('[CoverageTab] init failed', e);
    }
  }
})();
/* === END Coverage Tab === */
