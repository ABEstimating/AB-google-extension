// Content script to inject floating button and resizable panel with message bridging
(function() {
  'use strict';

  if (!window.__abSheetDebugListener) {
    window.__abSheetDebugListener = true;
    try {
      window.addEventListener('message', function (event) {
        try {
          if (!event || !event.data) return;
          const data = event.data;
          if (typeof data === 'string') {
            if (/sheet/i.test(data)) {
              console.log('[ExcelSummaryUI] frame message (string)', data.slice(0, 200));
            }
            return;
          }
          if (typeof data === 'object') {
            const keys = Object.keys(data || {});
            if (keys.some(k => /sheet/i.test(k) || /tab/i.test(k))) {
              console.log('[ExcelSummaryUI] frame message (object)', JSON.stringify(data).slice(0, 500));
            }
          }
        } catch (err) {
          console.warn('[ExcelSummaryUI] frame message inspect failed', err);
        }
      }, false);
    } catch (err) {
      console.warn('[ExcelSummaryUI] unable to attach frame message logger', err);
    }
  }

  let panelOpen = false;
  let panelDetached = false;
  let detachedPromptShown = false;
  let panel = null;
  let panelContainer = null;
  let floatingButton = null;
  let panelWidth = Math.max(360, Math.round((window.innerWidth || document.documentElement.clientWidth || 1200) * 0.33)); // 1/3 of screen, min 360
  let isResizing = false;
  let currentEmailData = {
    subject: '',
    sender: '',
    senderName: '',
    body: '',
    bodyHtml: '',
    attachments: [],
    sentAt: '',
    messageId: ''
  };

  // Recompute panel width on viewport changes
  function recomputePanelWidth() {
    const vw = window.innerWidth || document.documentElement.clientWidth || 1200;
    panelWidth = Math.max(360, Math.round(vw * 0.33));
    if (!panelOpen) {
      return;
    }

    const host = ensurePanelHost();
    if (!host) {
      return;
    }

    const widthValue = `${panelWidth}px`;
    host.style.width = widthValue;
    host.style.pointerEvents = 'auto';
    if (document.body) {
      document.body.style.marginRight = widthValue;
    }
  }

  // Update width on window resize
  function sendDetachButtonVisibility(show) {
    const iframe = document.getElementById('ab-panel-iframe');
    if (iframe && iframe.contentWindow) {
      iframe.contentWindow.postMessage({ type: 'AB_PANEL_DETACH_BUTTON', show }, '*');
    }
  }

  window.addEventListener('resize', () => {
    // lightweight throttle
    if (window.__ab_resize_ticking) return;
    window.__ab_resize_ticking = true;
    requestAnimationFrame(() => {
      recomputePanelWidth();
      window.__ab_resize_ticking = false;
    });
  });


  // Setup message bridge between iframe and background script
  // Setup message bridge between iframe and background script
function setupMessageBridge() {
  console.log('[Content] Setting up message bridge');

  window.addEventListener('message', async (event) => {
    const iframeWin = document.getElementById('ab-panel-iframe')?.contentWindow;
    if (event.source !== iframeWin) return;

    const data = event.data || {};

    if (data.type === 'AB_PANEL_CLOSE') {        // click on the white X
      closePanel();
      return;
    }

    if (data.type === 'AB_PANEL_DETACH') {
      requestDetach();
      return;
    }

    if (data.type === 'AB_REQUEST_EMAIL_DATA') {
      try { await sendEmailDataToPanel(); } catch (e) { console.error(e); }
      return;
    }

    if (data.type === 'AB_FORWARD_TO_BACKGROUND' && data.message) {
      try {
        const response = await chrome.runtime.sendMessage(data.message);
        event.source.postMessage({
          type: 'AB_RESPONSE',
          messageId: data.messageId,
          originalType: data.message.type,
          response
        }, '*');
      } catch (error) {
        event.source.postMessage({
          type: 'AB_RESPONSE',
          messageId: data.messageId,
          originalType: data.message.type,
          response: { success: false, error: String(error?.message || error) }
        }, '*');
      }
    }
  });
}

  // Initialize message bridge immediately when content script loads
  let messageBridgeInitialized = false;
  if (!messageBridgeInitialized) {
    setupMessageBridge();
    messageBridgeInitialized = true;
    console.log('[Content] Message bridge initialized on page load');
  }

  // Create floating button
 function createFloatingButton() {
  if (document.getElementById('ab-floating-button')) return;


  floatingButton = document.createElement('div');
  floatingButton.id = 'ab-floating-button';

  const logoUrl = chrome.runtime.getURL('logo.svg'); // keep icon size
  const logoImg = document.createElement('img');
  logoImg.src = logoUrl;
  logoImg.alt = 'Open AutoBuilders Panel';
  logoImg.style.width = '25px';
  logoImg.style.height = '25px';
  logoImg.style.display = 'block';
  logoImg.style.pointerEvents = 'none';

  floatingButton.appendChild(logoImg);

  Object.assign(floatingButton.style, {
    position: 'fixed',
    bottom: '45px',
    right: '45px',
    width: '45px',           // smaller circle
    height: '45px',
    borderRadius: '50%',
    background:
    '#0F6CBD',
    border: '0px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    boxShadow:
      '0 10px 22px rgba(0,0,0,.28), 0 3px 6px rgba(0,0,0,.12),' +
      'inset 0 2px 4px rgba(255,255,255,.25), inset 0 -4px 8px rgba(0,0,0,.25)',
    zIndex: 2147483647,
    transition: 'transform .15s ease, box-shadow .15s ease'
  });

  floatingButton.addEventListener('mouseenter', () => {
    floatingButton.style.transform = 'translateY(-1px)';
    floatingButton.style.boxShadow =
      '0 14px 28px rgba(0,0,0,.32), 0 6px 12px rgba(0,0,0,.14),' +
      'inset 0 3px 6px rgba(255,255,255,.28), inset 0 -5px 10px rgba(0,0,0,.28)';
  });
  floatingButton.addEventListener('mouseleave', () => {
    floatingButton.style.transform = 'translateY(0)';
    floatingButton.style.boxShadow =
      '0 10px 22px rgba(0,0,0,.28), 0 3px 6px rgba(0,0,0,.12),' +
      'inset 0 2px 4px rgba(255,255,255,.25), inset 0 -4px 8px rgba(0,0,0,.25)';
  });
  floatingButton.addEventListener('mousedown', () => {
    floatingButton.style.transform = 'translateY(0)';
    floatingButton.style.boxShadow =
      '0 6px 12px rgba(0,0,0,.22), inset 0 1px 3px rgba(255,255,255,.2),' +
      'inset 0 -3px 6px rgba(0,0,0,.35)';
  });
  floatingButton.addEventListener('mouseup', () => {
    floatingButton.style.transform = 'translateY(-1px)';
    floatingButton.style.boxShadow =
      '0 14px 28px rgba(0,0,0,.32), 0 6px 12px rgba(0,0,0,.14),' +
      'inset 0 3px 6px rgba(255,255,255,.28), inset 0 -5px 10px rgba(0,0,0,.28)';
  });

  floatingButton.addEventListener('click', togglePanel);
  document.body.appendChild(floatingButton);
}

// ==== Excel Online detection & Summary popup ====
function isExcelOnlinePage() {
  const h = location.hostname.toLowerCase();
  if (h.includes('sharepoint.com')) return /\/_layouts\/|\/doc\.aspx|\.xlsx/i.test(location.href);
  if (h.includes('excel.office.com')) return true;
  return false;
}
function getVisibleSheetName() {
  const tab = document.querySelector('[aria-selected="true"], [role="tab"], .tab-active, [class*="tab-active"], [data-tab-selected="true"]');
  if (tab && tab.textContent) return tab.textContent.trim();
  const titles = document.querySelectorAll('input[aria-label="Sheet name"], input[aria-label="Worksheet name"]');
  if (titles && titles.length === 1 && titles[0].value) {
    return titles[0].value.trim();
  }
  try {
    const ribbon = document.querySelector('[data-sheet-tab-name][data-tab-selected="true"], [data-sheet-tab-name]');
    if (ribbon && ribbon.getAttribute('data-sheet-tab-name')) {
      return ribbon.getAttribute('data-sheet-tab-name').trim();
    }
  } catch (_) {}
  try {
    const hash = location.hash || '';
    const hashMatch = hash.match(/sheet=([^&]+)/i);
    if (hashMatch) {
      return decodeURIComponent(hashMatch[1]).trim();
    }
    const pathMatch = location.pathname.match(/Sheet=([^&]+)/i);
    if (pathMatch) {
      return decodeURIComponent(pathMatch[1]).trim();
    }
  } catch (_) {}
  if (document.title) {
    const match = document.title.match(/-\s*([^\-]+)$/);
    if (match) return match[1].trim();
  }
  return '';
}
async function openExcelDivisionSummary() {
  const sheetName = getVisibleSheetName();
  try { console.log('[ExcelSummaryUI] sending summary for sheet:', sheetName || '(none)'); } catch (_) {}
  const resp = await chrome.runtime.sendMessage({ type: 'AB_EXCEL_SUMMARY', url: location.href, sheetName });
  if (!resp || !resp.success) { alert('Summary failed: ' + (resp && resp.error ? resp.error : 'Unknown error')); return; }
  showSummaryOverlay(resp.result);
}
function showSummaryOverlay(summary) {
  const existed = document.getElementById('ab-excel-summary-overlay'); if (existed) existed.remove();
  const overlay = document.createElement('div');
  Object.assign(overlay.style,{position:'fixed',inset:'0',background:'rgba(0,0,0,.35)',zIndex:2147483646,display:'flex',alignItems:'center',justifyContent:'center'});
  const panel = document.createElement('div');
  Object.assign(panel.style,{background:'#fff',borderRadius:'10px',width:'90%',maxWidth:'1100px',maxHeight:'80vh',overflow:'auto',boxShadow:'0 12px 28px rgba(0,0,0,.25)'});
  const header = document.createElement('div'); header.textContent=(summary.context?.sheetName||'Division')+' — Summary';
  Object.assign(header.style,{padding:'10px 14px',fontWeight:'600',background:'#0F6CBD',color:'#fff'});
  const close=document.createElement('button'); close.textContent='Close';
  Object.assign(close.style,{marginLeft:'auto',background:'rgba(255,255,255,0.2)',color:'#fff',border:0,padding:'6px 10px',borderRadius:'6px',cursor:'pointer',float:'right'});
  close.onclick=()=>overlay.remove(); header.appendChild(close);
  const tableWrap=document.createElement('div'); tableWrap.style.padding='12px'; tableWrap.style.overflow='auto';
  const table=document.createElement('table'); table.style.width='100%'; table.style.borderCollapse='collapse'; table.style.fontSize='12px'; tableWrap.appendChild(table);
  function th(t){const e=document.createElement('th');e.textContent=t;e.style.border='1px solid #ccc';e.style.padding='6px';e.style.background='#f0f0f0';e.style.position='sticky';e.style.top='0';return e;}
  function td(t){const e=document.createElement('td');e.textContent=t==null?'':String(t);e.style.border='1px solid #ddd';e.style.padding='6px';return e;}
  const thead=document.createElement('thead'); const trh=document.createElement('tr');
  const scopeHeader = (summary.context?.sheetName ? summary.context.sheetName + ' Scope Items' : 'Scope Items');
  [scopeHeader,'SF','QTY','UNIT','Cost'].forEach(h=>trh.appendChild(th(h)));
  (summary.companies||[]).forEach(c=>trh.appendChild(th(c.name))); thead.appendChild(trh); table.appendChild(thead);
  const tbody=document.createElement('tbody');
  (summary.items||[]).forEach(r=>{const tr=document.createElement('tr'); tr.appendChild(td(r.name)); tr.appendChild(td(r.sf)); tr.appendChild(td(r.qty)); tr.appendChild(td(r.unit)); tr.appendChild(td(r.cost)); (summary.companies||[]).forEach(c=>tr.appendChild(td(r.bids[c.name]||''))); tbody.appendChild(tr);});
  const trTotal=document.createElement('tr'); const tlabel=td('TOTAL BASE BID'); tlabel.style.fontWeight='600';
  trTotal.appendChild(tlabel); trTotal.appendChild(td('')); trTotal.appendChild(td('')); trTotal.appendChild(td('')); trTotal.appendChild(td(''));
  (summary.companies||[]).forEach(c=>{const cell=td(summary.baseBids?.[c.name]||''); cell.style.fontWeight='600'; cell.style.background='#fff8e1'; trTotal.appendChild(cell);}); tbody.appendChild(trTotal);
  table.appendChild(tbody); panel.appendChild(header); panel.appendChild(tableWrap); overlay.appendChild(panel); document.body.appendChild(overlay);
}
// Route AB button click: Excel pages open summary; others keep default panel
const _ab_old_togglePanel = (typeof togglePanel === 'function') ? togglePanel : null;
async function togglePanelOverride() { if (isExcelOnlinePage()) { try { await openExcelDivisionSummary(); } catch(e){ console.error(e);} return; } if (_ab_old_togglePanel) return _ab_old_togglePanel(); }
const _ab_orig_createFloatingButton = createFloatingButton;
createFloatingButton = function() {
  _ab_orig_createFloatingButton();
  const btn = document.getElementById('ab-floating-button');
  if (btn){
    floatingButton = btn;
    try { btn.removeEventListener('click', togglePanel); } catch (_) {}
    btn.addEventListener('click', togglePanelOverride);
  }
};

  function ensurePanelHost() {
    if (panelContainer && document.body?.contains(panelContainer)) {
      return panelContainer;
    }

    let existingHost = document.getElementById('ab-panel-host');
    if (!existingHost) {
      if (!document.body) return null;

      existingHost = document.createElement('div');
      existingHost.id = 'ab-panel-host';
      Object.assign(existingHost.style, {
        position: 'fixed',
        top: '0',
        right: '0',
        height: '100vh',
        width: '0px',
        zIndex: '2147483641',
        overflow: 'hidden',
        pointerEvents: 'none',
        display: 'flex',
        alignItems: 'stretch',
        justifyContent: 'flex-start',
        transition: 'width 0.3s ease',
        backgroundColor: 'transparent'
      });
      document.body.appendChild(existingHost);
    }

    panelContainer = existingHost;
    return panelContainer;
  }

  // Create resizable panel
  function createPanel() {
    if (panel) return;

    const panelHost = ensurePanelHost();
    if (!panelHost) {
      console.warn('[Content] Unable to create panel host (document.body not ready)');
      return;
    }

    panel = document.createElement('div');
    panel.id = 'ab-extension-panel';

    Object.assign(panel.style, {
      width: '100%',
      height: '100%',
      backgroundColor: 'white',
      boxShadow: '-4px 0 12px rgba(0,0,0,0.3)',
      borderLeft: '1px solid #ddd',
      display: 'flex',
      flexDirection: 'column',
      position: 'relative',
      overflow: 'hidden',
      pointerEvents: 'auto'
    });

    const resizeHandle = document.createElement('div');
    resizeHandle.id = 'ab-resize-handle';
    Object.assign(resizeHandle.style, {
      position: 'absolute',
      left: '0',
      top: '0',
      width: '12px',
      height: '100%',
      backgroundColor: 'transparent',
      cursor: 'ew-resize',
      zIndex: '2',
      transform: 'translateX(-50%)',
      touchAction: 'none',
      borderLeft: '1px solid rgba(0, 0, 0, 0.12)'
    });
    resizeHandle.className = 'ab-resize-handle';

    let panelUrl = null;
    try {
      panelUrl = chrome.runtime.getURL('panel.html');
    } catch (error) {
      console.error('[Content] Unable to resolve panel URL (extension context invalidated):', error);
      const errorNotice = document.createElement('div');
      Object.assign(errorNotice.style, {
        padding: '24px',
        color: '#b00020',
        fontSize: '14px',
        textAlign: 'center'
      });
      errorNotice.textContent = 'AutoBuilders extension reloaded. Refresh this tab to reopen the panel.';
      panel.appendChild(errorNotice);
      panelHost.appendChild(panel);
      panelContainer = panelHost;
      panelContainer.style.width = '0px';
      panelContainer.style.pointerEvents = 'none';
      if (document.body) {
        document.body.style.marginRight = '0px';
      }
      return;
    }

    const iframe = document.createElement('iframe');
    iframe.src = panelUrl;
    iframe.id = 'ab-panel-iframe';
    Object.assign(iframe.style, {
      width: '100%',
      height: '100%',
      border: 'none',
      backgroundColor: 'transparent'
    });

    panel.appendChild(resizeHandle);
    panel.appendChild(iframe);
    panelHost.appendChild(panel);

    panelContainer = panelHost;
    const initialWidth = panelOpen ? `${panelWidth}px` : '0px';
    panelContainer.style.width = initialWidth;
    panelContainer.style.pointerEvents = panelOpen ? 'auto' : 'none';
    if (document.body) {
      document.body.style.marginRight = initialWidth;
    }

    setupResize(resizeHandle);

    iframe.onload = () => {
      console.log('[Content] Panel iframe loaded');
      setTimeout(() => {
        console.log('[Content] Sending READY signal to iframe');
        iframe.contentWindow.postMessage({ type: 'AB_CONTENT_READY' }, '*');
      }, 500);
    };

    console.log('[Content] Panel created');
  }

  // Extract email data from Outlook Web
  function extractEmailData() {
    const data = {
      subject: '',
      sender: '',
      senderName: '',
      body: '',
      bodyHtml: '',
      sentAt: '',
      messageId: ''
    };

    // Try multiple selectors for subject
    const subjectSelectors = [
      '[role="main"] h1[role="heading"]',
      '[data-test-id="message-subject"]',
      'span[title*="Subject:"]',
      'div[aria-label*="Subject"]',
      '.ms-font-xl',
      'h1.hcptT'
    ];

    for (const selector of subjectSelectors) {
      const element = document.querySelector(selector);
      if (element && element.textContent.trim()) {
        data.subject = element.textContent.trim();
        console.log('[Content] Found subject:', data.subject);
        break;
      }
    }

    // Try multiple selectors for sender email
    const senderSelectors = [
      'button[aria-label*="From:"] span[title]',
      'span[aria-label*="From"]',
      'div[role="button"][title*="@"]',
      'button[title*="@"]',
      '.ms-Persona-primaryText',
      'span.lDdSm'
    ];

    for (const selector of senderSelectors) {
      const elements = document.querySelectorAll(selector);
      for (const element of elements) {
        const text = element.title || element.textContent || '';
        // Extract email from text like "John Doe <john@example.com>"
        const emailMatch = text.match(/<([^>]+@[^>]+)>/) || text.match(/([^\s]+@[^\s]+)/);
        if (emailMatch) {
          data.sender = emailMatch[1];
          // Extract name if available
          const nameMatch = text.match(/^([^<]+)</);
          if (nameMatch) {
            data.senderName = nameMatch[1].trim();
          }
          console.log('[Content] Found sender:', data.sender, 'Name:', data.senderName);
          break;
        }
      }
      if (data.sender) break;
    }

    // Try to get email body (for additional context)
    const bodySelectors = [
      '[role="document"]',
      '[aria-label="Message body"]',
      '.rps_68ae',
      'div[dir="ltr"][id*="UniqueMessageBody"]'
    ];

    for (const selector of bodySelectors) {
      const element = document.querySelector(selector);
      if (element) {
        data.body = element.textContent || '';
        data.bodyHtml = element.innerHTML || '';
        break;
      }
    }

    data.sentAt = findEmailTimestamp();
    data.messageId = extractMessageId();

    currentEmailData = data;
    return data;
  }

  // Send email data to panel
  
async function sendEmailDataToPanel() {
    const iframe = document.getElementById('ab-panel-iframe');
    if (!iframe) {
      console.log('[Content] Cannot send email data - iframe not found');
      return;
    }

    const emailData = extractEmailData();

    console.log('[Content] Extracting email data:', emailData);
    console.log('[Content] Found sender email:', emailData.sender);
    console.log('[Content] Found sender name:', emailData.senderName);

    const signatureDetails = parseSignatureDetails(emailData.bodyHtml || emailData.body);
    let attachments = [];
    try {
      attachments = await collectEmailAttachments();
    } catch (error) {
      console.error('[Content] Failed to collect attachments:', error);
    }

    iframe.contentWindow.postMessage({ 
      type: 'AB_EMAIL_DATA_RESPONSE',
      emailData: {
        from: emailData.sender,
        fromName: emailData.senderName,
        subject: emailData.subject,
        body: emailData.body,
        bodyHtml: emailData.bodyHtml,
        sentAt: emailData.sentAt,
        messageId: emailData.messageId,
        attachments,
        signature: signatureDetails
      }
    }, '*');

    console.log('[Content] Email data sent to panel with type: AB_EMAIL_DATA_RESPONSE');
  }





  function findEmailTimestamp() {
    const selectors = [
      'time[datetime]',
      '[data-test-id="message-date"]',
      'span[title*="Sent"]',
      'div[aria-label*="Sent"]'
    ];

    for (const selector of selectors) {
      const element = document.querySelector(selector);
      if (!element) continue;

      const datetime = element.getAttribute('datetime') || element.getAttribute('title') || element.textContent;
      if (datetime && datetime.trim()) {
        return datetime.trim();
      }
    }

    return new Date().toISOString();
  }

  function extractMessageId() {
    const attributeCandidates = ['data-message-id', 'data-item-id', 'data-qa-id'];
    for (const attribute of attributeCandidates) {
      const node = document.querySelector(`[${attribute}]`);
      if (node) {
        const value = node.getAttribute(attribute);
        if (value) return value;
      }
    }
    return '';
  }

  function parseSignatureDetails(body) {
    if (!body) {
      return { company: '', phone: '', name: '', phones: [], emails: [] };
    }

    const lines = body
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);

    if (lines.length === 0) {
      return { company: '', phone: '', name: '', phones: [], emails: [] };
    }

    let signatureStart = lines.findIndex(line => /^(thanks|regards|sincerely|best|thank you)/i.test(line));
    let signatureLines;
    if (signatureStart >= 0 && signatureStart < lines.length - 1) {
      signatureLines = lines.slice(signatureStart + 1, signatureStart + 9);
    } else {
      signatureLines = lines.slice(-8);
    }

    const emails = [];
    const emailRegex = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
    signatureLines.forEach(line => {
      let match;
      while ((match = emailRegex.exec(line))) {
        emails.push(match[1]);
      }
    });

    const phones = [];
    const phoneRegex = /(?:\+?1[\s.-]?)?(?:\(?\d{3}\)?[\s.-]?){2}\d{4}/g;
    signatureLines.forEach(line => {
      let match;
      while ((match = phoneRegex.exec(line))) {
        phones.push(match[0]);
      }
    });

    let name = '';
    for (const line of signatureLines) {
      if (/\d|@/.test(line)) continue;
      const words = line.split(/\s+/).filter(Boolean);
      if (words.length >= 2 && words.length <= 4 && words.every(word => /^[A-Za-z][A-Za-z'\-]*$/.test(word))) {
        name = line;
        break;
      }
    }

    let company = '';
    const companyCandidates = signatureLines.filter(line => line !== name);
    for (const candidate of companyCandidates) {
      if (!candidate) continue;
      if (candidate.includes('@')) continue;
      if (/\d{4}/.test(candidate)) continue;
      if (/^(phone|tel|email|mobile|office|fax)/i.test(candidate)) continue;
      if (/llc|inc|ltd|company|co\.|group|contracting|builders|construction|corp/i.test(candidate)) {
        company = candidate;
        break;
      }
    }

    if (!company) {
      const nameIndex = signatureLines.indexOf(name);
      if (nameIndex >= 0 && nameIndex + 1 < signatureLines.length) {
        const nextLine = signatureLines[nameIndex + 1];
        if (nextLine && !/\d|@/.test(nextLine)) {
          company = nextLine;
        }
      }
    }

    const uniquePhones = [...new Set(phones)];
    const uniqueEmails = [...new Set(emails)];

    return {
      company,
      phone: uniquePhones[0] || '',
      name,
      phones: uniquePhones,
      emails: uniqueEmails
    };
  }

    async function collectEmailAttachments() {
    const attachments = [];
    const selectorCandidates = [
      '[data-is-attachment="true"] a[href]',
      'div[role="listitem"][data-attachment-id] a[href]',
      'a[download]'
    ];

    const nodes = new Set();
    selectorCandidates.forEach(selector => {
      document.querySelectorAll(selector).forEach(node => {
        const anchor = node.closest('a') || node;
        if (anchor && anchor.href) {
          nodes.add(anchor);
        }
      });
    });

    for (const anchor of nodes) {
      const url = anchor.href;
      if (!url) continue;
      const name = anchor.getAttribute('download') || anchor.getAttribute('title') || anchor.textContent.trim() || 'attachment';

      try {
        const response = await fetch(url, { credentials: 'include' });
        if (!response.ok) {
          console.warn('[Content] Attachment fetch failed:', response.status, url);
          continue;
        }
        const blob = await response.blob();
        const arrayBuffer = await blob.arrayBuffer();
        const base64 = arrayBufferToBase64(arrayBuffer);
        const isInline = blob.size < 20480 && blob.type && blob.type.startsWith('image/');

        attachments.push({
          name,
          size: blob.size,
          contentType: blob.type || 'application/octet-stream',
          isInline,
          content: base64
        });
      } catch (error) {
        console.error('[Content] Error downloading attachment:', error);
      }
    }

    return attachments;
  }

  function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const length = bytes.byteLength;
    for (let i = 0; i < length; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

// Monitor for email changes - IMPROVED VERSION  // Monitor for email changes - IMPROVED VERSION
  function setupEmailMonitor() {
    let lastSubject = '';
    let lastSender = '';
    
    function checkAndSendEmailData() {
      const data = extractEmailData();
      
      // Only send if email actually changed
      if (data.subject !== lastSubject || data.sender !== lastSender) {
        console.log('[Content] New email detected:', data.subject, data.sender);
        lastSubject = data.subject;
        lastSender = data.sender;
        currentEmailData = data;
        sendEmailDataToPanel().catch(error => console.error('[Content] Failed to send email data:', error));
      }
    }
    
    // Initial check
    setTimeout(checkAndSendEmailData, 1000);
    
    // Poll every 2 seconds for email changes (more reliable than MutationObserver)
    setInterval(checkAndSendEmailData, 2000);
  }

  // Setup resize functionality
  function setupResize(handle) {
    let startX = 0;
    let startWidth = 0;
    let activePointerId = null;

    const handlePointerMove = (event) => {
      if (!isResizing || event.pointerId !== activePointerId) return;

      event.preventDefault();

      const deltaX = startX - event.clientX;
      const newWidth = Math.max(300, Math.min(1200, startWidth + deltaX));

      panelWidth = newWidth;
      if (panelContainer) {
        const widthValue = `${newWidth}px`;
        panelContainer.style.width = widthValue;
        if (document.body) {
          document.body.style.marginRight = widthValue;
        }
      }
    };

    const stopResize = (event) => {
      if (!isResizing || event.pointerId !== activePointerId) {
        return;
      }

      event.preventDefault();

      isResizing = false;
      activePointerId = null;

      try {
        handle.releasePointerCapture(event.pointerId);
      } catch (error) {
        // ignore release errors
      }

      document.body.style.cursor = '';
      document.body.style.userSelect = '';

      handle.removeEventListener('pointermove', handlePointerMove);
      handle.removeEventListener('pointerup', stopResize);
      handle.removeEventListener('pointercancel', stopResize);
    };

    handle.addEventListener('pointerdown', (event) => {
      if (event.button !== 0 && event.pointerType !== 'touch') {
        return;
      }

      if (isResizing) {
        return;
      }

      event.preventDefault();

      isResizing = true;
      activePointerId = event.pointerId;
      startX = event.clientX;
      const currentWidth = panelContainer
        ? panelContainer.getBoundingClientRect().width
        : panelWidth;
      startWidth = Math.max(0, currentWidth);

      document.body.style.cursor = 'ew-resize';
      document.body.style.userSelect = 'none';

      try {
        handle.setPointerCapture(activePointerId);
      } catch (error) {
        // ignore capture errors
      }

      handle.addEventListener('pointermove', handlePointerMove);
      handle.addEventListener('pointerup', stopResize);
      handle.addEventListener('pointercancel', stopResize);
    });
  }

  function requestDetach() {
    if (panelDetached) {
      if (chrome?.runtime?.sendMessage) {
        chrome.runtime.sendMessage({ type: 'AB_PANEL_FOCUS_POPUP' }).catch(() => {});
      }
      return;
    }
    panelDetached = true;
    closePanel();
    sendDetachButtonVisibility(false);
    if (chrome?.runtime?.sendMessage) {
      chrome.runtime.sendMessage({ type: 'AB_PANEL_DETACH_REQUEST' }).catch(error => {
        console.error('[Content] Detach request failed', error);
      });
    }
    if (!detachedPromptShown) {
      detachedPromptShown = true;
      setTimeout(() => {
        alert('AutoBuilders panel detached. Close the popup window to reattach it to Outlook.');
      }, 0);
    }
  }

  // Toggle panel open/close
  function togglePanel() {
    if (panelDetached) {
      if (chrome?.runtime?.sendMessage) {
        chrome.runtime.sendMessage({ type: 'AB_PANEL_FOCUS_POPUP' }).catch(() => {});
      }
      return;
    }
    if (panelOpen) {
      closePanel();
    } else {
      openPanel();
    }
  }

  // Open panel
function openPanel() {
  if (panelDetached) {
    if (chrome?.runtime?.sendMessage) {
      chrome.runtime.sendMessage({ type: 'AB_PANEL_FOCUS_POPUP' })
        .then(response => {
          if (!response || response.success !== true) {
            panelDetached = false;
            openPanel();
          }
        })
        .catch(() => {
          panelDetached = false;
          openPanel();
        });
    }
    return;
  }
  if (!panel) createPanel();
  panelOpen = true;
  recomputePanelWidth();

  if (panelContainer) {
    panelContainer.style.pointerEvents = 'auto';
  }

  if (floatingButton && floatingButton.style) {
    floatingButton.style.display = 'none';
  }

  setTimeout(() => {
    sendEmailDataToPanel().catch(err => console.error(err));
    sendDetachButtonVisibility(true);
  }, 500);
}

function closePanel() {
  if (!panelContainer) return;

  panelOpen = false;
  panelContainer.style.width = '0px';
  panelContainer.style.pointerEvents = 'none';
  if (document.body) {
    document.body.style.marginRight = '0px';
  }

  if (!document.getElementById('ab-floating-button')) {
    createFloatingButton();
  } else if (floatingButton && floatingButton.style) {
    floatingButton.style.display = 'flex';
    floatingButton.style.right = '20px';
  }
}

if (chrome?.runtime?.onMessage) {
    chrome.runtime.onMessage.addListener((message) => {
      if (!message || typeof message.type !== 'string') {
        return;
      }
      if (message.type === 'AB_PANEL_DETACH_CONFIRMED') {
        panelDetached = true;
        closePanel();
        sendDetachButtonVisibility(false);
        return;
      }
      if (message.type === 'AB_PANEL_ATTACH') {
        panelDetached = false;
        openPanel();
        sendDetachButtonVisibility(true);
        return;
      }
    });
  }

  // Initialize immediately
  console.log('[Content] AB Extension content script loaded');
  console.log('[Content] Creating floating button immediately');
  
  // Create button right away
 if (document.body) {
  createFloatingButton();
  setupEmailMonitor();
} else {
  const waitForBody = setInterval(() => {
    if (document.body) {
      clearInterval(waitForBody);
      createFloatingButton();
      setupEmailMonitor();
    }
  }, 100);
}

})();
