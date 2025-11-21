console.log('[AB] campaigns-bg loaded');
self.AB_CAMPAIGNS_BG_LOADED = true;
self.AB_CAMPAIGN_BG_LOADED = true;
/* campaigns-bg.js — background worker add-on for campaign automation */

(() => {
  const STORAGE_KEY = 'ab_campaigns';
  const POLL_ALARM = 'AB_CMP_POLL';
  // ===== SharePoint sync: /sites/estimating -> Documents drive =====
  const SITE_HOSTNAME = 'autobuildersbpos.sharepoint.com';
  const SITE_PATH = '/sites/estimating';
  const CMP_ROOT = 'Z - Extension Cache/campaigns';
  const DRIVE_CACHE_KEY = 'ab_drive_estimating';

  const SCHEDULE_MIN_LEAD_MS = 60 * 1000;

  const FOLLOW_UP_KEYS = ['rfi', 'bids'];
  const RECIPIENT_STATUS = {
    pending: 'pending',
    sent: 'sent',
    responded: 'responded',
    removed: 'removed'
  };

  async function callGraph(path, options = {}) {
    // Expect host background to expose callGraphAPI. If not, throw.
    if (typeof self.callGraphAPI !== 'function') throw new Error('callGraphAPI not available');
    return await self.callGraphAPI(path, options);
  }
  async function getEstimatingDriveId() {
    const cached = (await chrome.storage.local.get([DRIVE_CACHE_KEY]))[DRIVE_CACHE_KEY];
    if (cached) return cached;
    const site = await callGraph(`/sites/${SITE_HOSTNAME}:${encodeURIComponent(SITE_PATH)}?$select=id`);
    const siteId = site && site.id;
    if (!siteId) throw new Error('cannot resolve site id');
    const drive = await callGraph(`/sites/${siteId}/drive?$select=id,name`);
    const driveId = drive && drive.id;
    if (!driveId) throw new Error('cannot resolve drive id');
    await chrome.storage.local.set({ [DRIVE_CACHE_KEY]: driveId });
    return driveId;
  }
  async function spEnsureFolder(driveId, folderPath) {
    const segs = folderPath.split('/').filter(Boolean);
    let curr = '';
    for (const s of segs) {
      const next = curr ? `${curr}/${s}` : s;
      try {
        await callGraph(`/drives/${driveId}/root:/${encodeURIComponent(next)}`);
      } catch {
        await callGraph(`/drives/${driveId}/root:/${encodeURIComponent(curr || '')}:/children`, {
          method: 'POST',
          body: { name: s, folder: {}, '@microsoft.graph.conflictBehavior': 'replace' }
        });
      }
      curr = next;
    }
  }
  async function saveCampaignRemote(cmp) {
    try {
      const driveId = await getEstimatingDriveId();
      await spEnsureFolder(driveId, CMP_ROOT);
      
      // Load existing campaigns from single file
      let allCampaigns = {};
      try {
        const existingFile = await callGraph(`/drives/${driveId}/root:/${encodeURIComponent(CMP_ROOT)}/all-campaigns.json`);
        if (existingFile && existingFile['@microsoft.graph.downloadUrl']) {
          const resp = await fetch(existingFile['@microsoft.graph.downloadUrl']);
          if (resp.ok) {
            allCampaigns = await resp.json();
          }
        }
      } catch (_) {
        // File doesn't exist yet, start with empty object
      }
      
      // Update with current campaign
      allCampaigns[cmp.id] = cmp;
      
      // Save back to single file
      await callGraph(`/drives/${driveId}/root:/${encodeURIComponent(CMP_ROOT)}/all-campaigns.json:/content?@microsoft.graph.conflictBehavior=replace`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: allCampaigns
      });
    } catch (e) {
      console.warn('[AB CMP] remote save failed', e && e.message || e);
    }
  }
  async function listRemoteCampaigns() {
    try {
      const driveId = await getEstimatingDriveId();
      
      // Read from single all-campaigns.json file
      try {
        const item = await callGraph(`/drives/${driveId}/root:/${encodeURIComponent(CMP_ROOT)}/all-campaigns.json`);
        if (item && item['@microsoft.graph.downloadUrl']) {
          const resp = await fetch(item['@microsoft.graph.downloadUrl']);
          if (resp.ok) {
            const allCampaigns = await resp.json();
            if (allCampaigns && typeof allCampaigns === 'object') {
              return Object.values(allCampaigns).filter(c => c && c.id);
            }
          }
        }
      } catch (_) {
        // File doesn't exist yet
      }
      
      return [];
    } catch (e) {
      console.warn('[AB CMP] list remote failed', e && e.message || e); 
      return [];
    }
  }
  async function deleteCampaignRemote(campaignId) {
    try {
      const driveId = await getEstimatingDriveId();
      
      // Load existing campaigns from single file
      let allCampaigns = {};
      try {
        const existingFile = await callGraph(`/drives/${driveId}/root:/${encodeURIComponent(CMP_ROOT)}/all-campaigns.json`);
        if (existingFile && existingFile['@microsoft.graph.downloadUrl']) {
          const resp = await fetch(existingFile['@microsoft.graph.downloadUrl']);
          if (resp.ok) {
            allCampaigns = await resp.json();
          }
        }
      } catch (_) {
        // File doesn't exist, nothing to delete
        return;
      }
      
      // Remove the campaign
      if (allCampaigns[campaignId]) {
        delete allCampaigns[campaignId];
        
        // Save back to single file
        await callGraph(`/drives/${driveId}/root:/${encodeURIComponent(CMP_ROOT)}/all-campaigns.json:/content?@microsoft.graph.conflictBehavior=replace`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: allCampaigns
        });
      }
    } catch (e) {
      console.warn('[AB CMP] remote delete failed', e && e.message || e);
    }
  }


  // Utility: storage
  async function loadCampaigns() {
    const d = await chrome.storage.local.get([STORAGE_KEY]);
    const map = d[STORAGE_KEY] || {};
    return map;
  }
  async function saveCampaigns(map) {
    await chrome.storage.local.set({ [STORAGE_KEY]: map });
  }

  function normalizeEmail(value) {
    return String(value || '').trim().toLowerCase();
  }

  function dedupeEmails(list) {
    const out = [];
    const seen = new Set();
    (list || []).forEach(item => {
      const email = normalizeEmail(typeof item === 'string' ? item : item?.email);
      if (!email || seen.has(email)) return;
      seen.add(email);
      out.push(email);
    });
    return out;
  }

  function ensureRecipientLog(cmp) {
    if (!cmp) return {};
    const source = cmp.recipientLog && typeof cmp.recipientLog === 'object' ? cmp.recipientLog : {};
    const normalized = {};
    Object.keys(source).forEach(key => {
      if (!Object.prototype.hasOwnProperty.call(source, key)) return;
      const entry = normalizeRecipientLogEntry(key, source[key]);
      if (!entry) return;
      const logKey = normalizeEmail(entry.email || key);
      if (!logKey) return;
      normalized[logKey] = entry;
    });
    cmp.recipientLog = normalized;
    return normalized;
  }

  function normalizeRecipientLogEntry(key, entry) {
    if (!entry && !key) return null;
    const normalizedEmail = normalizeEmail(entry?.email || key);
    if (!normalizedEmail) return null;
    const deliveries = Array.isArray(entry?.deliveries)
      ? entry.deliveries.map(item => {
        const at = Number(item?.at);
        if (!Number.isFinite(at)) return null;
        return {
          at,
          type: typeof item?.type === 'string' ? item.type : 'invite',
          manual: !!item?.manual
        };
      }).filter(Boolean).sort((a, b) => a.at - b.at)
      : [];
    const lastDelivery = deliveries.length ? deliveries[deliveries.length - 1] : null;
    const status = typeof entry?.status === 'string' && RECIPIENT_STATUS[entry.status]
      ? entry.status
      : (entry?.respondedAt ? RECIPIENT_STATUS.responded : (deliveries.length ? RECIPIENT_STATUS.sent : RECIPIENT_STATUS.pending));
    const respondedAtMs = entry?.respondedAt != null
      ? Number(entry.respondedAt)
      : (entry?.lastResponseAt != null ? Number(entry.lastResponseAt) : null);
    return {
      email: entry?.email || key,
      status,
      addedAt: Number(entry?.addedAt) || Number(entry?.createdAt) || Date.now(),
      deliveries,
      lastDeliveryAt: Number(entry?.lastDeliveryAt) || (lastDelivery ? lastDelivery.at : null),
      lastDeliveryType: entry?.lastDeliveryType || (lastDelivery ? lastDelivery.type : null),
      respondedAt: Number.isFinite(respondedAtMs) ? respondedAtMs : null,
      lastResponseSubject: entry?.lastResponseSubject || entry?.responseSubject || null,
      lastResponsePreview: entry?.lastResponsePreview || entry?.responsePreview || null,
      removedAt: entry?.removedAt != null ? Number(entry.removedAt) : null
    };
  }

  function getRecipientLogEntry(cmp, email) {
    if (!cmp) return null;
    const normalizedEmail = normalizeEmail(email);
    if (!normalizedEmail) return null;
    const log = ensureRecipientLog(cmp);
    if (!log[normalizedEmail]) {
      log[normalizedEmail] = {
        email: email || normalizedEmail,
        status: RECIPIENT_STATUS.pending,
        addedAt: Date.now(),
        deliveries: [],
        lastDeliveryAt: null,
        lastDeliveryType: null,
        respondedAt: null,
        lastResponseSubject: null,
        lastResponsePreview: null,
        removedAt: null
      };
    } else if (!log[normalizedEmail].email && email) {
      log[normalizedEmail].email = email;
    }
    return log[normalizedEmail];
  }

  function recordRecipientAddition(cmp, email) {
    const entry = getRecipientLogEntry(cmp, email);
    if (!entry) return;
    if (entry.status !== RECIPIENT_STATUS.responded) {
      entry.status = RECIPIENT_STATUS.pending;
    }
    if (!entry.addedAt) {
      entry.addedAt = Date.now();
    }
    entry.removedAt = null;
  }

  function recordRecipientDelivery(cmp, email, details = {}) {
    const entry = getRecipientLogEntry(cmp, email);
    if (!entry) return;
    const when = Number(details.at) || Date.now();
    const type = typeof details.type === 'string' ? details.type : 'invite';
    entry.deliveries = Array.isArray(entry.deliveries) ? entry.deliveries : [];
    entry.deliveries.push({ at: when, type, manual: !!details.manual });
    entry.lastDeliveryAt = when;
    entry.lastDeliveryType = type;
    if (entry.status !== RECIPIENT_STATUS.responded) {
      entry.status = RECIPIENT_STATUS.sent;
    }
  }

  function recordRecipientResponse(cmp, email, details = {}) {
    const entry = getRecipientLogEntry(cmp, email);
    if (!entry) return;
    entry.status = RECIPIENT_STATUS.responded;
    const response = details?.response || {};
    let atMs = null;
    if (response.at) {
      const parsed = Date.parse(response.at);
      if (Number.isFinite(parsed)) atMs = parsed;
    }
    if (!atMs && details.at) {
      const numeric = Number(details.at);
      if (Number.isFinite(numeric)) atMs = numeric;
    }
    entry.respondedAt = atMs || Date.now();
    if (response.subject) entry.lastResponseSubject = response.subject;
    if (response.preview) entry.lastResponsePreview = response.preview;
    if (details.subject && !entry.lastResponseSubject) entry.lastResponseSubject = details.subject;
    if (details.preview && !entry.lastResponsePreview) entry.lastResponsePreview = details.preview;
    entry.removedAt = null;
  }

  function recordRecipientRemoval(cmp, email, details = {}) {
    const entry = getRecipientLogEntry(cmp, email);
    if (!entry) return;
    if (details.reason === 'response') return;
    if (entry.status === RECIPIENT_STATUS.responded) return;
    entry.status = RECIPIENT_STATUS.removed;
    entry.removedAt = Date.now();
  }

  function ensureCampaignShape(cmp) {
    if (!cmp) return cmp;
    const responses = cmp.responses && typeof cmp.responses === 'object' ? cmp.responses : {};
    const responded = new Set(Object.keys(responses).map(normalizeEmail));
    const seen = new Set();
    const recipients = Array.isArray(cmp.recipients) ? cmp.recipients : [];
    cmp.recipients = recipients.reduce((acc, rec) => {
      const email = normalizeEmail(typeof rec === 'string' ? rec : rec?.email);
      if (!email || seen.has(email) || responded.has(email)) return acc;
      seen.add(email);
      acc.push(Object.assign({}, rec, { email, addedAt: rec?.addedAt || rec?.createdAt || cmp.createdAt || Date.now() }));
      return acc;
    }, []);
    cmp.responses = responses;
    ensureRecipientLog(cmp);
    Object.keys(responses).forEach(addr => {
      recordRecipientResponse(cmp, addr, { response: responses[addr] });
    });
    cmp.schedule = normalizeSchedule(cmp.schedule || {});
    cmp.followUps = normalizeFollowUps(cmp.followUps || {}, Date.now());
    if (!cmp.nextSendType) cmp.nextSendType = cmp.paused ? null : "invite";
    const nextSendNumeric = cmp.nextSendAt != null ? Number(cmp.nextSendAt) : null;
    cmp.nextSendAt = Number.isFinite(nextSendNumeric) ? nextSendNumeric : null;
    cmp.defaults = cmp.defaults && typeof cmp.defaults === 'object' ? cmp.defaults : {};
    if (typeof cmp.paused !== 'boolean') cmp.paused = true;
    if (!cmp.id) cmp.id = newId();
      cmp.sentHistory = Array.isArray(cmp.sentHistory) ? cmp.sentHistory : (cmp.lastSentAt ? [cmp.lastSentAt] : []);
    return cmp;
  }

  function resolveCampaignId(message, payload, map) {
    const candidates = [
      message?.id,
      message?.campaignId,
      payload?.id,
      payload?.campaignId,
      payload?.item?.id,
      message?.item?.id
    ];
    for (const candidate of candidates) {
      const id = typeof candidate === 'string' ? candidate.trim() : '';
      if (!id) continue;
      if (map && map[id]) return id;
      if (id) return id;
    }
    const name = (payload?.name || message?.name || '').trim().toLowerCase();
    if (name && map) {
      const entry = Object.values(map).find(c => String(c?.name || '').trim().toLowerCase() === name);
      if (entry && entry.id) return entry.id;
    }
    return null;
  }

  function normalizeSchedule(schedule = {}, referenceMs) {
  const now = Number(referenceMs || Date.now());
  const minLead = SCHEDULE_MIN_LEAD_MS;
  const raw = schedule && typeof schedule === "object" ? Object.assign({}, schedule) : {};
  const customDates = [];
  const seen = new Set();
  const source = Array.isArray(raw.customDates) ? raw.customDates : [];
  for (const value of source) {
    let date;
    if (value instanceof Date) {
      date = value;
    } else if (typeof value === "number") {
      date = new Date(value);
    } else {
      date = new Date(value);
    }
    if (!date || isNaN(date.getTime())) continue;
    const ms = date.getTime();
    if (!isFinite(ms) || ms < now + minLead) continue;
    const iso = date.toISOString();
    if (seen.has(iso)) continue;
    seen.add(iso);
    customDates.push(iso);
  }
  customDates.sort();
  let mode = raw.mode === "custom" ? "custom" : "recurring";
  if (!customDates.length) mode = "recurring";
  let everyDays = Number(raw.everyDays);
  if (!isFinite(everyDays) || everyDays < 0) everyDays = 0;
  let untilIso = null;
  if (raw.until) {
    const untilDate = new Date(raw.until);
    if (untilDate && !isNaN(untilDate.getTime())) {
      untilIso = untilDate.toISOString();
    }
  }
  if (mode === "custom") {
    everyDays = 0;
    untilIso = null;
  }
  return {
    mode,
    everyDays,
    until: untilIso,
    customDates
  };
}

function normalizeFollowUps(followUps = {}, referenceMs) {
  const now = Number(referenceMs || Date.now());
  const minLead = SCHEDULE_MIN_LEAD_MS;
  const source = followUps && typeof followUps === "object" ? followUps : {};
  const out = {};
  FOLLOW_UP_KEYS.forEach(key => {
    const item = source[key] || {};
    const enabled = !!item.enabled;
    const subjectTpl = item.subjectTpl != null ? String(item.subjectTpl) : '';
    const htmlTpl = item.htmlTpl != null ? String(item.htmlTpl) : '';
    const rawDates = Array.isArray(item.dates) ? item.dates : (Array.isArray(item.customDates) ? item.customDates : []);
    const dates = [];
    const seen = new Set();
    rawDates.forEach(value => {
      let date;
      if (value instanceof Date) {
        date = value;
      } else if (typeof value === 'number') {
        date = new Date(value);
      } else {
        date = new Date(value);
      }
      if (!date || isNaN(date.getTime())) return;
      const ms = date.getTime();
      if (!isFinite(ms) || ms < now + minLead) return;
      const iso = date.toISOString();
      if (seen.has(iso)) return;
      seen.add(iso);
      dates.push(iso);
    });
    dates.sort();
    out[key] = {
      enabled: enabled,
      subjectTpl,
      htmlTpl,
      dates,
      lastSentAt: item.lastSentAt != null ? Number(item.lastSentAt) || null : null
    };
  });
  return out;
}

function computeNextEvent(cmp, referenceMs) {
    const now = Number(referenceMs || Date.now());
    if (!cmp || cmp.paused) return null;
    let nextTime = null;
    let nextType = null;
    const mainNext = computeNextSendAtForSchedule(cmp.schedule, now);
    if (mainNext && (!nextTime || mainNext < nextTime)) {
      nextTime = mainNext;
      nextType = 'invite';
    }
    FOLLOW_UP_KEYS.forEach(key => {
      const follow = cmp.followUps && cmp.followUps[key];
      if (!follow || !follow.enabled) return;
      const dates = Array.isArray(follow.dates) ? follow.dates : [];
      for (let i = 0; i < dates.length; i++) {
        const ms = Date.parse(dates[i]);
        if (!isFinite(ms)) continue;
        if (ms < now + SCHEDULE_MIN_LEAD_MS) continue;
        if (!nextTime || ms < nextTime) {
          nextTime = ms;
          nextType = key;
        }
        break;
      }
    });
    return nextTime ? { time: nextTime, type: nextType || 'invite' } : null;
  }

function computeNextSendAtForSchedule(schedule, referenceMs) {
    const now = Number(referenceMs || Date.now());
    if (!schedule || typeof schedule !== "object") return null;
    if (schedule.mode === "custom") {
      const dates = Array.isArray(schedule.customDates) ? schedule.customDates : [];
      for (let i = 0; i < dates.length; i++) {
        const ms = Date.parse(dates[i]);
        if (isFinite(ms) && ms >= now + SCHEDULE_MIN_LEAD_MS) return ms;
      }
      return null;
    }
    const every = Math.max(1, Number(schedule.everyDays || 0));
    if (!every || !isFinite(every)) return null;
    const untilMs = schedule.until ? Date.parse(schedule.until) : null;
    return nextSendTime(now, every, untilMs);
  }

  function applySchedule(cmp, schedule = {}, options = {}) {
    if (!cmp) return;
    const reference = Number(options.reference || Date.now());
    const merged = Object.assign({}, cmp.schedule || {}, schedule || {});
    cmp.schedule = normalizeSchedule(merged, reference);
    const followUpsInput = options.followUps != null ? options.followUps : cmp.followUps || {};
    cmp.followUps = normalizeFollowUps(followUpsInput, reference);
    if (cmp.paused) {
      cmp.nextSendAt = null;
      cmp.nextSendType = null;
    } else {
      const nextEvent = computeNextEvent(cmp, reference);
      cmp.nextSendAt = nextEvent ? nextEvent.time : null;
      cmp.nextSendType = nextEvent ? nextEvent.type : null;
    }
  }

  function updateAlarmsForCampaign(cmp) {
    if (!chrome?.alarms || !cmp?.id) return;
    chrome.alarms.clear(cmp.id);
    if (!cmp.paused && cmp.nextSendAt) {
      scheduleAlarmForCampaign(cmp);
    }
  }

  function addRecipientEmails(cmp, emails) {
    if (!cmp) return 0;
    const additions = dedupeEmails(emails);
    if (!additions.length) return 0;
    ensureRecipientLog(cmp);
    const existing = new Set((cmp.recipients || []).map(r => normalizeEmail(r.email)));
    const responded = new Set(Object.keys(cmp.responses || {}).map(normalizeEmail));
    responded.forEach(email => existing.add(email));
    const now = Date.now();
    let added = 0;
    additions.forEach(email => {
      if (!email || existing.has(email) || responded.has(email)) return;
      existing.add(email);
      cmp.recipients.push({ email, addedAt: now });
      recordRecipientAddition(cmp, email);
      added++;
    });
    return added;
  }

  function removeRecipientEmails(cmp, emails, options = {}) {
    if (!cmp || !Array.isArray(cmp.recipients)) return 0;
    const removal = new Set(dedupeEmails(emails));
    if (!removal.size) return 0;
    const before = cmp.recipients.length;
    const removedKeys = [];
    cmp.recipients = cmp.recipients.filter(rec => {
      const key = normalizeEmail(rec.email);
      if (removal.has(key)) {
        removedKeys.push(key);
        return false;
      }
      return true;
    });
    if (!options.skipLogUpdate && removedKeys.length) {
      removedKeys.forEach(email => recordRecipientRemoval(cmp, email, { reason: options.reason || 'manual' }));
    }
    return before - cmp.recipients.length;
  }

  async function persistCampaign(map, cmp, options = {}) {
    cmp = ensureCampaignShape(cmp);
    cmp.updatedAt = Date.now();
    map[cmp.id] = cmp;
    await saveCampaigns(map);
    if (!options.skipRemote) {
      try { await saveCampaignRemote(cmp); } catch (_) {}
    }
    if (!options.skipAlarms) {
      updateAlarmsForCampaign(cmp);
    }
      cmp.sentHistory = Array.isArray(cmp.sentHistory) ? cmp.sentHistory : (cmp.lastSentAt ? [cmp.lastSentAt] : []);
    return cmp;
  }

  async function sendEmails(cmp, emails, options = {}) {
    const normalized = dedupeEmails(emails);
    if (!normalized.length) return 0;
    const responded = new Set(Object.keys(cmp.responses || {}).map(normalizeEmail));
    const queue = normalized.filter(email => email && !responded.has(email));
    if (!queue.length) return 0;
    ensureRecipientLog(cmp);
    const tag = tagFor(cmp.id);
    const subject = renderTemplate(cmp.subjectTpl || '', cmp.defaults || {});
    const htmlBase = renderTemplate(cmp.htmlTpl || '', cmp.defaults || {});
    const html = htmlBase + '<div style="display:none;font-size:1px;color:#fff">' + tag + '</div>';
    const sendType = typeof options.type === 'string' ? options.type : 'invite';
    let sent = 0;
    for (const email of queue) {
      try {
        await sendGraphMail({ subject, html, to: [email] });
        recordRecipientDelivery(cmp, email, { at: Date.now(), type: sendType, manual: !!options.manual });
        sent++;
      } catch (err) {
        console.warn('[AB CMP] send mail failed', err && err.message ? err.message : err);
      }
    }
    cmp.updatedAt = Date.now();
    return sent;
  }
  // Utility: id and tag
  function newId() {
    return 'cmp_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8);
  }
  function tagFor(id) {
    return `[CMP-${id}]`;
  }

  // Utility: template replace
  function renderTemplate(tpl = '', vars = {}) {
    return String(tpl).replace(/\{\{\s*([a-z0-9_]+)\s*\}\}/gi, (_, key) => {
      const k = String(key || '').toLowerCase();
      const v = vars[k];
      return v == null ? '' : String(v);
    });
  }

  // Build and send one email via Microsoft Graph
  async function sendGraphMail({ subject, html, to }) {
    if (!Array.isArray(to) || !to.length) return;
    const body = {
      message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: to.map(addr => ({ emailAddress: { address: addr } }))
      },
      saveToSentItems: true
    };
    // callGraphAPI is defined in background.js
    await callGraphAPI('/me/sendMail', { method: 'POST', body });
  }

  // Schedule helpers
  function nextSendTime(fromMs, everyDays, untilMs) {
    const base = Math.max(Date.now(), fromMs || 0);
    const n = Math.max(1, Number(everyDays || 1));
    const next = base + n * 24 * 60 * 60 * 1000;
    if (untilMs && next > untilMs) return null;
    return next;
  }

  function scheduleAlarmForCampaign(cmp) {
    if (!cmp || !cmp.id) return;
    if (cmp.paused) return;
    if (!cmp.nextSendAt) return;
    chrome.alarms.create(cmp.id, { when: cmp.nextSendAt });
  }

  async function rescheduleIfNeeded(cmp) {
    if (!cmp) return;
    const now = Date.now();
    applySchedule(cmp, cmp.schedule, { reference: now, followUps: cmp.followUps });
    await upsertCampaign(cmp);
    updateAlarmsForCampaign(cmp);
  }

  // Maintain campaigns map with immutability
  async function upsertCampaign(cmp) {
    const map = await loadCampaigns();
    cmp = ensureCampaignShape(cmp);
    cmp.updatedAt = Date.now();
    map[cmp.id] = cmp;
    await saveCampaigns(map);
    try { await saveCampaignRemote(cmp); } catch (_) {}
    return map;
  }

  // Determine recipients to send to
  function pendingRecipients(cmp) {
    const responded = new Set(Object.keys(cmp.responses || {}));
    return (cmp.recipients || [])
      .filter(r => r.email)
      .filter(r => !responded.has(String(r.email).toLowerCase()));
  }

  function recordSendEvent(cmp, eventType, sentAt) {
    if (!cmp) return;
    const stamp = Number(sentAt) || Date.now();
    if (eventType === 'invite') {
      cmp.lastSentAt = stamp;
      cmp.sentHistory = Array.isArray(cmp.sentHistory) ? cmp.sentHistory : [];
      cmp.sentHistory.push(stamp);
      if (cmp.schedule && cmp.schedule.mode === 'custom' && Array.isArray(cmp.schedule.customDates) && cmp.schedule.customDates.length) {
        cmp.schedule.customDates.shift();
      }
    } else {
      cmp.followUps = cmp.followUps && typeof cmp.followUps === 'object' ? cmp.followUps : {};
      if (!cmp.followUps[eventType]) {
        cmp.followUps[eventType] = { enabled: false, dates: [] };
      }
      cmp.followUps[eventType].lastSentAt = stamp;
      if (Array.isArray(cmp.followUps[eventType].dates) && cmp.followUps[eventType].dates.length) {
        cmp.followUps[eventType].dates.shift();
      }
    }
  }

  // One send cycle: send to non-responders
  async function runSendCycle(map, cmp) {
    if (!cmp) return null;
    console.log('[runSendCycle] Starting for campaign:', cmp.name);
    const now = Date.now();
    cmp.schedule = normalizeSchedule(cmp.schedule || {}, now);
    cmp.followUps = normalizeFollowUps(cmp.followUps || {}, now);

    const nextEvent = computeNextEvent(cmp, now);

    console.log('[runSendCycle] Next event:', nextEvent);

    if (!nextEvent && !cmp.nextSendAt) {

      console.log('[runSendCycle] No next event - exiting');

      cmp.nextSendAt = null;

      cmp.nextSendType = null;

      await persistCampaign(map, cmp);

      return cmp;

    }



    const storedNextAt = Number(cmp.nextSendAt);

    const storedType = cmp.nextSendType || null;

    const hasStoredSchedule = Number.isFinite(storedNextAt);

    const timeDiff = hasStoredSchedule ? Math.abs(storedNextAt - now) : Infinity;

    const isWithinSendWindow = hasStoredSchedule && timeDiff < SCHEDULE_MIN_LEAD_MS * 2; // Within 2 minutes



    console.log('[runSendCycle] Time check - now:', now, 'nextSendAt:', cmp.nextSendAt, 'diff:', timeDiff, 'ms');



    let targetEventTime = nextEvent ? nextEvent.time : storedNextAt;

    let targetEventType = nextEvent ? (nextEvent.type || 'invite') : (storedType || 'invite');



    if (isWithinSendWindow) {

      if (hasStoredSchedule) targetEventTime = storedNextAt;

      if (storedType) targetEventType = storedType;

      console.log('[runSendCycle] Within send window - proceeding to send! Using stored event:', { time: targetEventTime, type: targetEventType });

    } else {

      console.log('[runSendCycle] Not within send window - updating schedule');

      if (!nextEvent) {

        console.log('[runSendCycle] No computed event available - exiting');

        cmp.nextSendAt = null;

        cmp.nextSendType = null;

        await persistCampaign(map, cmp);

        return cmp;

      }

      if (!cmp.nextSendAt || Math.abs(Number(cmp.nextSendAt) - nextEvent.time) > SCHEDULE_MIN_LEAD_MS || cmp.nextSendType !== nextEvent.type) {

        console.log('[runSendCycle] Updating next send time - not sending yet');

        cmp.nextSendAt = nextEvent.time;

        cmp.nextSendType = nextEvent.type;

        await persistCampaign(map, cmp);

        return cmp;

      }

      targetEventTime = nextEvent.time;

      targetEventType = nextEvent.type || 'invite';

    }



    if (!isWithinSendWindow && targetEventTime - now > SCHEDULE_MIN_LEAD_MS) {

      console.log('[runSendCycle] Next event too far in future - exiting');

      return cmp;

    }



    const batch = pendingRecipients(cmp);
    console.log('[runSendCycle] Pending recipients:', batch.length);
    if (!batch.length) {
      console.log('[runSendCycle] No pending recipients - removing date and exiting');
      if (targetEventType === 'invite' && cmp.schedule.mode === 'custom' && Array.isArray(cmp.schedule.customDates)) {
        cmp.schedule.customDates.shift();
      } else if (targetEventType !== 'invite' && cmp.followUps[targetEventType] && Array.isArray(cmp.followUps[targetEventType].dates)) {
        cmp.followUps[targetEventType].dates.shift();
      }
      applySchedule(cmp, cmp.schedule, { reference: now, followUps: cmp.followUps });
      await persistCampaign(map, cmp);
      return cmp;
    }

    console.log('[runSendCycle] Preparing to send emails...');
    const tag = tagFor(cmp.id);
    let subjectTpl = cmp.subjectTpl || '';
    let htmlTpl = cmp.htmlTpl || '';
    if (targetEventType !== 'invite') {
      const follow = cmp.followUps[targetEventType] || {};
      if (follow.subjectTpl) subjectTpl = follow.subjectTpl;
      if (follow.htmlTpl) htmlTpl = follow.htmlTpl;
    }
    const subj = renderTemplate(subjectTpl, cmp.defaults || {});
    const htmlBase = renderTemplate(htmlTpl, cmp.defaults || {});
    const html = `${htmlBase}<div style="display:none;font-size:1px;color:#fff">${tag}</div>`;
    console.log('[runSendCycle] Email subject:', subj);
    console.log('[runSendCycle] Email type:', targetEventType);
    
    let sentCount = 0;
    for (const recipient of batch) {
      const address = String(recipient && recipient.email || '').trim();
      if (!address) continue;
      console.log('[runSendCycle] Sending to:', address);
      try {
        await sendGraphMail({ subject: subj, html, to: [address] });
        recordRecipientDelivery(cmp, address, { at: Date.now(), type: targetEventType || 'invite', manual: false });
        sentCount++;
        console.log('[runSendCycle] ✅ Sent successfully to:', address);
      } catch (err) {
        console.error('[runSendCycle] ❌ Send failed to', address, ':', err && err.message ? err.message : err);
      }
    }

    console.log('[runSendCycle] Total sent:', sentCount);
    const sentAt = Date.now();
    recordSendEvent(cmp, targetEventType || 'invite', sentAt);

    applySchedule(cmp, cmp.schedule, { reference: Date.now(), followUps: cmp.followUps });
    await persistCampaign(map, cmp);
    return cmp;
  }

  // Poll responses by $search on the tag
  async function pollResponses() {
    const map = await loadCampaigns();
    const ids = Object.keys(map);
    if (!ids.length) return;

    let anyChanged = false;
    for (const id of ids) {
      const cmp = map[id];
      if (!cmp) continue;
      const tag = tagFor(id);
      // ConsistencyLevel header required for $search
      const res = await callGraphAPI(`/me/messages?$search="${encodeURIComponent(tag)}"&$top=25`, {
        headers: { 'ConsistencyLevel': 'eventual' }
      });

      let cmpChanged = false;
      const items = Array.isArray(res?.value) ? res.value : [];
      for (const msg of items) {
        const from = msg?.from?.emailAddress?.address || '';
        if (!from) continue;
        const addr = String(from).toLowerCase();
        // Ignore self
        if (addr.includes('@') === false) continue;

        // Mark responded
        cmp.responses = cmp.responses || {};
        if (!cmp.responses[addr]) {
          cmp.responses[addr] = {
            at: msg.receivedDateTime || new Date().toISOString(),
            subject: msg.subject || '',
            preview: msg.bodyPreview || ''
          };
          cmpChanged = true;
          recordRecipientResponse(cmp, addr, { response: cmp.responses[addr] });
        }
        removeRecipientEmails(cmp, [addr], { skipLogUpdate: true, reason: 'response' });
      }

      if (cmpChanged) {
        const shaped = ensureCampaignShape(cmp);
        map[id] = shaped;
        anyChanged = true;
        try { await saveCampaignRemote(shaped); } catch (_) {}
      }
    }

    if (anyChanged) {
      await saveCampaigns(map);
      try { chrome.runtime.sendMessage({ type: 'AB_CMP_REFRESH' }); } catch (_) {}
    }
  }  // Alarms
  chrome.runtime.onInstalled.addListener(() => {
    chrome.alarms.create(POLL_ALARM, { periodInMinutes: 5 });
  });
  chrome.runtime.onStartup.addListener(() => {
    chrome.alarms.create(POLL_ALARM, { periodInMinutes: 5 });
  });

  chrome.alarms.onAlarm.addListener(async (alarm) => {
    if (alarm.name === POLL_ALARM) {
      try { await pollResponses(); } catch (e) { /* no-op */ }
      return;
    }
    // Campaign alarm
    const map = await loadCampaigns();
    const cmp = map[alarm.name];
    if (!cmp) return;

    try {
      console.log('[CMP ALARM] Starting runSendCycle for:', cmp.name);
      await runSendCycle(map, cmp);
      console.log('[CMP ALARM] runSendCycle completed successfully');
    } catch (e) {
      console.error('❌ [CMP ALARM] runSendCycle FAILED:', e);
      console.error('   Error message:', e.message);
      console.error('   Error stack:', e.stack);
    }
    // schedule next
    await rescheduleIfNeeded(cmp);
  });

  // Message handlers for the UI
  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    var type = typeof message?.type === 'string' ? message.type : '';
    var isCampaignMessage = type.startsWith('AB_CAMPAIGN') || type.startsWith('AB_CMP') || type.startsWith('campaign:');
    if (!isCampaignMessage) return false;
    try { console.log('[AB CMP] message', type); } catch (_) {}
    var finished = false;
    function safeRespond(payload) {
      if (finished) return;
      finished = true;
      try { sendResponse(payload); } catch (err) { console.error('[AB CMP] sendResponse failed', err); }
    }
    (async () => {
      switch (message?.type) {
        case 'AB_CAMPAIGN_LIST':
        case 'AB_CMP_LIST': {
          const map = await loadCampaigns();
          let changed = false;
          try {
            const remote = await listRemoteCampaigns();
            remote.forEach(r => {
              if (!r || !r.id) return;
              const merged = ensureCampaignShape(Object.assign({}, map[r.id] || {}, r));
              map[r.id] = merged;
              changed = true;
            });
          } catch (_) {}
          if (changed) await saveCampaigns(map);
          const items = Object.values(map).map(ensureCampaignShape);
          items.sort((a, b) => (b.updatedAt || b.lastSentAt || b.createdAt || 0) - (a.updatedAt || a.lastSentAt || a.createdAt || 0));
          safeRespond({ success: true, campaigns: map, items });
          return;
        }
        case 'AB_CAMPAIGN_CREATE':
        case 'AB_CMP_CREATE': {
          const map = await loadCampaigns();
          const explicitId = typeof message.id === 'string' ? message.id.trim() : '';
          const id = explicitId && !map[explicitId] ? explicitId : newId();
          const now = Date.now();
          const schedule = normalizeSchedule(message.schedule || {}, now);
          const followUps = normalizeFollowUps(message.followUps || {}, now);
          const cmp = ensureCampaignShape({
            id,
            name: message.name || id,
            createdAt: now,
            updatedAt: now,
            paused: message.paused === true ? true : false,
            recipients: Array.isArray(message.recipients) ? message.recipients : [],
            subjectTpl: message.subjectTpl || '',
            htmlTpl: message.htmlTpl || '',
            defaults: Object.assign({}, message.defaults || {}),
            schedule: schedule,
            followUps: followUps,
            responses: {},
            lastSentAt: null,
            nextSendAt: null,
            source: message.source || 'manual'
          });
          applySchedule(cmp, schedule, { reference: now, followUps: followUps });
          if (cmp.paused) cmp.nextSendAt = null;
          await persistCampaign(map, cmp);
          const response = { success: true, item: cmp, campaign: cmp };
          safeRespond(response);
          return;
        }
        case 'AB_CAMPAIGN_PREVIEW':
        case 'campaign:preview': {
          const payload = message?.payload || {};
          const emails = dedupeEmails([]
            .concat(message?.manualEmails || [])
            .concat(payload.manualEmails || [])
            .concat(payload.manual || []));
          const preview = emails.map(email => ({ email }));
          safeRespond({ success: true, count: preview.length, preview });
          return;
        }
        case 'AB_CAMPAIGN_SAVE':
        case 'campaign:save':
        case 'AB_CMP_ADD_RECIPIENTS': {
          const payload = message?.payload || {};
          const map = await loadCampaigns();
          let id = resolveCampaignId(message, payload, map);
          let cmp = id ? map[id] : null;
          if (!cmp) {
            id = id || newId();
            const now = Date.now();
            cmp = ensureCampaignShape({
              id,
              name: payload.name || message.name || id,
              createdAt: now,
              updatedAt: now,
              paused: true,
              recipients: [],
              responses: {},
              subjectTpl: payload.subjectTpl || '',
              htmlTpl: payload.htmlTpl || '',
              defaults: Object.assign({}, payload.defaults || {}),
              schedule: {
                everyDays: Number(payload.schedule?.everyDays || 0) || 0,
                until: payload.schedule?.until || null
              },
              source: payload.source || message.source || 'manual',
              lastSentAt: null,
              nextSendAt: null
            });
          } else {
            cmp = ensureCampaignShape(Object.assign({}, cmp));
            if (payload.subjectTpl != null) cmp.subjectTpl = payload.subjectTpl;
            if (payload.htmlTpl != null) cmp.htmlTpl = payload.htmlTpl;
            if (payload.source != null) cmp.source = payload.source;
            if (payload.defaults) cmp.defaults = Object.assign({}, cmp.defaults, payload.defaults);
            if (payload.schedule) applySchedule(cmp, payload.schedule, { reference: Date.now() });
          }
          const additionsSource = message.type === 'AB_CMP_ADD_RECIPIENTS'
            ? message.emails
            : (payload.manualEmails || payload.manual || []);
          const addedCount = addRecipientEmails(cmp, additionsSource);
          await persistCampaign(map, cmp);
          safeRespond({ success: true, id: cmp.id, item: cmp, campaign: cmp, added: addedCount });
          return;
        }
        case 'AB_CMP_REMOVE_RECIPIENTS': {
          const map = await loadCampaigns();
          const id = resolveCampaignId(message, { id: message.id }, map);
          const cmp = id ? map[id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          const clone = ensureCampaignShape(Object.assign({}, cmp));
          const removed = removeRecipientEmails(clone, message.emails);
          await persistCampaign(map, clone);
          safeRespond({ success: true, item: clone, removed });
          return;
        }
        case 'AB_CMP_GET': {
          const map = await loadCampaigns();
          const id = resolveCampaignId(message, { id: message.id }, map);
          const cmp = id ? map[id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          const shaped = ensureCampaignShape(Object.assign({}, cmp));
          safeRespond({ success: true, item: shaped });
          return;
        }
        case 'AB_CMP_UPDATE': {
          const payload = {
            id: message.id,
            schedule: message.schedule,
            subjectTpl: message.subjectTpl,
            htmlTpl: message.htmlTpl,
            defaults: message.defaults,
            followUps: message.followUps
          };
          const map = await loadCampaigns();
          const id = resolveCampaignId(message, payload, map);
          const cmp = id ? map[id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          const clone = ensureCampaignShape(Object.assign({}, cmp));
          if (payload.subjectTpl != null) clone.subjectTpl = payload.subjectTpl;
          if (payload.htmlTpl != null) clone.htmlTpl = payload.htmlTpl;
          if (payload.defaults) clone.defaults = Object.assign({}, clone.defaults, payload.defaults);
          if (payload.followUps) clone.followUps = normalizeFollowUps(payload.followUps, Date.now());
          if (payload.schedule) {
            applySchedule(clone, payload.schedule, { reference: Date.now(), followUps: clone.followUps });
          } else {
            applySchedule(clone, clone.schedule, { reference: Date.now(), followUps: clone.followUps });
          }
          await persistCampaign(map, clone);
          safeRespond({ success: true, item: clone });
          return;
        }
        case 'AB_CMP_TOGGLE_PAUSE': {
          const map = await loadCampaigns();
          const id = resolveCampaignId(message, { id: message.id }, map);
          const cmp = id ? map[id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          const clone = ensureCampaignShape(Object.assign({}, cmp));
          clone.paused = !clone.paused;
          if (clone.paused) {
            clone.nextSendAt = null;
          } else {
            applySchedule(clone, clone.schedule, { reference: Date.now(), followUps: clone.followUps });
          }
          await persistCampaign(map, clone);
          safeRespond({ success: true, item: clone });
          return;
        }
        case 'AB_CMP_SEND_SELECTED': {
          const map = await loadCampaigns();
          const id = resolveCampaignId(message, { id: message.id }, map);
          const cmp = id ? map[id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          const clone = ensureCampaignShape(Object.assign({}, cmp));
          const now = Date.now();
          if (clone.nextSendAt && clone.nextSendAt - now > SCHEDULE_MIN_LEAD_MS) {
            safeRespond({ success: false, error: 'SEND_SCHEDULED_LATER', nextSendAt: clone.nextSendAt });
            return;
          }
          addRecipientEmails(clone, message.emails);
          const sendType = clone.nextSendType || 'invite';
          const sentCount = await sendEmails(clone, message.emails, { type: sendType, manual: true });
          recordSendEvent(clone, sendType, Date.now());
          applySchedule(clone, clone.schedule, { reference: Date.now(), followUps: clone.followUps });
          await persistCampaign(map, clone);
          safeRespond({ success: true, item: clone, sentCount });
          return;
        }
        case 'AB_CAMPAIGN_SEND':
        case 'campaign:send': {
          const payload = message?.payload || {};
          const map = await loadCampaigns();
          let id = resolveCampaignId(message, payload, map);
          let cmp = id ? map[id] : null;
          if (!cmp) {
            id = id || newId();
            const now = Date.now();
            cmp = ensureCampaignShape({
              id,
              name: payload.name || message.name || id,
              createdAt: now,
              updatedAt: now,
              paused: true,
              recipients: [],
              responses: {},
              subjectTpl: payload.subject || payload.subjectTpl || '',
              htmlTpl: payload.html || payload.htmlTpl || '',
              defaults: {},
              schedule: { everyDays: 0, until: null },
              source: payload.source || 'manual',
              lastSentAt: null,
              nextSendAt: null
            });
          } else {
            cmp = ensureCampaignShape(Object.assign({}, cmp));
            if (payload.subjectTpl != null || payload.subject != null) cmp.subjectTpl = payload.subjectTpl || payload.subject || cmp.subjectTpl;
            if (payload.htmlTpl != null || payload.html != null) cmp.htmlTpl = payload.htmlTpl || payload.html || cmp.htmlTpl;
          }
          const emails = []
            .concat(message?.manualEmails || [])
            .concat(payload.manualEmails || [])
            .concat(payload.manual || []);
          addRecipientEmails(cmp, emails);
          const targetList = emails.length ? emails : pendingRecipients(cmp).map(r => r.email);
          const sendType = cmp.nextSendType || 'invite';
          const sentCount = await sendEmails(cmp, targetList, { type: sendType, manual: true });
          recordSendEvent(cmp, sendType, Date.now());
          applySchedule(cmp, cmp.schedule, { reference: Date.now(), followUps: cmp.followUps });
          await persistCampaign(map, cmp);
          safeRespond({ success: true, id: cmp.id, item: cmp, sentCount });
          return;
        }
        case 'AB_CAMPAIGN_LIST_LISTS':
        case 'campaign:listLists': {
          safeRespond({ success: true, lists: [] });
          return;
        }
        case 'AB_CAMPAIGN_STATUS':
        case 'campaign:status': {
          const map = await loadCampaigns();
          const id = resolveCampaignId(message, { id: message?.id || message?.payload?.id }, map);
          const cmp = id ? map[id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          const shaped = ensureCampaignShape(Object.assign({}, cmp));
          const counts = { queued: (shaped.recipients || []).length };
          safeRespond({ success: true, id: shaped.id, counts, item: shaped, campaign: shaped });
          return;
        }
        case 'AB_CAMPAIGN_FORCE_SEND_NOW': {
          const map = await loadCampaigns();
          const cmp = message.id ? map[message.id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          await runSendCycle(cmp);
          await rescheduleIfNeeded(cmp);
          safeRespond({ success: true });
          return;
        }
        case 'AB_CAMPAIGN_PAUSE': {
          const map = await loadCampaigns();
          const cmp = message.id ? map[message.id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          cmp.paused = true;
          cmp.nextSendAt = null;
          await persistCampaign(map, cmp);
          safeRespond({ success: true, item: cmp });
          return;
        }
        case 'AB_CAMPAIGN_RESUME': {
          const map = await loadCampaigns();
          const cmp = message.id ? map[message.id] : null;
          if (!cmp) { safeRespond({ success: false, error: 'Not found' }); return; }
          cmp.paused = false;
          applySchedule(cmp, cmp.schedule, { reference: Date.now(), followUps: cmp.followUps });
          await persistCampaign(map, cmp);
          safeRespond({ success: true, item: cmp });
          return;
        }
        case 'AB_CMP_DELETE':
        case 'AB_CAMPAIGN_DELETE': {
          const map = await loadCampaigns();
          const id = resolveCampaignId(message, { id: message.id }, map) || message.id;
          if (id) {
            delete map[id];
            await saveCampaigns(map);
            chrome.alarms.clear(id);
            await deleteCampaignRemote(id);
          }
          safeRespond({ success: true });
          return;
        }
        case 'AB_CAMPAIGN_POLL_NOW': {
          await pollResponses();
          const map = await loadCampaigns();
          safeRespond({ success: true, campaigns: map, items: Object.values(map).map(ensureCampaignShape) });
          return;
        }
        default:
          safeRespond({ success: false, error: 'Unhandled campaign message' });
          return;
      }
    })().catch(function (err) {
      console.error('[AB CMP] handler failed', err);
      safeRespond({ success: false, error: String(err?.message || err) });
    });
    return true; // async
  });
})();


