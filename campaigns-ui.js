
/* campaigns-ui.js â€” stable ES5 build */

(function () {
  // ---------- Helpers ----------
  function $(id) { return document.getElementById(id); }
  function show(el, flag) { if (!el) return; el.style.display = flag ? "" : "none"; }
  function fmtDate(ms) {
    if (!ms) return "â€”";
    try { return new Date(ms).toLocaleString(); } catch (e) { return "â€”"; }
  }
  function formatDistanceMiles(value) {
    if (value === null || value === undefined) return "";
    var num = Number(value);
    if (!isFinite(num) || num < 0) return "";
    var digits = num < 10 ? 2 : 1;
    return num.toFixed(digits);
  }
  function uniqueEmailsFromInput(value) {
    var out = [], map = {};
    String(value || "").split(/[,\s;]+/g).forEach(function (s) {
      var e = String(s || "").trim().toLowerCase();
      if (!e) return;
      if (map[e]) return;
      map[e] = true;
      out.push(e);
    });
    return out;
  }

  function escapeHtml(str) {
    return String(str || "").replace(/[&<>"']/g, function (ch) {
      switch (ch) {
        case "&": return "&amp;";
        case "<": return "&lt;";
        case ">": return "&gt;";
        case "\"": return "&quot;";
        case "'": return "&#39;";
        default: return ch;
      }
    });
  }

  function normalizeText(value) {
    if (value == null) return '';
    return String(value)
      .trim()
      .toLowerCase()
      .replace(/[\u2010-\u2015\u2212]/g, '-')
      .replace(/\s+/g, ' ');
  }

  function sanitizeZip(value) {
    return String(value || "").replace(/[^0-9]/g, "").slice(0, 5);
  }

  function firstNonEmpty(obj, keys) {
    if (!obj || !keys || !keys.length) return "";
    for (var i = 0; i < keys.length; i++) {
      var val = obj[keys[i]];
      if (val != null) {
        var str = String(val);
        if (str.trim()) return str;
      }
    }
    return "";
  }
  function initRichEditor(wrapper) {
    if (!wrapper) return;
    var targetId = wrapper.getAttribute("data-target");
    if (!targetId) return;
    var textarea = $(targetId);
    var editor = wrapper.querySelector(".rich-content");
    if (!textarea || !editor) return;

    if (textarea.value) {
      editor.innerHTML = textarea.value;
    }

    editor.addEventListener("input", function () {
      textarea.value = editor.innerHTML;
    });
    editor.addEventListener("blur", function () {
      textarea.value = editor.innerHTML;
    });

    var buttons = wrapper.querySelectorAll(".rich-btn");
    Array.prototype.forEach.call(buttons, function (btn) {
      btn.addEventListener("click", function () {
        var cmd = btn.getAttribute("data-command");
        if (!cmd) return;
        editor.focus();
        if (cmd === "createLink") {
          var url = prompt("Enter link URL", "https://");
          if (url) {
            if (window.getSelection && window.getSelection().isCollapsed) {
              var safeUrl = String(url).trim();
              if (!safeUrl) return;
              var html = '<a href="' + safeUrl.replace(/["<>]/g, '') + '" target="_blank" rel="noopener">' + safeUrl + '</a>';
              document.execCommand('insertHTML', false, html);
            } else {
              document.execCommand('createLink', false, url);
            }
          }
        } else if (cmd === "unlink") {
          document.execCommand('unlink', false, null);
        } else {
          document.execCommand(cmd, false, null);
        }
        textarea.value = editor.innerHTML;
      });
    });
  }

  function getEditorValue(editor, textarea) {
    if (editor && typeof editor.innerHTML === "string") return editor.innerHTML;
    if (textarea && typeof textarea.value === "string") return textarea.value;
    return "";
  }

  function setEditorValue(editor, textarea, value) {
    if (textarea) textarea.value = value || "";
    if (editor) editor.innerHTML = value || "";
  }

  function getRuntimeAssetUrl(relativePath) {
    if (!relativePath) return '';
    var cleanPath = String(relativePath).replace(/^\/+/,'');
    if (typeof chrome !== "undefined" && chrome.runtime && typeof chrome.runtime.getURL === "function") {
      return chrome.runtime.getURL(cleanPath);
    }
    return cleanPath;
  }

  function stripDiacritics(value) {
    try {
      return String(value || '').normalize('NFKD').replace(/[\u0300-\u036f]/g, '');
    } catch (error) {
      return String(value || '');
    }
  }

  function sanitizeCityLabel(value) {
    return stripDiacritics(value).replace(/[^A-Za-z0-9'\-\.\s]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function normalizeProjectKey(value) {
    var normalized = normalizeText(value);
    if (!normalized) return '';
    var match = normalized.match(/(\d{2})[-\s]?(\d{2,3})/);
    if (match && match[1]) {
      var code = match[1] + '-' + match[2];
      var pattern = new RegExp(match[1] + '[-\\s]?' + match[2]);
      var remainder = normalized.replace(pattern, '').trim();
      normalized = code + (remainder ? ' ' + remainder : '');
    }
    return normalized;
  }

  function buildUsCityDatasetOptions(payload) {
    var seen = {};
    var options = [];
    (payload || []).forEach(function (entry) {
      if (!entry) return;
      var cityName = sanitizeCityLabel(entry.city || entry.City || '');
      var state = entry.state_code || entry.state || entry.stateCode || '';
      var stateAbbrev = String(state || '').trim().toUpperCase();
      if (!cityName || !stateAbbrev || stateAbbrev === 'HI') return;
      var normalizedCity = normalizeText(cityName);
      if (!normalizedCity) return;
      var key = normalizedCity + '|' + stateAbbrev;
      if (seen[key]) return;
      seen[key] = true;
      var label = cityName + ', ' + stateAbbrev;
      options.push({
        key: key,
        label: label,
        value: normalizedCity,
        stateAbbrev: stateAbbrev,
        search: normalizeText(label)
      });
    });
    options.sort(function (a, b) { return a.label.localeCompare(b.label); });
    return options;
  }

  function syncProjectCacheFromGlobal() {
    if (typeof window === 'undefined') return;
    if (Array.isArray(window.allProjects) && window.allProjects.length) {
      cmpProjectCache = window.allProjects.slice();
    }
  }

  function formatProjectOptionLabel(name) {
    var str = String(name || '').trim();
    if (!str) return '';
    var match = str.match(/(?:^|[\s(])(\d{2})[-\s]?(\d{2,3})(?:\))?\s*$/);
    if (!match) return str;
    var code = match[1] + '-' + match[2];
    var pattern = new RegExp('(?:^|[\\s(])' + match[1] + '[-\\s]?' + match[2] + '(?:\\))?\\s*$', 'i');
    var cleaned = str.replace(pattern, '').trim();
    return cleaned ? code + ' ' + cleaned : code;
  }

  function getProjectSortValue(value) {
    if (!value) return 0;
    var str = String(value);
    var match = str.match(/(^|[^0-9])(\d{2})\s*[-\u2010-\u2015\u2212]\s*(\d{3})(?![0-9])/);
    if (match && match[2]) {
      var major = parseInt(match[2], 10);
      var minor = parseInt(match[3], 10);
      if (!isNaN(major) && !isNaN(minor)) {
        return (major * 1000) + minor;
      }
    }
    return 0;
  }

  function compareProjectLabelsDesc(labelA, labelB) {
    var aVal = getProjectSortValue(labelA);
    var bVal = getProjectSortValue(labelB);
    var aHasCode = aVal > 0;
    var bHasCode = bVal > 0;
    if (aHasCode && bHasCode) {
      if (aVal !== bVal) return bVal - aVal;
      return String(labelA || '').localeCompare(String(labelB || ''));
    }
    if (aHasCode) return -1;
    if (bHasCode) return 1;
    return String(labelA || '').localeCompare(String(labelB || ''));
  }

  function normalizeCustomDates(list) {
    var out = [];
    var seen = {};
    (Array.isArray(list) ? list : []).forEach(function (value) {
      if (value == null) return;
      var date;
      if (value instanceof Date) {
        date = value;
      } else if (typeof value === "number") {
        date = new Date(value);
      } else {
        date = new Date(value);
      }
      if (!date || isNaN(date.getTime())) return;
      var iso = date.toISOString();
      if (!iso || seen[iso]) return;
      seen[iso] = true;
      out.push(iso);
    });
    out.sort();
    return out;
  }

  function createScheduleState() {
    return {
      main: { mode: "recurring", customDates: [] },
      followUps: {
        rfi: [],
        bids: []
      }
    };
  }

var FOLLOW_UP_KEYS = ['rfi', 'bids'];

var FOLLOW_UP_DEFAULT_SUBJECT = { rfi: "Reminder: RFI's Due", bids: "Reminder: Bids Due" };

  var CMP_SCHEDULE_MIN_LEAD_MS = 60 * 1000;
  var cmpFilterSelections = { city: [], state: [], project: [], division: [] };
  var cmpMultiSelectOptions = { city: null, state: null, project: null, division: null };
  var cmpMultiSelectOptionMap = { city: {}, state: {}, project: {}, division: {} };
  var cmpMultiSelectFields = {};
  var cmpActiveDropdownField = null;
  var cmpProjectCache = (typeof window !== 'undefined' && Array.isArray(window.allProjects)) ? window.allProjects.slice() : [];
  var cmpUsCityOptionCache = null;
  var cmpUsCityLoadPromise = null;
  var cmpDivisionOptionsSource = 'unknown';
  var cmpDivisionRefreshTimer = null;

  if (typeof window !== 'undefined') {
    window.__AB_CMP = window.__AB_CMP || {};
    window.__AB_CMP.getState = function () {
      return {
        cityDatasetSize: Array.isArray(window.AB_US_CITY_DATA) ? window.AB_US_CITY_DATA.length : 0,
        cityOptionCacheSize: cmpUsCityOptionCache ? cmpUsCityOptionCache.length : 0,
        multiSelectOptions: cmpMultiSelectOptions,
        divisionSource: cmpDivisionOptionsSource,
        filters: cmpFilterSelections
      };
    };
    window.ensureCampaignContactCache = ensureContactCache;
    window.getCampaignFilters = function () { return cmpFilterSelections; };
  }
  var cmpScheduleState = {
    edit: createScheduleState(),
    new: createScheduleState()
  };

  function setScheduleMode(context, mode) {
    var state = cmpScheduleState[context];
    if (!state) return;
    var normalizedMode = mode === "custom" ? "custom" : "recurring";
    state.main.mode = normalizedMode;
    if (context === "edit") {
      if (cmpScheduleModeRecurring) cmpScheduleModeRecurring.checked = normalizedMode === "recurring";
      if (cmpScheduleModeCustom) cmpScheduleModeCustom.checked = normalizedMode === "custom";
    } else {
      if (cmpScheduleModeNewRecurring) cmpScheduleModeNewRecurring.checked = normalizedMode === "recurring";
      if (cmpScheduleModeNewCustom) cmpScheduleModeNewCustom.checked = normalizedMode === "custom";
    }
    (cmpSchedulePanels[context] || []).forEach(function (panel) {
      if (!panel || typeof panel.classList === "undefined") return;
      var target = panel.getAttribute("data-panel");
      if (target === normalizedMode) {
        panel.classList.add("is-active");
      } else {
        panel.classList.remove("is-active");
      }
    });
  }

  function renderMainCustomList(context) {
    var state = cmpScheduleState[context];
    if (!state) return;
    state.main.customDates = normalizeCustomDates(state.main.customDates);
    var listEl = context === "edit" ? cmpScheduleCustomList : cmpScheduleNewCustomList;
    var emptyEl = context === "edit" ? cmpScheduleCustomEmpty : cmpScheduleNewCustomEmpty;
    if (!listEl) return;
    if (!state.main.customDates.length) {
      listEl.innerHTML = "";
      if (emptyEl) emptyEl.style.display = "";
      return;
    }
    if (emptyEl) emptyEl.style.display = "none";
    listEl.innerHTML = state.main.customDates.map(function (iso) {
      var label;
      try { label = new Date(iso).toLocaleString(); }
      catch (err) { label = iso; }
      return '<li class="cmp-schedule-custom-item" data-value="' + escapeHtml(iso) + '"><span>' + escapeHtml(label) + '</span><button type="button" class="link-btn cmp-schedule-remove" data-action="remove" data-value="' + escapeHtml(iso) + '">Remove</button></li>';
    }).join("");
  }

  function addCustomDate(context) {
    var input = context === "edit" ? cmpScheduleCustomInput : cmpScheduleNewCustomInput;
    if (!input) return;
    var raw = input.value;
    if (!raw) { alert("Choose a date and time."); return; }
    var date = new Date(raw);
    if (!date || isNaN(date.getTime())) { alert("Choose a valid date and time."); return; }
    if (date.getTime() <= Date.now() + 60000) { alert("Pick a time at least one minute in the future."); return; }
    var iso = date.toISOString();
    var state = cmpScheduleState[context];
    if (!state) return;
    if (state.main.customDates.indexOf(iso) !== -1) { alert("That date is already added."); input.value = ""; return; }
    state.main.customDates.push(iso);
    state.main.customDates.sort();
    input.value = "";
    renderMainCustomList(context);
    setScheduleMode(context, "custom");
  }

  function removeCustomDate(context, iso) {
    var state = cmpScheduleState[context];
    if (!state) return;
    state.main.customDates = (state.main.customDates || []).filter(function (value) { return value !== iso; });
    renderMainCustomList(context);
  }

  function handleMainCustomListClick(context, event) {
    var target = event && event.target;
    if (!target || typeof target.getAttribute !== "function") return;
    if (target.getAttribute("data-action") === "remove") {
      removeCustomDate(context, target.getAttribute("data-value"));
    }
  }

  function setFollowUpEnabled(context, key, enabled) {
    var ctrl = followUpControls[context] && followUpControls[context][key];
    if (!ctrl) return;
    if (ctrl.enable) ctrl.enable.checked = !!enabled;
    if (ctrl.block) {
      if (enabled) {
        ctrl.block.removeAttribute("hidden");
      } else {
        ctrl.block.setAttribute("hidden", "");
      }
    }
    if (enabled && ctrl.subject && !ctrl.subject.value && FOLLOW_UP_DEFAULT_SUBJECT[key]) {
      ctrl.subject.value = FOLLOW_UP_DEFAULT_SUBJECT[key];
    }
    renderFollowUpList(context, key);
  }

  function renderFollowUpList(context, key) {
    var ctrl = followUpControls[context] && followUpControls[context][key];
    if (!ctrl) return;
    var state = cmpScheduleState[context];
    if (!state) return;
    var list = state.followUps[key] = normalizeCustomDates(state.followUps[key] || []);
    if (!ctrl.list) return;
    if (!list.length) {
      ctrl.list.innerHTML = "";
      if (ctrl.empty) ctrl.empty.style.display = "";
      return;
    }
    if (ctrl.empty) ctrl.empty.style.display = "none";
    ctrl.list.innerHTML = list.map(function (iso) {
      var label;
      try { label = new Date(iso).toLocaleString(); }
      catch (err) { label = iso; }
      return '<li class="cmp-schedule-custom-item" data-value="' + escapeHtml(iso) + '"><span>' + escapeHtml(label) + '</span><button type="button" class="link-btn cmp-schedule-remove" data-action="remove" data-value="' + escapeHtml(iso) + '">Remove</button></li>';
    }).join("");
  }

  function addFollowUpDate(context, key) {
    var ctrl = followUpControls[context] && followUpControls[context][key];
    if (!ctrl) return;
    var input = ctrl.dateInput;
    if (!input) return;
    var raw = input.value;
    if (!raw) { alert("Choose a date and time."); return; }
    var date = new Date(raw);
    if (!date || isNaN(date.getTime())) { alert("Choose a valid date and time."); return; }
    if (date.getTime() <= Date.now() + 60000) { alert("Pick a time at least one minute in the future."); return; }
    var iso = date.toISOString();
    var state = cmpScheduleState[context];
    if (!state) return;
    var list = state.followUps[key] = state.followUps[key] || [];
    if (list.indexOf(iso) !== -1) { alert("That date is already added."); input.value = ""; return; }
    list.push(iso);
    list.sort();
    input.value = "";
    setFollowUpEnabled(context, key, true);
    renderFollowUpList(context, key);
  }

  function removeFollowUpDate(context, key, iso) {
    var state = cmpScheduleState[context];
    if (!state) return;
    state.followUps[key] = (state.followUps[key] || []).filter(function (value) { return value !== iso; });
    renderFollowUpList(context, key);
  }

  function handleFollowUpListClick(context, key, event) {
    var target = event && event.target;
    if (!target || typeof target.getAttribute !== "function") return;
    if (target.getAttribute("data-action") === "remove") {
      removeFollowUpDate(context, key, target.getAttribute("data-value"));
    }
  }

  function readScheduleFromUI(context) {
    var state = cmpScheduleState[context] || createScheduleState();
    var everyInput = context === "edit" ? cmpEveryDaysEdit : cmpEveryDays;
    var untilInput = context === "edit" ? cmpUntilEdit : cmpUntil;
    var everyVal = everyInput ? Number(everyInput.value || 0) : 0;
    if (state.main.mode === "recurring") {
      if (!everyVal || !isFinite(everyVal)) {
        alert("Enter how many days between sends (minimum 1).");
        if (everyInput) everyInput.focus();
        return null;
      }
    }
    var every = Math.max(1, Math.round(everyVal || 1));
    var untilIso = null;
    if (untilInput && untilInput.value) {
      var untilDate = new Date(untilInput.value);
      if (untilDate && !isNaN(untilDate.getTime())) {
        untilIso = untilDate.toISOString();
      }
    }
    return {
      mode: state.main.mode,
      everyDays: state.main.mode === "custom" ? 0 : every,
      until: state.main.mode === "custom" ? null : untilIso,
      customDates: state.main.customDates.slice()
    };
  }

  function readFollowUpsFromUI(context) {
    var result = {};
    FOLLOW_UP_KEYS.forEach(function (key) {
      var ctrl = followUpControls[context] && followUpControls[context][key];
      if (!ctrl) return;
      var enabled = ctrl.enable ? ctrl.enable.checked : false;
      var subject = ctrl.subject ? ctrl.subject.value : "";
      var html = getEditorValue(ctrl.htmlEditor, ctrl.html);
      var dates = (cmpScheduleState[context].followUps[key] || []).slice();
      result[key] = {
        enabled: enabled,
        subjectTpl: subject,
        htmlTpl: html,
        dates: dates
      };
    });
    return result;
  }

  function applyScheduleToUI(schedule, context) {
    var state = cmpScheduleState[context];
    if (!state) return;
    var mode = schedule && schedule.mode === "custom" ? "custom" : "recurring";
    state.main.mode = mode;
    state.main.customDates = normalizeCustomDates(schedule && schedule.customDates || []);
    if (context === "edit") {
      if (cmpEveryDaysEdit) cmpEveryDaysEdit.value = schedule && schedule.everyDays ? String(schedule.everyDays) : "";
      if (cmpUntilEdit) cmpUntilEdit.value = schedule && schedule.until ? schedule.until.split('T')[0] : "";
    } else {
      if (cmpEveryDays) cmpEveryDays.value = schedule && schedule.everyDays ? String(schedule.everyDays) : (cmpEveryDays && cmpEveryDays.value ? cmpEveryDays.value : "2");
      if (cmpUntil) cmpUntil.value = schedule && schedule.until ? schedule.until.split('T')[0] : "";
    }
    setScheduleMode(context, mode);
    renderMainCustomList(context);
  }

  function applyFollowUpsToUI(followUps, context) {
    var state = cmpScheduleState[context];
    if (!state) return;
    FOLLOW_UP_KEYS.forEach(function (key) {
      var ctrl = followUpControls[context] && followUpControls[context][key];
      if (!ctrl) return;
      var item = followUps && followUps[key] ? followUps[key] : {};
      var dates = normalizeCustomDates(item.dates || item.customDates || []);
      state.followUps[key] = dates;
      if (ctrl.subject && item.subjectTpl != null) ctrl.subject.value = item.subjectTpl;
      if (ctrl.html) setEditorValue(ctrl.htmlEditor, ctrl.html, item.htmlTpl || "");
      setFollowUpEnabled(context, key, !!item.enabled);
      renderFollowUpList(context, key);
    });
  }

  function resetScheduleUI(context) {
    cmpScheduleState[context] = createScheduleState();
    if (context === "edit") {
      if (cmpEveryDaysEdit) cmpEveryDaysEdit.value = "";
      if (cmpUntilEdit) cmpUntilEdit.value = "";
    } else {
      if (cmpEveryDays) cmpEveryDays.value = cmpEveryDays ? (cmpEveryDays.value || "2") : "2";
      if (cmpUntil) cmpUntil.value = "";
    }
    setScheduleMode(context, "recurring");
    renderMainCustomList(context);
    FOLLOW_UP_KEYS.forEach(function (key) {
      var ctrl = followUpControls[context] && followUpControls[context][key];
      if (!ctrl) return;
      if (ctrl.subject) ctrl.subject.value = "";
      if (ctrl.html) setEditorValue(ctrl.htmlEditor, ctrl.html, "");
      cmpScheduleState[context].followUps[key] = [];
      setFollowUpEnabled(context, key, false);
      renderFollowUpList(context, key);
    });
  }

  function toggleCampaignRequirements(hasCampaign) {
    var active = hasCampaign && cmpCurrentMode === "existing";
    cmpRequiresCampaignEls.forEach(function (el) { show(el, active); });
    if (active) {
      if (!cmpContactResultsData.length) {
        setContactResultsMessage("Run a search to see contacts.");
      }
    } else if (cmpCurrentMode === "existing") {
      setContactResultsMessage("Select a campaign to search contacts.");
    }
    if (cmpAddBtn) cmpAddBtn.disabled = !active;
  }

  Array.prototype.forEach.call(document.querySelectorAll('.rich-editor'), initRichEditor);

  // ---------- Messaging ----------

  function send(message, timeout) {
    timeout = timeout || 30000;

    // Prefer direct runtime messaging in the extension context
    if (typeof chrome !== "undefined" && chrome.runtime && typeof chrome.runtime.sendMessage === "function") {
      return new Promise(function (resolve) {
        try { chrome.runtime.sendMessage(message, function (response) { resolve(response); }); }
        catch (e) { resolve(null); }
      });
    }

    // Fallback bridge inside iframe -> background
    return new Promise(function (resolve, reject) {
      var finished = false;
      var id = String(Date.now()) + "_" + String(Math.random());
      var timer = setTimeout(function () {
        if (!finished) {
          finished = true;
          try { window.removeEventListener("message", handler); } catch (e) {}
          reject(new Error("timeout"));
        }
      }, timeout);

      function handler(event) {
        var data = event && event.data;
        if (!data || data.type !== "AB_RESPONSE" || data.messageId !== id) return;
        try { window.removeEventListener("message", handler); } catch (e) {}
        clearTimeout(timer);
        finished = true;
        resolve(data.response);
      }

      try { window.addEventListener("message", handler); } catch (e) {}
      try { window.postMessage({ type: "AB_FORWARD_TO_BACKGROUND", message: message, messageId: id }, "*"); } catch (e) {}
      try {
        if (window.parent && window.parent !== window) {
          window.parent.postMessage({ type: "AB_FORWARD_TO_BACKGROUND", message: message, messageId: id }, "*");
        }
      } catch (e) {}
    });
  }

  // ---------- Elements ----------
  var modeRadios = Array.prototype.slice.call(document.querySelectorAll('input[name="cmpMode"]') || []);
  var cmpSelectRow = $("cmpSelectRow");
  var cmpSelect = $("cmpSelect");
  var cmpLoaded = $("cmpLoaded");

  var cmpRefreshBtn = $("cmpRefreshBtn");
  var cmpDeleteBtn = $("cmpDeleteBtn");

  var cmpCreateBtn = $("cmpCreateBtn");
  var cmpName = $("cmpName");
  var cmpSubjectTpl = $("cmpSubjectTpl");
  var cmpHtmlTpl = $("cmpHtmlTpl");
  var cmpHtmlTplEditor = $("cmpHtmlTplEditor");
  var cmpEveryDays = $("cmpEveryDays");
  var cmpUntil = $("cmpUntil");
  var cmpNewRecipients = $("cmpNewRecipients");

  var cmpAddEmails = $("cmpAddEmails");
  var cmpAddBtn = $("cmpAddBtn");
var cmpRecipients = $("cmpRecipients");
var cmpModeExistingEls = Array.prototype.slice.call(document.querySelectorAll('.cmp-mode-existing') || []);
var cmpModeNewEls = Array.prototype.slice.call(document.querySelectorAll('.cmp-mode-new') || []);
var cmpRequiresCampaignEls = Array.prototype.slice.call(document.querySelectorAll('.cmp-requires-campaign') || []);

var cmpSuccessModal = $("cmpSuccessModal");
var cmpSuccessMessage = $("cmpSuccessMessage");
var cmpSuccessContinueBtn = $("cmpSuccessContinueBtn");

var cmpScheduleModeRecurring = $("cmpScheduleModeRecurring");
var cmpScheduleModeCustom = $("cmpScheduleModeCustom");
var cmpScheduleCustomInput = $("cmpScheduleCustomInput");
var cmpScheduleCustomAddBtn = $("cmpScheduleCustomAddBtn");
var cmpScheduleCustomList = $("cmpScheduleCustomList");
var cmpScheduleCustomEmpty = $("cmpScheduleCustomEmpty");
var cmpScheduleModeNewRecurring = $("cmpScheduleModeNewRecurring");
var cmpScheduleModeNewCustom = $("cmpScheduleModeNewCustom");
var cmpScheduleNewCustomInput = $("cmpScheduleNewCustomInput");
var cmpScheduleNewCustomAddBtn = $("cmpScheduleNewCustomAddBtn");
var cmpScheduleNewCustomList = $("cmpScheduleNewCustomList");
var cmpScheduleNewCustomEmpty = $("cmpScheduleNewCustomEmpty");

var followUpControls = {
  new: {
    rfi: {
      enable: $("cmpRfiEnableNew"),
      block: $("cmpRfiBlockNew"),
      subject: $("cmpRfiSubjectNew"),
      html: $("cmpRfiHtmlNew"),
      htmlEditor: $("cmpRfiHtmlNewEditor"),
      dateInput: $("cmpRfiDateNew"),
      addBtn: $("cmpRfiAddDateNew"),
      list: $("cmpRfiListNew"),
      empty: $("cmpRfiEmptyNew")
    },
    bids: {
      enable: $("cmpBidsEnableNew"),
      block: $("cmpBidsBlockNew"),
      subject: $("cmpBidsSubjectNew"),
      html: $("cmpBidsHtmlNew"),
      htmlEditor: $("cmpBidsHtmlNewEditor"),
      dateInput: $("cmpBidsDateNew"),
      addBtn: $("cmpBidsAddDateNew"),
      list: $("cmpBidsListNew"),
      empty: $("cmpBidsEmptyNew")
    }
  },
  edit: {
    rfi: {
      enable: $("cmpRfiEnable"),
      block: $("cmpRfiBlock"),
      subject: $("cmpRfiSubjectEdit"),
      html: $("cmpRfiHtmlEdit"),
      htmlEditor: $("cmpRfiHtmlEditEditor"),
      dateInput: $("cmpRfiDateEdit"),
      addBtn: $("cmpRfiAddDateEdit"),
      list: $("cmpRfiListEdit"),
      empty: $("cmpRfiEmptyEdit")
    },
    bids: {
      enable: $("cmpBidsEnable"),
      block: $("cmpBidsBlock"),
      subject: $("cmpBidsSubjectEdit"),
      html: $("cmpBidsHtmlEdit"),
      htmlEditor: $("cmpBidsHtmlEditEditor"),
      dateInput: $("cmpBidsDateEdit"),
      addBtn: $("cmpBidsAddDateEdit"),
      list: $("cmpBidsListEdit"),
      empty: $("cmpBidsEmptyEdit")
    }
  }
};

var cmpSchedulePanels = {
  edit: Array.prototype.slice.call(document.querySelectorAll('.cmp-schedule-panel[data-context="edit"]') || []),
  new: Array.prototype.slice.call(document.querySelectorAll('.cmp-schedule-panel[data-context="new"]') || [])
};

var cmpCurrentMode = "existing";


  var cmpContactSearchForm = $("cmpContactSearchForm");
  var cmpSearchZip = $("cmpSearchZip");
  var cmpSearchRadius = $("cmpSearchRadius");
  var cmpSearchCity = $("cmpSearchCity");
  var cmpSearchCityInput = $("cmpSearchCityInput");
  var cmpSearchCityDropdown = $("cmpSearchCityDropdown");
  var cmpSearchCityChips = $("cmpSearchCityChips");
  var cmpSearchState = $("cmpSearchState");
  var cmpSearchStateInput = $("cmpSearchStateInput");
  var cmpSearchStateDropdown = $("cmpSearchStateDropdown");
  var cmpSearchStateChips = $("cmpSearchStateChips");
  var cmpSearchProject = $("cmpSearchProject");
  var cmpSearchProjectInput = $("cmpSearchProjectInput");
  var cmpSearchProjectDropdown = $("cmpSearchProjectDropdown");
  var cmpSearchProjectChips = $("cmpSearchProjectChips");
  var cmpSearchDivision = $("cmpSearchDivision");
  var cmpSearchDivisionInput = $("cmpSearchDivisionInput");
  var cmpSearchDivisionDropdown = $("cmpSearchDivisionDropdown");
  var cmpSearchDivisionChips = $("cmpSearchDivisionChips");
  var cmpContactResultsContainer = $("cmpContactResults");
  var cmpContactResultsCount = $("cmpContactResultsCount");
  var cmpContactResetBtn = $("cmpContactResetBtn");
  var cmpContactSelectAll = $("cmpContactSelectAll");
  var cmpContactClearSelection = $("cmpContactClearSelection");
  var cmpAddSelectedContactsBtn = $("cmpAddSelectedContacts");

  var cmpSubjectTplEdit = $("cmpSubjectTplEdit");
  var cmpHtmlTplEdit = $("cmpHtmlTplEdit");
  var cmpHtmlTplEditEditor = $("cmpHtmlTplEditEditor");

  var cmpEveryDaysEdit = $("cmpEveryDaysEdit");
  var cmpUntilEdit = $("cmpUntilEdit");
  var cmpNextSend = $("cmpNextSend");

  var cmpSaveBtn = $("cmpSaveBtn");
  var cmpSendNowBtn = $("cmpSendNowBtn");
  var cmpToggleBtn = $("cmpToggleBtn");

  var kpiTotal = $("cmpKpiTotal");
  var kpiPending = $("cmpKpiPending");
  var kpiResp = $("cmpKpiResp");
  var kpiOpen = $("cmpKpiOpen");
  var kpiClick = $("cmpKpiClick");
  var kpiLast = $("cmpKpiLast");
  var kpiNext = $("cmpKpiNext");
  var cmpRecipientSummary = $("cmpRecipientSummary");

  // Review modal
  var modal = document.getElementById("cmpSendModal");
  var modalClose = document.getElementById("cmpSendModalClose");
  var pendingList = document.getElementById("cmpPendingList");
  var pendingCount = document.getElementById("cmpPendingCount");
  var checkAllBtn = document.getElementById("cmpCheckAll");
  var uncheckAllBtn = document.getElementById("cmpUncheckAll");
  var sendSelectedBtn = document.getElementById("cmpSendSelectedBtn");
  var cmpSendNotice = document.getElementById("cmpSendNotice");

  cmpMultiSelectFields = {
    city: {
      input: cmpSearchCityInput,
      dropdown: cmpSearchCityDropdown,
      chips: cmpSearchCityChips,
      hidden: cmpSearchCity,
      loader: loadCityOptions,
      maxResults: 500
    },
    state: {
      input: cmpSearchStateInput,
      dropdown: cmpSearchStateDropdown,
      chips: cmpSearchStateChips,
      hidden: cmpSearchState,
      loader: loadStateOptions,
      maxResults: 60
    },
    project: {
      input: cmpSearchProjectInput,
      dropdown: cmpSearchProjectDropdown,
      chips: cmpSearchProjectChips,
      hidden: cmpSearchProject,
      loader: loadProjectOptions,
      maxResults: 400
    },
    division: {
      input: cmpSearchDivisionInput,
      dropdown: cmpSearchDivisionDropdown,
      chips: cmpSearchDivisionChips,
      hidden: cmpSearchDivision,
      loader: loadDivisionOptions,
      maxResults: 400
    }
  };
  Object.keys(cmpMultiSelectFields).forEach(function (field) {
    initMultiSelect(field);
  });

  document.addEventListener('click', function (evt) {
    var target = evt && evt.target;
    if (!target) return;
    if (target.closest && target.closest('.cmp-multiselect')) return;
    closeAllMultiSelectDropdowns();
  });

  preloadCampaignFilters();

  // History UI
  var cmpHistoryToggle = $("cmpHistoryToggle");
  var cmpSendHistory = $("cmpSendHistory");
  var cmpScheduleSummary = $("cmpScheduleSummary");
  function renderHistory() {
    if (!cmpSendHistory) return;
    cmpSendHistory.innerHTML = "";
    var list = (current && current.sentHistory) ? current.sentHistory.slice().sort(function(a,b){return b-a;}) : [];
    if (!list.length) {
      var li = document.createElement("li");
      li.textContent = "No sends yet.";
      cmpSendHistory.appendChild(li);
      return;
    }
    for (var i=0;i<list.length;i++) {
      var ts = list[i];
      var li = document.createElement("li");
      li.textContent = fmtDate(ts);
      cmpSendHistory.appendChild(li);
    }
  }
  if (cmpHistoryToggle) cmpHistoryToggle.addEventListener("click", function(){
    if (!cmpSendHistory) return;
    var isHidden = cmpSendHistory.hasAttribute("hidden");
    if (isHidden) {
      renderHistory();
      cmpSendHistory.removeAttribute("hidden");
      cmpHistoryToggle.textContent = "â€“ Hide history";
    } else {
      cmpSendHistory.setAttribute("hidden","");
      cmpHistoryToggle.textContent = "+ Show history";
    }
  });
// ---------- State ----------
  var campaigns = [];
  var current = null;

  var cmpContactCache = null;
  var cmpContactsLoadingPromise = null;
  var cmpContactResultsData = [];
  var cmpContactDisplayedData = [];
  var cmpContactDisplayedMap = {};
  var cmpContactSelectedMap = {};
  var cmpContactLastTotal = 0;
  var cmpContactLastDisplayed = 0;
  var cmpZipCacheLocal = {};
  var cmpZipLookupPending = {};

  function syncZipCacheFromGlobal() {
    if (typeof window === "undefined") return;
    if (!window.zipCoordCache || typeof window.zipCoordCache !== "object") return;
    for (var key in window.zipCoordCache) {
      if (Object.prototype.hasOwnProperty.call(window.zipCoordCache, key) && !cmpZipCacheLocal[key]) {
        cmpZipCacheLocal[key] = window.zipCoordCache[key];
      }
    }
  }

  function storeZipInGlobal(key, location) {
    if (typeof window === "undefined") return;
    if (!window.zipCoordCache || typeof window.zipCoordCache !== "object") {
      window.zipCoordCache = {};
    }
    window.zipCoordCache[key] = location;
  }

  function lookupZipCoords(zip) {
    var key = sanitizeZip(zip);
    if (!key) return Promise.resolve(null);
    syncZipCacheFromGlobal();
    if (cmpZipCacheLocal[key]) return Promise.resolve(cmpZipCacheLocal[key]);
    if (cmpZipLookupPending[key]) return cmpZipLookupPending[key];
    var promise = send({ type: "AB_LOOKUP_ZIP", zip: key }).then(function (res) {
      if (res && res.success && res.location) {
        cmpZipCacheLocal[key] = res.location;
        storeZipInGlobal(key, res.location);
        return res.location;
      }
      return null;
    }).catch(function () { return null; });
    cmpZipLookupPending[key] = promise;
    var clear = function () { delete cmpZipLookupPending[key]; };
    promise.then(clear, clear);
    return promise;
  }

  function computeDistanceMiles(pointA, pointB) {
    if (!pointA || !pointB) return Infinity;
    var lat1 = Number(pointA.lat != null ? pointA.lat : pointA.latitude);
    var lon1 = Number(pointA.lon != null ? pointA.lon : (pointA.lng != null ? pointA.lng : pointA.longitude));
    var lat2 = Number(pointB.lat != null ? pointB.lat : pointB.latitude);
    var lon2 = Number(pointB.lon != null ? pointB.lon : (pointB.lng != null ? pointB.lng : pointB.longitude));
    if (!isFinite(lat1) || !isFinite(lon1) || !isFinite(lat2) || !isFinite(lon2)) return Infinity;
    var toRad = function (deg) { return deg * (Math.PI / 180); };
    var dLat = toRad(lat2 - lat1);
    var dLon = toRad(lon2 - lon1);
    var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
      Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
      Math.sin(dLon / 2) * Math.sin(dLon / 2);
    var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return 3958.8 * c;
  }

  function ensureContactCache() {
    if (cmpContactCache && cmpContactCache.length) return Promise.resolve(cmpContactCache);
    if (typeof window !== "undefined" && Array.isArray(window.allContacts) && window.allContacts.length) {
      cmpContactCache = window.allContacts.slice();
      syncZipCacheFromGlobal();
      return Promise.resolve(cmpContactCache);
    }
    if (cmpContactsLoadingPromise) return cmpContactsLoadingPromise;
    cmpContactsLoadingPromise = send({ type: "AB_GET_CACHED_DATA" }).then(function (res) {
      if (res && res.success) {
        cmpContactCache = Array.isArray(res.contacts) ? res.contacts : [];
        if (Array.isArray(res.projects) && res.projects.length) {
          cmpProjectCache = res.projects.slice();
        }
        if (res.zipCache && typeof res.zipCache === "object") {
          cmpZipCacheLocal = {};
          for (var key in res.zipCache) {
            if (Object.prototype.hasOwnProperty.call(res.zipCache, key)) {
              cmpZipCacheLocal[key] = res.zipCache[key];
              storeZipInGlobal(key, res.zipCache[key]);
            }
          }
        }
        // Reload project and division options now that we have data
        loadProjectOptions(true).catch(function() {});
        loadDivisionOptions(true).catch(function() {});
      } else {
        cmpContactCache = [];
      }
      return cmpContactCache;
    }).catch(function () {
      cmpContactCache = [];
      return cmpContactCache;
    });
    cmpContactsLoadingPromise.then(function () { cmpContactsLoadingPromise = null; }, function () { cmpContactsLoadingPromise = null; });
    return cmpContactsLoadingPromise;
  }

  function collectProjectTokens(contact) {
    var tokens = [];
    if (!contact || typeof contact !== "object") return tokens;
    var fields = ["projectDisplayName", "projectName", "project", "projectFolderName", "projectFolder", "projectCity", "projectKey", "projectNumber"];
    for (var i = 0; i < fields.length; i++) {
      var val = contact[fields[i]];
      if (!val && contact.projectInfo && contact.projectInfo[fields[i]]) {
        val = contact.projectInfo[fields[i]];
      }
      if (!val) continue;
      var norm = normalizeText(val);
      if (norm && tokens.indexOf(norm) === -1) tokens.push(norm);
      var normalizedKey = normalizeProjectKey(val);
      if (normalizedKey && tokens.indexOf(normalizedKey) === -1) tokens.push(normalizedKey);
    }
    return tokens;
  }

  function shapeContact(contact) {
    if (!contact || typeof contact !== "object") return null;
    var emailRaw = firstNonEmpty(contact, ["email", "Email", "emailAddress", "emailNormalized"]);
    var email = normalizeText(emailRaw);
    if (!email) return null;
    var company = firstNonEmpty(contact, ["company", "Company", "companyName", "CompanyName"]);
    var name = firstNonEmpty(contact, ["name", "Name", "contactName", "ContactName"]);
    var city = firstNonEmpty(contact, ["city", "City", "town"]);
    var state = firstNonEmpty(contact, ["state", "State", "province", "Region"]);
    var zipRaw = firstNonEmpty(contact, ["zip", "Zip", "postal", "Postal", "zipCode", "ZipCode", "zipcode"]);
    var projectValue = firstNonEmpty(contact, ["projectDisplayName", "projectName", "project", "projectFolderName"]);
    var division = firstNonEmpty(contact, ["division", "Division", "divisionName", "divisionKey", "sheetName"]);
    var divisionKey = normalizeText(division);
    var tokens = collectProjectTokens(contact);

    if (!projectValue && tokens.length) {

      projectValue = contact.projectDisplayName || contact.projectName || contact.project || "";

    }

    var normalizedProject = normalizeProjectKey(projectValue || '');

    if (normalizedProject && tokens.indexOf(normalizedProject) === -1) {

      tokens.push(normalizedProject);

    }

    var cityKeyRaw = sanitizeCityLabel(city) || city;
    var cityKey = normalizeText(cityKeyRaw);

    return {

      email: email,

      emailRaw: emailRaw || email,

      company: company,

      name: name,

      city: city,

      cityKey: cityKey,

      state: state,

      stateKey: normalizeText(state),

      stateAbbrev: state ? String(state).trim().toUpperCase() : "",

      zip: zipRaw,

      zipKey: sanitizeZip(zipRaw),

      project: projectValue,

      projectKey: normalizedProject,

      division: division,

      divisionKey: divisionKey,

      projectTokens: tokens,

      original: contact

    };
  }

  function projectMatches(shaped, query) {

    if (!query) return true;

    var normalizedQuery = normalizeProjectKey(query);

    var tokens = shaped.projectTokens || [];

    for (var i = 0; i < tokens.length; i++) {

      var token = tokens[i];

      if (token && token.indexOf(normalizedQuery) !== -1) return true;

    }

    if (shaped.projectKey && shaped.projectKey.indexOf(normalizedQuery) !== -1) return true;

    return false;

  }



  function ensureUsCityOptionList() {
    if (cmpUsCityOptionCache && cmpUsCityOptionCache.length) return Promise.resolve(cmpUsCityOptionCache);
    if (typeof window !== "undefined" && Array.isArray(window.AB_US_CITY_DATA) && window.AB_US_CITY_DATA.length) {
      var datasetOptions = buildUsCityDatasetOptions(window.AB_US_CITY_DATA);
      if (datasetOptions && datasetOptions.length) {
        cmpUsCityOptionCache = datasetOptions;
        window.AB_US_CITY_OPTIONS = datasetOptions;
        return Promise.resolve(datasetOptions);
      }
    }
    if (cmpUsCityLoadPromise) return cmpUsCityLoadPromise;
    var url = getRuntimeAssetUrl('data/us_cities_min.json');
    cmpUsCityLoadPromise = fetch(url).then(function (response) {
      if (!response || !response.ok) throw new Error('Failed to load US city list');
      return response.json();
    }).then(function (payload) {
      var options = buildUsCityDatasetOptions(payload);
      cmpUsCityOptionCache = options;
      if (typeof window !== 'undefined') {
        if (Array.isArray(payload)) {
          window.AB_US_CITY_DATA = payload;
        }
        window.AB_US_CITY_OPTIONS = options;
      }
      return options;
    }).catch(function (error) {
      console.warn('[Campaigns] Unable to load packaged US city list:', error);
      cmpUsCityOptionCache = [];
      return [];
    }).then(function (result) {
      cmpUsCityLoadPromise = null;
      return result;
    });
    return cmpUsCityLoadPromise;
  }

  function buildCityOptionsFromContacts() {
    return ensureContactCache().then(function (contacts) {
      var map = {};
      (contacts || []).forEach(function (contact) {
        var shaped = shapeContact(contact);
        if (!shaped || !shaped.cityKey) return;
        var stateAbbrev = shaped.stateAbbrev || '';
        if (stateAbbrev === 'HI') return;
        var key = shaped.cityKey + '|' + stateAbbrev;
        if (map[key]) return;
        var labelCity = sanitizeCityLabel(shaped.city || shaped.cityKey);
        var label = labelCity || shaped.cityKey;
        if (stateAbbrev) label = label + ', ' + stateAbbrev;
        map[key] = {
          key: key,
          label: label,
          value: shaped.cityKey,
          stateAbbrev: stateAbbrev,
          search: normalizeText(label + ' ' + (shaped.state || ''))
        };
      });
      var list = Object.keys(map).map(function (key) { return map[key]; });
      list.sort(function (a, b) { return a.label.localeCompare(b.label); });
      return list;
    }).catch(function () {
      return [];
    });
  }

  function loadCityOptions(forceRefresh) {
    if (!forceRefresh && cmpMultiSelectOptions.city && cmpMultiSelectOptions.city.length && cmpMultiSelectOptionMap.city && Object.keys(cmpMultiSelectOptionMap.city).length) return Promise.resolve(cmpMultiSelectOptions.city);
    return ensureUsCityOptionList().then(function (options) {
      if (options && options.length) {
        var map = {};
        options.forEach(function (option) { map[option.key] = option; });
        cmpMultiSelectOptions.city = options;
        cmpMultiSelectOptionMap.city = map;
        console.log('[CMP] City options loaded from US dataset:', options.length);
        return options;
      }
      return buildCityOptionsFromContacts().then(function (fallback) {
        var fallbackMap = {};
        fallback.forEach(function (option) { fallbackMap[option.key] = option; });
        if (fallback.length) {
          cmpMultiSelectOptions.city = fallback;
          console.log('[CMP] City options loaded from contacts:', fallback.length);
        } else {
          cmpMultiSelectOptions.city = null;
          console.warn('[CMP] No city options available');
        }
        cmpMultiSelectOptionMap.city = fallbackMap;
        return fallback;
      });
    }).catch(function (err) {
      console.error('[CMP] Failed to load city options:', err);
      cmpMultiSelectOptions.city = null;
      cmpMultiSelectOptionMap.city = {};
      return [];
    });
  }

  function loadStateOptions() {
    if (cmpMultiSelectOptions.state) return Promise.resolve(cmpMultiSelectOptions.state);
    return ensureContactCache().then(function (contacts) {
      var map = {};
      (contacts || []).forEach(function (contact) {
        var shaped = shapeContact(contact);
        if (!shaped) return;
        var label = shaped.state || shaped.stateAbbrev;
        if (!label) return;
        var normalized = normalizeText(label);
        if (!normalized) return;
        var abbr = shaped.stateAbbrev ? shaped.stateAbbrev.toUpperCase() : '';
        if (abbr === 'HI' || normalized === 'hawaii') return;
        var key = abbr || normalized;
        if (map[key]) return;
        var display = abbr && label.toUpperCase().indexOf(abbr) === -1 ? label + ' (' + abbr + ')' : label;
        map[key] = {
          key: key,
          label: display,
          value: normalized,
          abbr: abbr,
          search: normalizeText(display)
        };
      });
      var list = Object.keys(map).map(function (key) { return map[key]; });
      list.sort(function (a, b) { return a.label.localeCompare(b.label); });
      cmpMultiSelectOptions.state = list;
      cmpMultiSelectOptionMap.state = map;
      return list;
    });
  }

  function ensureProjectList() {
    syncProjectCacheFromGlobal();
    if (typeof window !== "undefined" && typeof window.buildProjectNameList === "function") {
      var source = Array.isArray(window.allProjects) && window.allProjects.length ? window.allProjects : cmpProjectCache;
      if (Array.isArray(source) && source.length) {
        var builtList = window.buildProjectNameList(source) || [];
        return builtList.slice().sort(compareProjectLabelsDesc);
      }
    }
    var names = [];
    (cmpProjectCache || []).forEach(function (proj) {
      var label = proj && (proj.displayName || proj.projectDisplayName || proj.name || proj.projectName);
      if (!label) return;
      names.push(label);
    });
    names = names.filter(function (val, idx, arr) { return val && arr.indexOf(val) === idx; });
    names.sort(compareProjectLabelsDesc);
    return names;
  }

  function loadProjectOptions(forceRefresh) {
    if (!forceRefresh && cmpMultiSelectOptions.project && cmpMultiSelectOptions.project.length && cmpMultiSelectOptionMap.project && Object.keys(cmpMultiSelectOptionMap.project).length) return Promise.resolve(cmpMultiSelectOptions.project);
    
    // First sync from global if available
    syncProjectCacheFromGlobal();
    
    var prepare = (typeof window !== 'undefined' && Array.isArray(window.allProjects) && window.allProjects.length) || (cmpProjectCache && cmpProjectCache.length)
      ? Promise.resolve()
      : ensureContactCache().catch(function (err) { 
          console.warn('[CMP] Failed to load contact cache for projects:', err); 
          return []; 
        });
    return prepare.then(function () {
      // Sync again after potential contact cache load
      syncProjectCacheFromGlobal();
      
      var names = ensureProjectList();
      console.log('[CMP] Building project options from', names.length, 'projects');
      
      var map = {};
      names.forEach(function (name) {
        var formatted;
        if (typeof window !== "undefined" && typeof window.labelProjectCodeFirst === "function") {
          formatted = window.labelProjectCodeFirst(name);
        } else {
          formatted = formatProjectOptionLabel(name);
        }
        if (!formatted) formatted = String(name || '');
        var normalized = normalizeText(formatted);
        if (!normalized || map[normalized]) return;
        map[normalized] = {
          key: normalized,
          label: formatted,
          value: normalized,
          search: normalizeText(formatted)
        };
      });
      var list = Object.keys(map).map(function (key) { return map[key]; });
      list.sort(function (a, b) { return a.label.localeCompare(b.label); });
      if (list.length) {
        cmpMultiSelectOptions.project = list;
        cmpMultiSelectOptionMap.project = map;
        console.log('[CMP] Project options loaded:', list.length, 'options in map');
      } else {
        cmpMultiSelectOptions.project = null;
        cmpMultiSelectOptionMap.project = {};
        console.warn('[CMP] No project options could be built');
      }
      return list;
    }).catch(function (err) {
      console.error('[CMP] Failed to load project options:', err);
      cmpMultiSelectOptions.project = null;
      cmpMultiSelectOptionMap.project = {};
      return [];
    });
  }

  function buildDivisionOptionsFromMapping() {
    if (typeof window === "undefined" || typeof window.getDivisionFolderData !== "function") return [];
    var source = window.getDivisionFolderData();
    var map = {};
    if (source && typeof source === 'object') {
      Object.keys(source).forEach(function (siteKey) {
        var divisions = source[siteKey];
        if (!divisions || typeof divisions !== 'object') return;
        Object.keys(divisions).forEach(function (divisionName) {
          var mapped = divisions[divisionName];
          var label = mapped || divisionName;
          var normalized = normalizeText(label);
          if (!normalized || map[normalized]) return;
          map[normalized] = {
            key: normalized,
            label: label,
            value: normalized,
            search: normalizeText(label)
          };
        });
      });
    }
    return Object.keys(map).map(function (key) { return map[key]; }).sort(function (a, b) { return a.label.localeCompare(b.label); });
  }

  function buildDivisionOptionsFromContacts() {
    return ensureContactCache().then(function (contacts) {
      var map = {};
      (contacts || []).forEach(function (contact) {
        var shaped = shapeContact(contact);
        if (!shaped || !shaped.divisionKey) return;
        if (map[shaped.divisionKey]) return;
        var label = shaped.division || shaped.divisionKey;
        map[shaped.divisionKey] = {
          key: shaped.divisionKey,
          label: label,
          value: shaped.divisionKey,
          search: normalizeText(label)
        };
      });
      var list = Object.keys(map).map(function (key) { return map[key]; });
      list.sort(function (a, b) { return compareProjectLabelsDesc(a.label, b.label); });
      return list;
    }).catch(function () {
      return [];
    });
  }

  function waitForDivisionMapping(maxAttempts, delayMs) {
    return new Promise(function (resolve) {
      var attempts = 0;
      function check() {
        var options = buildDivisionOptionsFromMapping();
        if (options && options.length) {
          resolve(options);
          return;
        }
        attempts += 1;
        if (attempts > maxAttempts) {
          resolve(null);
          return;
        }
        setTimeout(check, delayMs);
      }
      check();
    });
  }

  function scheduleDivisionMappingRefresh() {
    if (cmpDivisionOptionsSource !== 'fallback') return;
    if (cmpDivisionRefreshTimer) return;
    cmpDivisionRefreshTimer = setTimeout(function () {
      cmpDivisionRefreshTimer = null;
      waitForDivisionMapping(15, 300).then(function (options) {
        if (options && options.length) {
          var map = {};
          options.forEach(function (option) { map[option.key] = option; });
          cmpDivisionOptionsSource = 'mapping';
          cmpMultiSelectOptions.division = options;
          cmpMultiSelectOptionMap.division = map;
          renderMultiSelectChips('division');
        } else if (cmpDivisionOptionsSource === 'fallback') {
          scheduleDivisionMappingRefresh();
        }
      });
    }, 1500);
  }

  function loadDivisionOptions(forceRefresh) {
    if (!forceRefresh && cmpMultiSelectOptions.division && cmpMultiSelectOptions.division.length && cmpDivisionOptionsSource === 'mapping') {
      return Promise.resolve(cmpMultiSelectOptions.division);
    }
    return waitForDivisionMapping(10, 200).then(function (mappingOptions) {
      if (mappingOptions && mappingOptions.length) {
        var mappingMap = {};
        mappingOptions.forEach(function (option) { mappingMap[option.key] = option; });
        cmpDivisionOptionsSource = 'mapping';
        cmpMultiSelectOptions.division = mappingOptions;
        cmpMultiSelectOptionMap.division = mappingMap;
        return mappingOptions;
      }
      return buildDivisionOptionsFromContacts().then(function (fallback) {
        var fallbackMap = {};
        fallback.forEach(function (option) { fallbackMap[option.key] = option; });
        if (fallback.length) {
          cmpMultiSelectOptions.division = fallback;
        } else {
          cmpMultiSelectOptions.division = null;
        }
        cmpDivisionOptionsSource = 'fallback';
        cmpMultiSelectOptionMap.division = fallbackMap;
        scheduleDivisionMappingRefresh();
        return fallback;
      });
    });
  }

  function preloadCampaignFilters() {
    console.log('[CMP] Preloading campaign filters...');
    Promise.all([
      loadCityOptions(true).catch(function(e) { console.error('[CMP] City options failed:', e); }),
      loadStateOptions().catch(function(e) { console.error('[CMP] State options failed:', e); }),
      loadDivisionOptions().catch(function(e) { console.error('[CMP] Division options failed:', e); }),
      loadProjectOptions(true).catch(function(e) { console.error('[CMP] Project options failed:', e); })
    ]).then(function() {
      console.log('[CMP] Filter preload complete');
      console.log('[CMP] - City options:', cmpMultiSelectOptions.city ? cmpMultiSelectOptions.city.length : 0);
      console.log('[CMP] - State options:', cmpMultiSelectOptions.state ? cmpMultiSelectOptions.state.length : 0);
      console.log('[CMP] - Division options:', cmpMultiSelectOptions.division ? cmpMultiSelectOptions.division.length : 0);
      console.log('[CMP] - Project options:', cmpMultiSelectOptions.project ? cmpMultiSelectOptions.project.length : 0);
    });
  }

  function getSelectedFilters(field) {
    var list = cmpFilterSelections[field] || [];
    return list.slice();
  }

  function updateMultiSelectHiddenValue(field) {
    var cfg = cmpMultiSelectFields[field];
    if (!cfg || !cfg.hidden) return;
    var labels = (cmpFilterSelections[field] || []).map(function (item) { return item.label; });
    cfg.hidden.value = labels.join(', ');
  }

  function renderMultiSelectChips(field) {
    var cfg = cmpMultiSelectFields[field];
    if (!cfg || !cfg.chips) return;
    var selections = cmpFilterSelections[field] || [];
    if (!selections.length) {
      cfg.chips.innerHTML = '';
      updateMultiSelectHiddenValue(field);
      return;
    }
    cfg.chips.innerHTML = selections.map(function (entry) {
      return '<span class="cmp-chip" data-key="' + entry.key + '">' + escapeHtml(entry.label) + '<button type="button" class="cmp-chip-remove" data-remove="' + entry.key + '">&times;</button></span>';
    }).join('');
    updateMultiSelectHiddenValue(field);
  }

  function closeMultiSelectDropdown(field) {
    var cfg = cmpMultiSelectFields[field];
    if (cfg && cfg.dropdown) {
      cfg.dropdown.classList.remove('is-open');
      cfg.dropdown.innerHTML = '';
    }
    if (cmpActiveDropdownField === field) {
      cmpActiveDropdownField = null;
    }
  }

  function closeAllMultiSelectDropdowns(exceptField) {
    Object.keys(cmpMultiSelectFields).forEach(function (name) {
      if (name === exceptField) return;
      closeMultiSelectDropdown(name);
    });
  }

  function openMultiSelectDropdown(field, query) {
    var cfg = cmpMultiSelectFields[field];
    if (!cfg) return;
    closeAllMultiSelectDropdowns(field);
    cmpActiveDropdownField = field;
    renderMultiSelectDropdown(field, query);
  }

  function filterMultiOptions(options, query, selections, maxResults) {
    var search = normalizeText(query || '');
    var taken = new Set((selections || []).map(function (item) { return item.key; }));
    var limit = (typeof maxResults === 'number' && maxResults > 0) ? maxResults : 12;
    var filtered = options.filter(function (option) {
      if (!option || taken.has(option.key)) return false;
      if (!search) return true;
      return option.search.indexOf(search) !== -1;
    });
    if (filtered.length > limit) {
      return filtered.slice(0, limit);
    }
    return filtered;
  }

  function renderMultiSelectDropdown(field, query) {
    var cfg = cmpMultiSelectFields[field];
    if (!cfg || !cfg.dropdown) return;
    var loader = cfg.loader;
    if (typeof loader !== 'function') return;
    loader().then(function (options) {
      if (cmpActiveDropdownField !== field) return;
      var selections = cmpFilterSelections[field] || [];
      var filtered = filterMultiOptions(options || [], query || '', selections, cfg.maxResults);
      if (!filtered.length) {
        closeMultiSelectDropdown(field);
        return;
      }
      cfg.dropdown.innerHTML = filtered.map(function (option) {
        var meta = option.stateAbbrev || option.abbr || '';
        var metaHtml = meta ? '<div class="cmp-multiselect-option-meta">' + escapeHtml(meta) + '</div>' : '';
        var titleText = option.label || '';
        if (meta) titleText += ' (' + meta + ')';
        var optionTitle = escapeHtml(titleText);
        return '<div class="cmp-multiselect-option" data-key="' + option.key + '" title="' + optionTitle + '"><div class="cmp-multiselect-option-label">' + escapeHtml(option.label) + '</div>' + metaHtml + '</div>';
      }).join('');
      cfg.dropdown.classList.add('is-open');
    });
  }

  function selectMultiOption(field, key) {
    if (!key) return;
    var cfg = cmpMultiSelectFields[field];
    var options = cmpMultiSelectOptionMap[field] || {};
    var option = options[key];
    if (!cfg || !option) return;
    var selections = cmpFilterSelections[field] || [];
    var exists = selections.some(function (entry) { return entry.key === key; });
    if (exists) return;
    selections.push(option);
    cmpFilterSelections[field] = selections;
    renderMultiSelectChips(field);
    closeMultiSelectDropdown(field);
    if (cfg.input) {
      cfg.input.value = '';
      cfg.input.focus();
    }
  }

  function removeMultiSelection(field, key) {
    var selections = cmpFilterSelections[field] || [];
    cmpFilterSelections[field] = selections.filter(function (entry) { return entry.key !== key; });
    renderMultiSelectChips(field);
  }

  function clearMultiSelect(field) {
    cmpFilterSelections[field] = [];
    renderMultiSelectChips(field);
    closeMultiSelectDropdown(field);
  }

  function initMultiSelect(field) {
    var cfg = cmpMultiSelectFields[field];
    if (!cfg || !cfg.input || !cfg.dropdown || !cfg.chips) return;
    renderMultiSelectChips(field);
    cfg.input.addEventListener('focus', function () {
      if (cmpActiveDropdownField !== field) {
        openMultiSelectDropdown(field, cfg.input.value);
      }
    });
    cfg.input.addEventListener('input', function () {
      openMultiSelectDropdown(field, cfg.input.value);
    });
    cfg.input.addEventListener('click', function () {
      var isOpen = cfg.dropdown.classList.contains('is-open');
      if (cmpActiveDropdownField === field && isOpen) {
        closeMultiSelectDropdown(field);
      } else {
        openMultiSelectDropdown(field, cfg.input.value);
      }
    });
    cfg.input.addEventListener('keydown', function (evt) {
      if (evt.key === 'Enter') {
        evt.preventDefault();
        var options = cfg.dropdown.querySelectorAll('.cmp-multiselect-option');
        if (options.length) {
          selectMultiOption(field, options[0].getAttribute('data-key'));
        }
      } else if (evt.key === 'Escape') {
        closeMultiSelectDropdown(field);
      } else if (evt.key === 'Backspace' && !cfg.input.value) {
        var selections = cmpFilterSelections[field] || [];
        if (selections.length) {
          var last = selections[selections.length - 1];
          removeMultiSelection(field, last.key);
        }
      }
    });
    cfg.dropdown.addEventListener('mousedown', function (evt) {
      var target = evt.target;
      var optionEl = target && target.closest ? target.closest('.cmp-multiselect-option') : null;
      if (!optionEl) return;
      evt.preventDefault();
      selectMultiOption(field, optionEl.getAttribute('data-key'));
    });
    cfg.chips.addEventListener('click', function (evt) {
      var target = evt.target;
      if (!target || !target.getAttribute) return;
      if (target.classList.contains('cmp-chip-remove')) {
        var key = target.getAttribute('data-remove');
        removeMultiSelection(field, key);
      }
    });
  }

  function filterContactsByRadius(list, centerZip, radiusMiles) {
    if (!Array.isArray(list) || !list.length) return Promise.resolve([]);
    var safeZip = sanitizeZip(centerZip);
    if (!safeZip || !radiusMiles) return Promise.resolve(list);
    radiusMiles = Math.max(0, Number(radiusMiles) || 0);
    if (!radiusMiles) return Promise.resolve(list);
    return lookupZipCoords(safeZip).then(function (center) {
      if (!center) {
        var err = new Error("CENTER_UNRESOLVED");
        err.code = "CENTER_UNRESOLVED";
        throw err;
      }
      return new Promise(function (resolve) {
        var results = [];
        var pending = 0;
        var finish = function () {
          if (!pending) resolve(results);
        };
        for (var i = 0; i < list.length; i++) {
          var contact = list[i];
          if (!contact || !contact.zipKey) continue;
          if (contact._coord) {
            var cachedDistance = computeDistanceMiles(center, contact._coord);
            if (!isNaN(cachedDistance) && cachedDistance <= radiusMiles) {
              contact.distanceMiles = cachedDistance;
              results.push(contact);
            }
            continue;
          }
          if (contact.original && contact.original._cmpCoord) {
            contact._coord = contact.original._cmpCoord;
            var storedDistance = computeDistanceMiles(center, contact._coord);
            if (!isNaN(storedDistance) && storedDistance <= radiusMiles) {
              contact.distanceMiles = storedDistance;
              results.push(contact);
            }
            continue;
          }
          pending += 1;
          (function (shapedRef) {
            lookupZipCoords(shapedRef.zipKey).then(function (coords) {
              if (coords) {
                shapedRef._coord = coords;
                if (shapedRef.original) shapedRef.original._cmpCoord = coords;
                var dist = computeDistanceMiles(center, coords);
                if (!isNaN(dist) && dist <= radiusMiles) {
                  shapedRef.distanceMiles = dist;
                  results.push(shapedRef);
                }
              }
            }).catch(function () { }).then(function () {
              pending -= 1;
              finish();
            });
          })(contact);
        }
        finish();
      });
    });
  }

  function setContactResultsMessage(message) {
    cmpContactResultsData = [];
    cmpContactDisplayedData = [];
    cmpContactDisplayedMap = {};
    cmpContactSelectedMap = {};
    cmpContactLastTotal = 0;
    cmpContactLastDisplayed = 0;
    if (cmpContactResultsContainer) {
      cmpContactResultsContainer.className = "cmp-contact-results empty";
      cmpContactResultsContainer.innerHTML = '<p class="cmp-contact-placeholder">' + escapeHtml(message || "") + '</p>';
    }
    if (cmpContactResultsCount) {
      cmpContactResultsCount.textContent = message || "";
    }
    if (cmpContactSelectAll) cmpContactSelectAll.disabled = true;
    if (cmpContactClearSelection) cmpContactClearSelection.disabled = true;
    if (cmpAddSelectedContactsBtn) cmpAddSelectedContactsBtn.disabled = true;
  }

  function renderContactResults(list) {
    if (!cmpContactResultsContainer) return;
    if (!current) {
      setContactResultsMessage("Select a campaign to search contacts.");
      return;
    }
    var slice = Array.isArray(list) ? list.slice() : [];
    cmpContactResultsData = slice;
    cmpContactLastTotal = slice.length;
    var MAX_RENDER = 500;
    cmpContactDisplayedData = slice.slice(0, MAX_RENDER);
    cmpContactLastDisplayed = cmpContactDisplayedData.length;
    cmpContactDisplayedMap = {};
    var hasDistance = cmpContactDisplayedData.some(function (item) {
      return item && item.distanceMiles != null && isFinite(item.distanceMiles);
    });
    if (!cmpContactDisplayedData.length) {
      setContactResultsMessage("No contacts match your filters.");
      return;
    }
    var rows = [];
    rows.push('<table class="cmp-contact-table"><thead><tr>');
    var headerCells = [
      '<th class="checkbox-col"><input type="checkbox" disabled></th>',
      '<th>Company</th>',
      '<th>Name</th>',
      '<th>Email</th>',
      '<th>City</th>',
      '<th>State</th>',
      '<th>Zip</th>'
    ];
    if (hasDistance) headerCells.push('<th class="distance-col">Distance (mi)</th>');
    headerCells.push('<th>Division</th>');
    headerCells.push('<th>Project</th>');
    rows.push(headerCells.join(''));
    rows.push('</tr></thead><tbody>');

    var buildCell = function (value, extraClass, titleValue) {
      var hasValue = value != null && value !== '';
      var textValue = hasValue ? String(value) : '';
      var display = hasValue ? escapeHtml(textValue) : '&mdash;';
      var rawTitle = titleValue != null ? String(titleValue) : (hasValue ? textValue : '');
      var titleAttr = rawTitle && rawTitle.trim() ? ' title="' + escapeHtml(rawTitle) + '"' : '';
      var classAttr = extraClass ? ' class="' + extraClass + '"' : '';
      return '<td' + classAttr + titleAttr + '>' + display + '</td>';
    };

    for (var i = 0; i < cmpContactDisplayedData.length; i++) {
      var contact = cmpContactDisplayedData[i];
      cmpContactDisplayedMap[contact.email] = contact;
      var checked = Object.prototype.hasOwnProperty.call(cmpContactSelectedMap, contact.email);
      rows.push('<tr>');
      rows.push('<td class="checkbox-col"><input type="checkbox" class="cmp-contact-checkbox" data-email="' + escapeHtml(contact.email) + '"' + (checked ? ' checked' : '') + '></td>');
      rows.push(buildCell(contact.company, '', contact.company));
      rows.push(buildCell(contact.name, '', contact.name));
      var emailLabel = contact.emailRaw || contact.email;
      rows.push(buildCell(emailLabel, '', emailLabel));
      rows.push(buildCell(contact.city, '', contact.city));
      rows.push(buildCell(contact.state, '', contact.state));
      rows.push(buildCell(contact.zip, '', contact.zip));
      if (hasDistance) {
        var distanceText = '';
        if (contact.distanceMiles != null && isFinite(contact.distanceMiles)) {
          distanceText = formatDistanceMiles(contact.distanceMiles);
        }
        rows.push(buildCell(distanceText, 'distance-col', distanceText));
      }
      rows.push(buildCell(contact.division, '', contact.division));
      rows.push(buildCell(contact.project, '', contact.project));
      rows.push('</tr>');
    }
    rows.push('</tbody></table>');
    cmpContactResultsContainer.innerHTML = rows.join("");
    var extra = "";
    if (cmpContactLastTotal > cmpContactLastDisplayed) {
      extra = "Showing first " + cmpContactLastDisplayed + " of " + cmpContactLastTotal + " contacts";
    }
    updateContactSelectionSummary(extra);
  }

  function updateContactSelectionSummary(extraMessage) {
    var selected = 0;
    for (var key in cmpContactSelectedMap) {
      if (Object.prototype.hasOwnProperty.call(cmpContactSelectedMap, key)) selected += 1;
    }
    if (cmpContactResultsCount) {
      if (cmpContactLastDisplayed) {
        var base = cmpContactLastDisplayed + " contact" + (cmpContactLastDisplayed !== 1 ? "s" : "") + " listed";
        if (cmpContactLastTotal > cmpContactLastDisplayed) {
          base += " (of " + cmpContactLastTotal + ")";
        }
        if (selected) {
          base += " · " + selected + " selected";
        }
        if (extraMessage) {
          base += " · " + extraMessage;
        }
        cmpContactResultsCount.textContent = base;
      } else {
        cmpContactResultsCount.textContent = extraMessage || "No contacts found";
      }
    }
    if (cmpContactSelectAll) cmpContactSelectAll.disabled = !cmpContactLastDisplayed;
    if (cmpContactClearSelection) cmpContactClearSelection.disabled = !selected;
    if (cmpAddSelectedContactsBtn) cmpAddSelectedContactsBtn.disabled = !selected;
  }

  function selectAllDisplayedContacts() {
    if (!cmpContactDisplayedData.length) return;
    for (var i = 0; i < cmpContactDisplayedData.length; i++) {
      var contact = cmpContactDisplayedData[i];
      cmpContactSelectedMap[contact.email] = contact;
    }
    if (cmpContactResultsContainer) {
      var nodes = cmpContactResultsContainer.querySelectorAll("input.cmp-contact-checkbox");
      Array.prototype.forEach.call(nodes, function (node) { node.checked = true; });
    }
    updateContactSelectionSummary();
  }

  function clearContactSelection() {
    cmpContactSelectedMap = {};
    if (cmpContactResultsContainer) {
      var nodes = cmpContactResultsContainer.querySelectorAll("input.cmp-contact-checkbox");
      Array.prototype.forEach.call(nodes, function (node) { node.checked = false; });
    }
    updateContactSelectionSummary();
  }

  function handleContactCheckboxChange(target) {
    if (!target || !target.classList || !target.classList.contains("cmp-contact-checkbox")) return;
    var email = normalizeText(target.getAttribute("data-email"));
    if (!email) return;
    var contact = cmpContactDisplayedMap[email];
    if (!contact) return;
    if (target.checked) {
      cmpContactSelectedMap[email] = contact;
    } else {
      delete cmpContactSelectedMap[email];
    }
    updateContactSelectionSummary();
  }

  function addSelectedContactsToCampaign() {
    if (!current) { alert("Select a campaign first."); return; }
    var emails = [];
    for (var key in cmpContactSelectedMap) {
      if (Object.prototype.hasOwnProperty.call(cmpContactSelectedMap, key)) emails.push(key);
    }
    if (!emails.length) { alert("Select at least one contact."); return; }
    send({ type: "AB_CMP_ADD_RECIPIENTS", id: current.id, emails: emails }).then(function (res) {
      if (res && res.success) {
        current = res.item;
        cmpContactSelectedMap = {};
        renderRecipients();
        updateStats();
        renderContactResults(cmpContactResultsData);
        alert("Added " + emails.length + " contact" + (emails.length !== 1 ? "s" : "") + " to the campaign.");
      } else {
        alert("Unable to add contacts. Please try again.");
      }
    });
  }

  function resetContactSearch() {
    if (cmpSearchZip) cmpSearchZip.value = "";
    if (cmpSearchRadius) cmpSearchRadius.value = "";
    clearMultiSelect('city');
    clearMultiSelect('state');
    clearMultiSelect('project');
    clearMultiSelect('division');
    setContactResultsMessage("Run a search to see contacts.");
  }

  function runContactSearch() {
    if (cmpCurrentMode !== "existing" || !current) {
      if (cmpCurrentMode === "existing") {
        setContactResultsMessage("Select a campaign to search contacts.");
      } else {
        setContactResultsMessage("Switch to an existing campaign to search contacts.");
      }
      return;
    }
    var zip = cmpSearchZip ? sanitizeZip(cmpSearchZip.value) : "";
    if (cmpSearchZip) cmpSearchZip.value = zip;
    var radiusVal = cmpSearchRadius ? Number(cmpSearchRadius.value) : 0;
    if (!isFinite(radiusVal)) radiusVal = 0;
    if (radiusVal < 0) radiusVal = Math.abs(radiusVal);
    var cityFilters = getSelectedFilters('city');
    var stateFilters = getSelectedFilters('state');
    var projectFilters = getSelectedFilters('project');
    var divisionFilters = getSelectedFilters('division');
    if (!zip && !cityFilters.length && !stateFilters.length && !projectFilters.length && !divisionFilters.length) {
      setContactResultsMessage("Enter at least one filter to search.");
      return;
    }
    if (zip && !radiusVal) {
      setContactResultsMessage("Enter a radius when searching by zip code.");
      return;
    }
    setContactResultsMessage("Searching contacts...");
    ensureContactCache().then(function (contacts) {
      var filtered = [];
      var seen = {};
      for (var i = 0; i < contacts.length; i++) {
        var shaped = shapeContact(contacts[i]);
        if (!shaped) continue;
        if (cityFilters.length) {
          var matchesCity = shaped.cityKey && cityFilters.some(function (filter) {
            var match = shaped.cityKey.indexOf(filter.value) !== -1;
            if (!match) return false;
            if (filter.stateAbbrev && shaped.stateAbbrev) {
              return shaped.stateAbbrev === filter.stateAbbrev;
            }
            return true;
          });
          if (!matchesCity) continue;
        }
        if (stateFilters.length) {
          var matchesState = shaped.stateKey && stateFilters.some(function (filter) {
            if (filter.abbr && shaped.stateAbbrev) {
              return shaped.stateAbbrev === filter.abbr;
            }
            return shaped.stateKey.indexOf(filter.value) !== -1;
          });
          if (!matchesState) continue;
        }
        if (projectFilters.length) {
          var matchesProject = projectFilters.some(function (filter) { return projectMatches(shaped, filter.value); });
          if (!matchesProject) continue;
        }
        if (divisionFilters.length) {
          var matchesDivision = shaped.divisionKey && divisionFilters.some(function (filter) { return shaped.divisionKey.indexOf(filter.value) !== -1; });
          if (!matchesDivision) continue;
        }
        if (seen[shaped.email]) continue;
        seen[shaped.email] = true;
        filtered.push(shaped);
      }
      if (!filtered.length) {
        setContactResultsMessage("No contacts match your filters.");
        return;
      }
      var finalize = function (list) {
        if (!list || !list.length) {
          setContactResultsMessage("No contacts match your filters.");
          return;
        }
        var sorted = list.slice().sort(function (a, b) {
          var da = (a && a.distanceMiles != null && isFinite(a.distanceMiles)) ? a.distanceMiles : Infinity;
          var db = (b && b.distanceMiles != null && isFinite(b.distanceMiles)) ? b.distanceMiles : Infinity;
          if (da !== db) return da - db;
          var companyCompare = String((a && a.company) || '').localeCompare(String((b && b.company) || ''));
          if (companyCompare) return companyCompare;
          var nameCompare = String((a && a.name) || '').localeCompare(String((b && b.name) || ''));
          if (nameCompare) return nameCompare;
          return String((a && a.email) || '').localeCompare(String((b && b.email) || ''));
        });
        cmpContactSelectedMap = {};
        renderContactResults(sorted);
      };
      if (zip && radiusVal) {
        filterContactsByRadius(filtered, zip, radiusVal).then(function (radiusMatches) {
          finalize(radiusMatches);
        }).catch(function (err) {
          if (err && err.code === "CENTER_UNRESOLVED") {
            setContactResultsMessage("Unable to locate the center zip code.");
          } else {
            setContactResultsMessage("Failed to apply radius filter.");
          }
        });
      } else {
        finalize(filtered);
      }
    }).catch(function () {
      setContactResultsMessage("Contact search failed. Please try again.");
    });
  }
  function setMode(mode) {
    cmpCurrentMode = mode;
    var isExisting = mode === "existing";
    cmpModeExistingEls.forEach(function (el) { show(el, isExisting); });
    cmpModeNewEls.forEach(function (el) { show(el, !isExisting); });
    show(cmpLoaded, isExisting && !!current);
    toggleCampaignRequirements(isExisting && !!current);
    if (!isExisting) {
      setEditorValue(cmpHtmlTplEditor, cmpHtmlTpl, getEditorValue(cmpHtmlTplEditor, cmpHtmlTpl));
    } else if (cmpHtmlTplEditEditor && cmpHtmlTplEdit) {
      setEditorValue(cmpHtmlTplEditEditor, cmpHtmlTplEdit, getEditorValue(cmpHtmlTplEditEditor, cmpHtmlTplEdit));
    }
  }

  function populateSelect() {
    if (!cmpSelect) return;
    cmpSelect.innerHTML = "";
    var opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "Select a campaign";
    cmpSelect.appendChild(opt);
    campaigns.forEach(function (c) {
      var o = document.createElement("option");
      var rec = (c.recipients || []).length;
      o.value = c.id;
      o.textContent = c.name + " (" + rec + ")";
      cmpSelect.appendChild(o);
    });
  }

  function getPendingEmails() {
    if (!current) return [];
    var responded = {};
    Object.keys(current.responses || {}).forEach(function (k) { responded[String(k).toLowerCase()] = true; });
    return (current.recipients || []).map(function (r) { return r.email; }).filter(function (e) {
      return e && !responded[String(e).toLowerCase()];
    });
  }

  function manualSendBlockedUntil() {
    if (!current || !current.nextSendAt || current.paused) return null;
    var delta = Number(current.nextSendAt) - Date.now();
    if (!isFinite(delta) || delta <= CMP_SCHEDULE_MIN_LEAD_MS) return null;
    return current.nextSendAt;
  }

  function updateManualSendNotice() {
    if (!sendSelectedBtn) return;
    var blockedUntil = manualSendBlockedUntil();
    if (blockedUntil) {
      sendSelectedBtn.disabled = true;
      if (cmpSendNotice) {
        cmpSendNotice.textContent = 'Next send scheduled for ' + fmtDate(blockedUntil) + '. Manual send is disabled until that time.';
        cmpSendNotice.classList.add('warning');
      }
    } else {
      sendSelectedBtn.disabled = false;
      if (cmpSendNotice) {
        cmpSendNotice.textContent = 'Select recipients to send immediately.';
        cmpSendNotice.classList.remove('warning');
      }
    }
  }

  function renderPendingModal(preset) {
    if (!pendingList) return;
    var items = Array.isArray(preset) ? preset : getPendingEmails();
    pendingList.innerHTML = "";
    if (!items.length) {
      pendingList.innerHTML = '<div style="padding:8px 0; color:#6b7280;">No non-responders. Nothing to send.</div>';
    } else {
      for (var i = 0; i < items.length; i++) {
        var email = items[i];
        var row = document.createElement("label");
        row.style.display = "flex";
        row.style.alignItems = "center";
        row.style.gap = "10px";
        row.style.borderBottom = "1px solid #f3f4f6";
        row.style.padding = "8px 0";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.className = "cmp-check";
        cb.checked = true;
        cb.value = email;
        var span = document.createElement("span");
        span.textContent = email;
        row.appendChild(cb);
        row.appendChild(span);
        pendingList.appendChild(row);
      }
    }
    updatePendingCount();
    updateManualSendNotice();
  }

  function updatePendingCount() {
    if (!pendingList || !pendingCount) return;
    var cbs = pendingList.querySelectorAll("input.cmp-check");
    var n = 0; Array.prototype.forEach.call(cbs, function (cb) { if (cb.checked) n++; });
    pendingCount.textContent = String(n) + " selected";
  }

  function renderRecipients() {
    if (!cmpRecipients) return;
    cmpRecipients.innerHTML = "";
    var list = (current && current.recipients) || [];
    list.forEach(function (r) {
      var row = document.createElement("div");
      row.style.display = "flex";
      row.style.justifyContent = "space-between";
      row.style.alignItems = "center";
      row.style.borderBottom = "1px solid #eee";
      row.style.padding = "6px 4px";
      var emailSpan = document.createElement("span");
      emailSpan.textContent = r.email;
      var btn = document.createElement("button");
      btn.className = "btn btn-secondary";
      btn.textContent = "Remove";
      btn.style.padding = "4px 8px";
      btn.onclick = function () {
        send({ type: "AB_CMP_REMOVE_RECIPIENTS", id: current.id, emails: [r.email] }).then(function (res) {
          if (res && res.success) {
            current = res.item;
            updateStats();
            renderRecipients();
          }
        });
      };
      row.appendChild(emailSpan);
      row.appendChild(btn);
      cmpRecipients.appendChild(row);
    });
  }

  function renderResponses() {
    var box = $("cmpResponses");
    if (!box) return;
    box.innerHTML = "";
    var resp = (current && current.responses) || {};
    var entries = [];
    for (var k in resp) if (Object.prototype.hasOwnProperty.call(resp, k)) entries.push([k, resp[k]]);
    if (!entries.length) { box.textContent = "No responses yet."; return; }
    entries.sort(function (a, b) {
      var ta = Date.parse(a[1].at || "") || 0;
      var tb = Date.parse(b[1].at || "") || 0;
      return tb - ta;
    });
    entries.forEach(function (kv) {
      var email = kv[0], info = kv[1];
      var row = document.createElement("div");
      row.style.borderBottom = "1px solid #eee";
      row.style.padding = "6px 4px";
      row.innerHTML = "<strong>" + email + "</strong> â€” " + (info.subject || "") + " <small>(" + (info.at || "") + ")</small><br>" + String(info.preview || "").replace(/[<>]/g, "");
      box.appendChild(row);
    });
  }

  function computeRecipientStats() {
    var stats = { pending: 0, sent: 0, responded: 0 };
    var log = (current && current.recipientLog) || {};
    for (var key in log) {
      if (!Object.prototype.hasOwnProperty.call(log, key)) continue;
      var entry = log[key];
      if (!entry) continue;
      var status = entry.status || 'pending';
      if (status === 'responded') {
        stats.responded += 1;
      } else if (status === 'sent') {
        stats.sent += 1;
      } else if (status === 'pending') {
        stats.pending += 1;
      }
    }
    return stats;
  }

  function renderRecipientSummary() {
    if (!cmpRecipientSummary) return;
    var log = (current && current.recipientLog) || {};
    var entries = [];
    for (var key in log) {
      if (Object.prototype.hasOwnProperty.call(log, key)) {
        entries.push(log[key]);
      }
    }
    if (!entries.length) {
      cmpRecipientSummary.innerHTML = '<div class="cmp-summary-empty">No delivery activity yet.</div>';
      return;
    }
    entries.sort(function (a, b) {
      var aTime = a.lastDeliveryAt || a.respondedAt || a.addedAt || 0;
      var bTime = b.lastDeliveryAt || b.respondedAt || b.addedAt || 0;
      return bTime - aTime;
    });
    var rows = ['<div class="cmp-summary-row cmp-summary-header"><div>Recipient</div><div>Status</div><div>Last Delivery</div><div>Response</div></div>'];
    for (var i = 0; i < entries.length && i < 200; i++) {
      var entry = entries[i];
      var deliveries = Array.isArray(entry.deliveries) ? entry.deliveries.length : 0;
      var deliveryText = entry.lastDeliveryAt ? fmtDate(entry.lastDeliveryAt) : '—';
      var responseText = entry.respondedAt ? fmtDate(entry.respondedAt) : '—';
      var status = entry.status || 'pending';
      rows.push(
        '<div class="cmp-summary-row">' +
          '<div><div class="cmp-summary-recipient">' + escapeHtml(entry.email || '') + '</div>' +
          '<div class="cmp-summary-meta">Deliveries: ' + deliveries + '</div></div>' +
          '<div><span class="cmp-status-chip status-' + status + '">' + status + '</span></div>' +
          '<div>' + deliveryText + '</div>' +
          '<div>' + responseText + '</div>' +
        '</div>'
      );
    }
    cmpRecipientSummary.innerHTML = rows.join('');
  }

  
  function collectUpcomingEvents(cmp) {
    var events = [];
    if (!cmp) return events;
    var now = Date.now();
    var baseSubject = (cmp.subjectTpl || '').trim() || 'Campaign Invite';

    function addEvent(type, isoString, subject) {
      if (!isoString) return;
      var ms = Date.parse(isoString);
      if (!isFinite(ms) || ms < now) return;
      events.push({ type: type || 'invite', time: ms, subject: subject || baseSubject });
    }

    var schedule = cmp.schedule || {};
    if (schedule.mode === 'custom') {
      (Array.isArray(schedule.customDates) ? schedule.customDates : []).forEach(function (iso) {
        addEvent('invite', iso, baseSubject);
      });
    } else if (cmp.nextSendAt) {
      var iso = new Date(Number(cmp.nextSendAt)).toISOString();
      addEvent(cmp.nextSendType || 'invite', iso, baseSubject);
    }

    FOLLOW_UP_KEYS.forEach(function (key) {
      var follow = cmp.followUps && cmp.followUps[key];
      if (!follow || !Array.isArray(follow.dates)) return;
      var subject = (follow.subjectTpl || '').trim() || baseSubject;
      follow.dates.forEach(function (iso) { addEvent(key, iso, subject); });
    });

    events.sort(function (a, b) { return a.time - b.time; });
    return events;
  }

  function getEventLabel(type) {
    if (type === 'invite') return 'Invite';
    if (type === 'rfi') return 'RFI Reminder';
    if (type === 'bids') return 'Bids Reminder';
    return (type || 'event').charAt(0).toUpperCase() + (type || 'event').slice(1);
  }

  function renderScheduleSummary() {
    if (!cmpScheduleSummary) return;
    if (!current) {
      cmpScheduleSummary.innerHTML = '<div class="cmp-summary-empty">No campaign selected.</div>';
      cmpScheduleSummary.classList.add('empty');
      return;
    }
    var events = collectUpcomingEvents(current);
    if (!events.length) {
      cmpScheduleSummary.innerHTML = '<div class="cmp-summary-empty">No upcoming emails scheduled.</div>';
      cmpScheduleSummary.classList.add('empty');
      return;
    }
    var rows = ['<table><thead><tr><th>Type</th><th>Subject</th><th>Date</th></tr></thead><tbody>'];
    events.forEach(function (event) {
      rows.push('<tr>' +
        '<td><span class="cmp-event-type type-' + event.type + '">' + getEventLabel(event.type) + '</span></td>' +
        '<td>' + escapeHtml(event.subject || '') + '</td>' +
        '<td>' + fmtDate(event.time) + '</td>' +
        '</tr>');
    });
    rows.push('</tbody></table>');
    cmpScheduleSummary.innerHTML = rows.join('');
    cmpScheduleSummary.classList.remove('empty');
  }

function updateStats() {
    var recipientStats = computeRecipientStats();
    if (kpiTotal) kpiTotal.textContent = String(recipientStats.sent);
    if (kpiResp) kpiResp.textContent = String(recipientStats.responded);
    if (kpiPending) kpiPending.textContent = String(recipientStats.pending);
    if (kpiOpen) kpiOpen.textContent = "N/A";
    if (kpiClick) kpiClick.textContent = "N/A";
    if (kpiLast) kpiLast.textContent = fmtDate(current && current.lastSentAt);
    if (kpiNext) kpiNext.textContent = fmtDate(current && current.nextSendAt);

    if (cmpSubjectTplEdit) cmpSubjectTplEdit.value = (current && current.subjectTpl) || "";
    setEditorValue(cmpHtmlTplEditEditor, cmpHtmlTplEdit, (current && current.htmlTpl) || "");
    if (cmpNextSend) cmpNextSend.value = current && current.nextSendAt ? fmtDate(current.nextSendAt) : "";

    if (current && current.schedule) {
      applyScheduleToUI(current.schedule, "edit");
    } else {
      resetScheduleUI("edit");
    }
    if (current && current.followUps) {
      applyFollowUpsToUI(current.followUps, "edit");
    } else {
      applyFollowUpsToUI({}, "edit");
    }

    show(cmpLoaded, !!current);
    toggleCampaignRequirements(cmpCurrentMode === "existing" && !!current);
    renderRecipientSummary();
    renderScheduleSummary();
    updateManualSendNotice();
  }

  function pollNow() {
    return send({ type: "AB_CAMPAIGN_POLL_NOW" }).catch(function () { return null; });
  }

  function loadCurrent(id) {
    if (!id) {
      current = null;
      show(cmpLoaded, false);
      toggleCampaignRequirements(false);
      setContactResultsMessage("Select a campaign to search contacts.");
      return Promise.resolve();
    }
    return send({ type: "AB_CMP_GET", id: id }).then(function (res) {
      if (res && res.success) {
        current = res.item;
        setContactResultsMessage("Run a search to see contacts.");
        updateStats();
        renderRecipients();
        renderResponses();
      } else {
        current = null;
        setContactResultsMessage("Select a campaign to search contacts.");
        updateStats();
        renderRecipients();
        renderResponses();
      }
    });
  }

  function refreshList() {
    return send({ type: "AB_CMP_LIST" }).then(function (res) {
      campaigns = (res && res.success && res.items) ? res.items : [];
      campaigns.sort(function (a, b) { return (b.lastSentAt || 0) - (a.lastSentAt || 0); });
      populateSelect();
    });
  }

  // ---------- Events ----------
  modeRadios.forEach(function (r) {
    r.addEventListener("change", function (e) { setMode(e.target.value); });
  });

  if (cmpSelect) cmpSelect.addEventListener("change", function () {
    pollNow().finally(function () { loadCurrent(cmpSelect.value); });
  });
  if (cmpRefreshBtn) cmpRefreshBtn.addEventListener("click", function () {
    pollNow().then(function () {
      refreshList();
    });
  });
  if (cmpDeleteBtn) cmpDeleteBtn.addEventListener("click", function () {
    if (!current) return;
    if (!confirm("Delete this campaign?")) return;
    send({ type: "AB_CMP_DELETE", id: current.id }).then(function (res) {
      if (res && res.success) {
        current = null;
        refreshList().then(function () { setMode("existing"); show(cmpLoaded, false); });
      }
    });
  });

  if (cmpCreateBtn) cmpCreateBtn.addEventListener("click", function () {
    var name = (cmpName && cmpName.value || "").trim();
    if (!name) { alert("Campaign name is required"); return; }
    var subjectTpl = cmpSubjectTpl && cmpSubjectTpl.value || "";
    var htmlTpl = getEditorValue(cmpHtmlTplEditor, cmpHtmlTpl);
    setEditorValue(cmpHtmlTplEditor, cmpHtmlTpl, htmlTpl);
    var schedule = readScheduleFromUI("new");
    if (!schedule) return;
    var followUps = readFollowUpsFromUI("new");
    var recipients = uniqueEmailsFromInput(cmpNewRecipients && cmpNewRecipients.value);
    send({
      type: "AB_CMP_CREATE",
      name: name,
      subjectTpl: subjectTpl,
      htmlTpl: htmlTpl,
      schedule: schedule,
      followUps: followUps,
      recipients: recipients.map(function (e) { return { email: e }; })
    }).then(function (res) {
      if (res && res.success) {
        var created = res.item || res.campaign || {};
        var id = created.id || res.id || name;
        var campaignName = created.name || name;
        
        console.log('[Campaign Create] Success! ID:', id, 'Name:', campaignName);
        
        // Store the created campaign ID for the modal button
        if (cmpSuccessModal) {
          cmpSuccessModal.dataset.createdId = id;
          console.log('[Campaign Create] Stored ID in modal dataset:', cmpSuccessModal.dataset.createdId);
        }
        
        // Update success message
        if (cmpSuccessMessage) {
          cmpSuccessMessage.textContent = "Campaign '" + campaignName + "' has been created successfully.";
        }
        
        // Clear the form fields
        if (cmpName) cmpName.value = "";
        if (cmpSubjectTpl) cmpSubjectTpl.value = "";
        setEditorValue(cmpHtmlTplEditor, cmpHtmlTpl, "");
        if (cmpNewRecipients) cmpNewRecipients.value = "";
        resetScheduleUI("new");  // This also resets follow-ups
        
        // Refresh the campaign list
        refreshList().then(function () {
          console.log('[Campaign Create] List refreshed, showing modal');
          // Show success modal
          if (cmpSuccessModal) {
            cmpSuccessModal.style.display = "flex";
          }
        });
      } else {
        alert("Failed to create campaign. Please try again.");
      }
    }).catch(function (err) {
      alert("Error creating campaign: " + (err && err.message || err));
    });
  });
  if (cmpAddBtn) cmpAddBtn.addEventListener("click", function () {
    if (!current) return;
    var emails = uniqueEmailsFromInput(cmpAddEmails && cmpAddEmails.value);
    if (!emails.length) return;
    send({ type: "AB_CMP_ADD_RECIPIENTS", id: current.id, emails: emails }).then(function (res) {
      if (res && res.success) {
        current = res.item;
        if (cmpAddEmails) cmpAddEmails.value = "";
        updateStats();
        renderRecipients();
      }
    });
  });

    if (cmpContactSearchForm) cmpContactSearchForm.addEventListener("submit", function (e) {
    if (e && typeof e.preventDefault === "function") e.preventDefault();
    runContactSearch();
  });
  if (cmpContactResetBtn) cmpContactResetBtn.addEventListener("click", function () {
    resetContactSearch();
  });
  if (cmpContactResultsContainer) cmpContactResultsContainer.addEventListener("change", function (e) {
    handleContactCheckboxChange(e && e.target);
  });
  if (cmpContactSelectAll) cmpContactSelectAll.addEventListener("click", function () {
    selectAllDisplayedContacts();
  });
  if (cmpContactClearSelection) cmpContactClearSelection.addEventListener("click", function () {
    clearContactSelection();
  });
  if (cmpAddSelectedContactsBtn) cmpAddSelectedContactsBtn.addEventListener("click", function () {
    addSelectedContactsToCampaign();
  });
  if (cmpScheduleModeRecurring) cmpScheduleModeRecurring.addEventListener("change", function () {
    if (cmpScheduleModeRecurring.checked) setScheduleMode("edit", "recurring");
  });
  if (cmpScheduleModeCustom) cmpScheduleModeCustom.addEventListener("change", function () {
    if (cmpScheduleModeCustom.checked) setScheduleMode("edit", "custom");
  });
  if (cmpScheduleModeNewRecurring) cmpScheduleModeNewRecurring.addEventListener("change", function () {
    if (cmpScheduleModeNewRecurring.checked) setScheduleMode("new", "recurring");
  });
  if (cmpScheduleModeNewCustom) cmpScheduleModeNewCustom.addEventListener("change", function () {
    if (cmpScheduleModeNewCustom.checked) setScheduleMode("new", "custom");
  });
  if (cmpScheduleCustomAddBtn) cmpScheduleCustomAddBtn.addEventListener("click", function () { addCustomDate("edit"); });
  if (cmpScheduleNewCustomAddBtn) cmpScheduleNewCustomAddBtn.addEventListener("click", function () { addCustomDate("new"); });
  if (cmpScheduleCustomList) cmpScheduleCustomList.addEventListener("click", function (e) { handleMainCustomListClick("edit", e); });
  if (cmpScheduleNewCustomList) cmpScheduleNewCustomList.addEventListener("click", function (e) { handleMainCustomListClick("new", e); });

FOLLOW_UP_KEYS.forEach(function (key) {
  var newCtrl = followUpControls.new[key];
  if (newCtrl) {
    if (newCtrl.enable) newCtrl.enable.addEventListener("change", function () { setFollowUpEnabled("new", key, !!newCtrl.enable.checked); });
    if (newCtrl.addBtn) newCtrl.addBtn.addEventListener("click", function () { addFollowUpDate("new", key); });
    if (newCtrl.list) newCtrl.list.addEventListener("click", function (e) { handleFollowUpListClick("new", key, e); });
  }
  var editCtrl = followUpControls.edit[key];
  if (editCtrl) {
    if (editCtrl.enable) editCtrl.enable.addEventListener("change", function () { setFollowUpEnabled("edit", key, !!editCtrl.enable.checked); });
    if (editCtrl.addBtn) editCtrl.addBtn.addEventListener("click", function () { addFollowUpDate("edit", key); });
    if (editCtrl.list) editCtrl.list.addEventListener("click", function (e) { handleFollowUpListClick("edit", key, e); });
  }
});

if (cmpSaveBtn) cmpSaveBtn.addEventListener("click", function () {
    if (!current) return;
    var schedule = readScheduleFromUI("edit");
    if (!schedule) return;
    var followUps = readFollowUpsFromUI("edit");
    var subjectTpl = (cmpSubjectTplEdit && cmpSubjectTplEdit.value) || "";
    var htmlTpl = getEditorValue(cmpHtmlTplEditEditor, cmpHtmlTplEdit);
    setEditorValue(cmpHtmlTplEditEditor, cmpHtmlTplEdit, htmlTpl);
    send({ type: "AB_CMP_UPDATE", id: current.id, schedule: schedule, followUps: followUps, subjectTpl: subjectTpl, htmlTpl: htmlTpl }).then(function (res) {
      if (res && res.success) { current = res.item; updateStats(); }
    });
  });
  if (cmpSendNowBtn) cmpSendNowBtn.addEventListener("click", function (e) {
    if (e && typeof e.preventDefault === "function") e.preventDefault();
    if (!current) { alert("Select a campaign to review."); return; }
    setEditorValue(cmpHtmlTplEditEditor, cmpHtmlTplEdit, getEditorValue(cmpHtmlTplEditEditor, cmpHtmlTplEdit));
    var pending = getPendingEmails();
    renderPendingModal(pending);
    if (modal) modal.style.display = "flex";
  });

  if (modalClose) modalClose.addEventListener("click", function () {
    if (modal) modal.style.display = "none";
  });

  if (cmpSuccessContinueBtn) cmpSuccessContinueBtn.addEventListener("click", function () {
    // Close the modal
    if (cmpSuccessModal) {
      cmpSuccessModal.style.display = "none";
    }
    
    // Get the created campaign ID
    var id = cmpSuccessModal && cmpSuccessModal.dataset.createdId;
    
    console.log('[Campaign Success] Continue clicked, campaign ID:', id);
    
    // Check the "Existing" radio button
    modeRadios.forEach(function (r) {
      if (r.value === "existing") {
        r.checked = true;
      }
    });
    
    // Switch to existing mode
    setMode("existing");
    
    // Scroll to top of the campaign tab
    var campaignTab = document.getElementById("campaign-tab");
    if (campaignTab) {
      console.log('[Campaign Success] Scrolling to top');
      campaignTab.scrollTop = 0;
    }
    
    // Wait a moment for the mode switch to complete, then select and load
    setTimeout(function() {
      // Select the newly created campaign in dropdown
      if (cmpSelect && id) {
        console.log('[Campaign Success] Selecting campaign in dropdown:', id);
        cmpSelect.value = id;
      }
      
      // Load the campaign
      if (id) {
        console.log('[Campaign Success] Loading campaign:', id);
        loadCurrent(id);
      }
    }, 100);
  });

  if (pendingList) pendingList.addEventListener("change", function (e) {
    if (e && e.target && e.target.classList && e.target.classList.contains("cmp-check")) updatePendingCount();
  });
  if (checkAllBtn) checkAllBtn.addEventListener("click", function () {
    var cbs = pendingList ? pendingList.querySelectorAll("input.cmp-check") : [];
    Array.prototype.forEach.call(cbs, function (cb) { cb.checked = true; });
    updatePendingCount();
  });
  if (uncheckAllBtn) uncheckAllBtn.addEventListener("click", function () {
    var cbs = pendingList ? pendingList.querySelectorAll("input.cmp-check") : [];
    Array.prototype.forEach.call(cbs, function (cb) { cb.checked = false; });
    updatePendingCount();
  });
  if (sendSelectedBtn) sendSelectedBtn.addEventListener("click", function () {
    if (!current) return;
    var cbs = pendingList ? pendingList.querySelectorAll("input.cmp-check") : [];
    var selected = [];
    Array.prototype.forEach.call(cbs, function (cb) { if (cb.checked) selected.push(cb.value); });
    if (!selected.length) { alert("No recipients selected."); return; }
    send({ type: "AB_CMP_SEND_SELECTED", id: current.id, emails: selected }).then(function (res) {
      if (res && res.success) {
        current = res.item;
        updateStats();
        if (modal) modal.style.display = "none";
        alert("Queued " + String(res.sentCount || 0) + " emails.");
      } else if (res && res.error === 'SEND_SCHEDULED_LATER') {
        var when = res.nextSendAt ? fmtDate(res.nextSendAt) : 'the next scheduled window';
        alert('Manual sends are disabled until ' + when + '.');
        updateManualSendNotice();
      } else {
        alert("Send failed");
      }
    });
  });

  if (cmpToggleBtn) cmpToggleBtn.addEventListener("click", function () {
    if (!current) return;
    send({ type: "AB_CMP_TOGGLE_PAUSE", id: current.id }).then(function (res) {
      if (res && res.success) {
        current = res.item;
        if (cmpToggleBtn) cmpToggleBtn.textContent = current.paused ? "Resume" : "Pause";
        updateStats();
      }
    });
  });

  
  resetScheduleUI("edit");
  resetScheduleUI("new");
  toggleCampaignRequirements(false);

  chrome.runtime.onMessage.addListener(function(m){ if(m && m.type==='AB_CMP_REFRESH'){ refreshList().then(function(){ if(current){ loadCurrent(current.id); } }); }});
// ---------- Init ----------
  // Only initialize campaign UI if the campaign tab is currently active
  function initCampaignUI() {
    setMode("existing");
    refreshList().then(function () {
      if (campaigns.length && cmpSelect) {
        cmpSelect.value = campaigns[0].id;
        loadCurrent(campaigns[0].id);
      }
    });
  }
  
  // Check if campaign tab is active before initializing
  var campaignTab = document.getElementById('campaign-tab');
  if (campaignTab && campaignTab.classList.contains('active')) {
    initCampaignUI();
  }
  
  // Listen for tab changes and initialize when campaign tab is activated
  document.addEventListener('click', function(e) {
    var target = e.target;
    // Check if a tab button was clicked
    while (target && !target.classList.contains('tab') && !target.classList.contains('more-item')) {
      target = target.parentElement;
    }
    if (target && (target.dataset.tab === 'campaign' || target.dataset.targetTab === 'campaign')) {
      // Campaign tab was clicked, initialize if not already done
      if (!campaigns || campaigns.length === 0) {
        initCampaignUI();
      }
    }
  });
})();





















