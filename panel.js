let autoLoadLastProjectRan = false;

function autoLoadLastProject() {
  try {
    if (autoLoadLastProjectRan) return;
    autoLoadLastProjectRan = true;
    if (typeof localStorage === 'undefined') return;
    const last = localStorage.getItem("ab_last_project");
    if (!last) return;
    const input = $("bidderProjectInput");
    if (input && !input.value) input.value = last;
  } catch (e) {
    console.warn('[Panel] autoLoadLastProject skipped:', e);
  }
}
// Panel.js - AutoBuilders Extension - COMPLETE FILE







const urlParams = new URLSearchParams(window.location.search);
const isPopupMode = urlParams.get('mode') === 'popup';
const isInIframe = window !== window.parent;

console.log('[Panel] Panel loaded, isInIframe:', isInIframe, 'isPopupMode:', isPopupMode);

window.addEventListener('message', (event) => {
  const data = event.data || {};
  if (data.type === 'AB_PANEL_DETACH_BUTTON') {
    const detachBtn = document.getElementById('detachPanelBtn');
    const closeBtn = document.getElementById('closePanel');
    const hide = data.show === false;
    if (detachBtn) {
      detachBtn.hidden = hide;
    }
    if (closeBtn) {
      closeBtn.hidden = hide;
    }
  }
});







window.displayContactsList ||= function(){};

document.addEventListener('DOMContentLoaded', () => {
  const detachBtn = document.getElementById('detachPanelBtn');
  const closeBtn = document.getElementById('closePanel');
  if (detachBtn) {
    detachBtn.hidden = isPopupMode;
  }
  if (closeBtn) {
    closeBtn.hidden = isPopupMode;
  }
});











const $ = (id) => document.getElementById(id);







// Close button posts a message to the parent (content.js)



document.addEventListener('click', (event) => {
  const detachBtn = event.target?.closest('#detachPanelBtn');
  if (detachBtn) {
    event.preventDefault();
    if (!isPopupMode) {
      detachBtn.hidden = true;
      try {
        window.parent.postMessage({ type: 'AB_PANEL_DETACH' }, '*');
      } catch (error) {
        console.error('[Panel] Failed to request detach', error);
      }
    }
    return;
  }

  if (event.target?.closest('#closePanel')) {
    event.preventDefault();
    if (isPopupMode) {
      window.close();
    } else {
      try {
        window.parent.postMessage({ type: 'AB_PANEL_CLOSE' }, '*');
      } catch (error) {
        console.error('[Panel] Failed to notify parent to close panel', error);
      }
    }
    return;
  }
});







let isConnected = false;



let allProjects = [];



let allContacts = [];



let selectedContacts = [];



let availableDivisions = [];



let availableStatesList = [];



let bulkStateScanActive = false;



let currentProject = null;



let emailData = {



  subject: '',



  sender: '',



  senderName: '',



  body: '',



  bodyHtml: '',



  attachments: [],



  sentAt: '',



  messageId: '',



  forceOverride: false,



  signature: {



    company: '',



    phone: '',



    name: ''



  }



};







let currentProjectContacts = [];







let currentProjectContext = {



  name: null,



  project: null,



  divisions: [],



  contactsByDivision: {},



  companyIndex: new Map()



};

const AUDIT_LOG_FETCH_LIMIT = 1000;
let auditLogEntries = [];







let zipCoordCache = {};







let initialDataLoaded = false;







const loadingOverlayRefs = {



  root: null,



  spinner: null,



  title: null,



  message: null,



  actionBtn: null



};







let loadingOverlayActionBound = false;







function getLoadingOverlayRefs() {



  if (loadingOverlayRefs.root) {



    if (loadingOverlayRefs.actionBtn && !loadingOverlayActionBound) {



      loadingOverlayRefs.actionBtn.addEventListener('click', handleLoadingOverlayAction);



      loadingOverlayActionBound = true;



    }



    return loadingOverlayRefs;



  }







  loadingOverlayRefs.root = $('appLoadingOverlay');



  loadingOverlayRefs.spinner = $('appLoadingSpinner');



  loadingOverlayRefs.title = $('appLoadingTitle');



  loadingOverlayRefs.message = $('appLoadingSubtext');



  loadingOverlayRefs.actionBtn = $('appLoadingAction');







  if (loadingOverlayRefs.actionBtn && !loadingOverlayActionBound) {



    loadingOverlayRefs.actionBtn.addEventListener('click', handleLoadingOverlayAction);



    loadingOverlayActionBound = true;



  }







  return loadingOverlayRefs;



}







function showLoadingOverlay(options = {}) {



  const refs = getLoadingOverlayRefs();



  if (!refs || !refs.root) return;







  const {



    title = 'Loading company data...',



    message = '',



    actionLabel = '',



    actionType = '',



    showSpinner = true



  } = options;







  refs.root.classList.remove('app-hidden');



  refs.root.setAttribute('aria-hidden', 'false');







  if (refs.title) refs.title.textContent = title;



  if (refs.message) {



    refs.message.textContent = message || '';



    refs.message.classList.toggle('app-hidden', !message);



  }



  if (refs.spinner) {



    refs.spinner.style.display = showSpinner ? '' : 'none';



  }



  if (refs.actionBtn) {



    if (actionLabel && actionType) {



      refs.actionBtn.textContent = actionLabel;



      refs.actionBtn.dataset.action = actionType;



      refs.actionBtn.classList.remove('app-hidden');



      refs.actionBtn.disabled = false;



    } else {



      refs.actionBtn.classList.add('app-hidden');



      refs.actionBtn.disabled = false;



      delete refs.actionBtn.dataset.action;



    }



  }



}







function hideLoadingOverlay() {



  const refs = getLoadingOverlayRefs();



  if (!refs || !refs.root) return;



  refs.root.classList.add('app-hidden');



  refs.root.setAttribute('aria-hidden', 'true');



  if (refs.actionBtn) {



    refs.actionBtn.disabled = false;



    delete refs.actionBtn.dataset.action;



  }



}







function setLoadingOverlayActionDisabled(disabled) {



  const refs = getLoadingOverlayRefs();



  if (!refs?.actionBtn || refs.actionBtn.classList.contains('app-hidden')) return;



  refs.actionBtn.disabled = !!disabled;



}







async function handleLoadingOverlayAction(event) {



  const action = event?.currentTarget?.dataset?.action;



  if (!action) return;







  setLoadingOverlayActionDisabled(true);



  try {



    if (action === 'retry') {



      await loadCachedDataOnStartup({ showOverlay: true, requireData: true, maxAttempts: 20, reason: 'retry' });



    } else if (action === 'connect') {



      await connectToSharePoint({ initiatedByOverlay: true });



    }



  } catch (error) {



    console.error('[Panel] Overlay action failed:', error);



  } finally {



    if (!initialDataLoaded) {



      setLoadingOverlayActionDisabled(false);



    }



  }



}







function delay(ms = 0) {



  return new Promise((resolve) => setTimeout(resolve, Math.max(0, ms)));



}







let matchedEmailContact = null;







const dropdownDataMap = new Map();







const DROPDOWN_DEFAULT_OPTIONS = {



  allowMultiple: true,



  minChars: 1,



  maxResults: 75,



  showAllOnFocus: true,



  filterMode: 'contains',



  onSelect: null



};







function escapeHtml(text) {



  if (!text) return '';



  const div = document.createElement('div');



  div.textContent = text;



  return div.innerHTML;



}







function sanitizeDataArray(dataArray = []) {



  const seen = new Set();



  const result = [];







  dataArray.forEach(item => {



    if (item === undefined || item === null) return;



    const value = String(item).trim();



    if (!value) return;



    const key = value.toLowerCase();



    if (seen.has(key)) return;



    seen.add(key);



    result.push(value);



  });







  return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));



}







function uniquePreserveOrder(dataArray = []) {



  const seen = new Set();



  const result = [];



  (dataArray || []).forEach(item => {



    if (item === undefined || item === null) return;



    const value = String(item).trim();



    if (!value) return;



    const key = normalizeText(value) || value.toLowerCase();



    if (seen.has(key)) return;



    seen.add(key);



    result.push(value);



  });



  return result;



}







function normalizeDivisionKey(value) {



  if (!value) return '';



  return String(value).toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, '');



}







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







const FALLBACK_DIVISION_KEY_INDEX = new Map();



FALLBACK_DIVISION_ORDER.forEach((name, index) => {



  const key = normalizeDivisionKey(name);



  if (!FALLBACK_DIVISION_KEY_INDEX.has(key)) {



    FALLBACK_DIVISION_KEY_INDEX.set(key, index);



  }



});







function getFallbackDivisionIndex(name) {



  if (!name) return -1;



  const key = normalizeDivisionKey(name);



  return FALLBACK_DIVISION_KEY_INDEX.has(key) ? FALLBACK_DIVISION_KEY_INDEX.get(key) : -1;



}







function extractDivisionCodeValue(label) {



  if (!label) return null;



  const match = String(label).trim().match(/^(\d{2})[-\s]?(\d{2,3})/);



  if (!match) return null;



  return (parseInt(match[1], 10) * 1000) + parseInt(match[2], 10);



}







function labelProjectCodeFirst(name) {



  if (!name) return '';



  const str = String(name).trim();



  if (!str) return '';



  const match = str.match(/(?:^|[\s(])(\d{2})[-\s]?(\d{2,3})(?:\))?\s*$/);



  if (!match) return str;



  const code = `${match[1]}-${match[2]}`;



  const pattern = new RegExp(`(?:^|[\\s(])${match[1]}[-\\s]?${match[2]}(?:\\))?\\s*$`, 'i');



  const cleaned = str.replace(pattern, '').trim();



  return cleaned ? `${code} ${cleaned}` : code;



}



function extractProjectCode(value) {



  if (!value) return null;



  const match = String(value).match(/(\d{2})[-\s]?(\d{2,3})/);



  if (!match) return null;



  return `${match[1]}-${match[2]}`;



}











function normalizeProjectCode(value) {



  if (!value) return null;



  const text = String(value).trim();



  const match = text.match(/^(\d{2})[-\s]?(\d{2,3})$/);



  if (!match) return null;



  return `${match[1]}-${match[2]}`;



}







function parseProjectCodeForSort(name) {



  if (!name) return null;



  const match = String(name).match(/(\d{2})-(\d{2,3})/);



  if (!match) return null;



  return (parseInt(match[1], 10) * 1000) + parseInt(match[2], 10);



}







function formatProjectLabel(project, baseName) {



  const raw = String(baseName || '').trim();



  let normalizedCode = normalizeProjectCode(project?.projectNumber);



  if (!normalizedCode) {



    const codeCandidates = [



      project?.projectNumber,



      project?.displayName,



      project?.name,



      project?.projectName,



      project?.projectFolderName,



      project?.fileName,



      raw



    ];



    for (const candidate of codeCandidates) {



      const extracted = extractProjectCode(candidate);



      if (extracted) {



        normalizedCode = extracted;



        break;



      }



    }



  }







  let label = raw;







  if (normalizedCode) {



    const lower = raw.toLowerCase();



    if (!lower.includes(normalizedCode.toLowerCase())) {



      label = raw ? `${normalizedCode} ${raw}` : normalizedCode;



    }



  }







  if (!label) {



    return normalizedCode || '';



  }







  const normalizedLabel = labelProjectCodeFirst(label);



  if (normalizedCode && !normalizedLabel.startsWith(normalizedCode)) {



    return `${normalizedCode} ${normalizedLabel}`.trim();



  }



  return normalizedLabel;



}







function getProjectDisplayName(project = null) {



  if (!project) return '';



  if (typeof project === 'string') {



    return labelProjectCodeFirst(project);



  }



  const candidates = [



    project.displayName,



    project.name,



    project.projectName,



    project.projectFolderName,



    project.fileName



  ];



  let baseName = '';



  for (const candidate of candidates) {



    if (!candidate) continue;



    const value = String(candidate).trim();



    if (!value) continue;



    if (!baseName) baseName = value;



    if (extractProjectCode(value)) {



      baseName = value;



      break;



    }



  }



  const label = formatProjectLabel(project, baseName);



  if (label) return label;



  return baseName || '';



}











const LEGACY_PROJECT_NAME_FALLBACKS = new Set(
  [
    '64-17 Tire Kingdom 6417 Venice',
    '60-44 Tire Kingdom 6044 Cape Coral',
    '57-21 21-042 57-21 Aldi Miramar Store #57',
    '40-21 21-025 40-21 PBC Fire Station #40',
    '30-19 Grieco Delray Chevrolet Estimating Workbook 1-30-19 Buy-out.xlsx',
    '28-20 Rutledge Center Austin Estimating Workbook 1-28-20.xlsx',
    '26-22 Tesla Center Westbury Workbook 8-26-22.xlsx',
    '26-21 Vandergriff Body Shop Budget 1-26-21.xlsx'
  ]
    .map(name => normalizeText(name))
    .filter(Boolean)
);

function compareProjectNamesDesc(a, b) {



  if (!a && !b) return 0;

  if (!a) return 1;

  if (!b) return -1;



  const normalizedA = normalizeText(a);

  const normalizedB = normalizeText(b);

  const legacyA = LEGACY_PROJECT_NAME_FALLBACKS.has(normalizedA);

  const legacyB = LEGACY_PROJECT_NAME_FALLBACKS.has(normalizedB);



  if (legacyA && !legacyB) return 1;

  if (!legacyA && legacyB) return -1;



  const codeA = parseProjectCodeForSort(a);

  const codeB = parseProjectCodeForSort(b);



  if (codeA !== null && codeB !== null) {

    if (codeA !== codeB) return codeB - codeA;

  } else if (codeA !== null) {

    return -1;

  } else if (codeB !== null) {

    return 1;

  }



  return b.localeCompare(a, undefined, { sensitivity: 'base' });

}



function normalizeProjectNameList(list = []) {



  const map = new Map();



  list.forEach(name => {



    const value = String(name || '').trim();



    if (!value) return;



    const formatted = labelProjectCodeFirst(value);



    if (isBlockedProjectName(formatted) || isExcludedProjectLabel(formatted)) return;



    const key = normalizeText(formatted) || formatted.toLowerCase();



    if (!map.has(key)) {



      map.set(key, formatted);



    }



  });



  const unique = Array.from(map.values());



  unique.sort(compareProjectNamesDesc);



  return unique;



}







function compareDivisionEntries(a, b) {



  if (!a && !b) return 0;



  if (!a) return 1;



  if (!b) return -1;



  const orderA = typeof a.codeValue === 'number' ? a.codeValue : (typeof a.order === 'number' ? a.order : Number.MAX_SAFE_INTEGER);



  const orderB = typeof b.codeValue === 'number' ? b.codeValue : (typeof b.order === 'number' ? b.order : Number.MAX_SAFE_INTEGER);



  if (orderA !== orderB) return orderA - orderB;



  const nameA = (a.displayName || a.name || a.sheetName || a.key || '').toString();



  const nameB = (b.displayName || b.name || b.sheetName || b.key || '').toString();



  return nameA.localeCompare(nameB, undefined, { sensitivity: 'base' });



}
function isNotUsedDivisionEntry(division) {
  if (!division) return false;
  const label = division.displayName || division.name || division.sheetName || division.key || '';
  const normalized = normalizeText(label);
  if (normalized.startsWith('not used')) return true;
  const keyNormalized = normalizeText(division.key || division.sheetName || '');
  return keyNormalized.startsWith('not used');
}

function partitionDivisionsByUsage(divisions = []) {
  const active = [];
  const notUsed = [];
  (divisions || []).forEach(division => {
    if (!division) return;
    if (isNotUsedDivisionEntry(division)) {
      notUsed.push(division);
    } else {
      active.push(division);
    }
  });
  return { active, notUsed };
}









function buildProjectNameList(projects = []) {



  const map = new Map();



  (projects || []).forEach(project => {



    if (hasBlockedProjectName(project) || hasExcludedProjectCode(project)) return;



    const displayName = getProjectDisplayName(project);



    if (!displayName || isBlockedProjectName(displayName) || isExcludedProjectLabel(displayName)) return;



    const key = normalizeText(displayName) || displayName.toLowerCase();



    if (!map.has(key)) {



      map.set(key, displayName);



    }



  });



  const unique = Array.from(map.values());



  unique.sort(compareProjectNamesDesc);



  return unique;



}







function escapeRegex(value) {



  return value.replace(/[\^$.*+?()[\]{}|/-]/g, '\\$&');



}







async function fetchCachedData({ timeout = 5000, retries = 2 } = {}) {



  let attempt = 0;



  let lastError = null;







  while (attempt <= retries) {



    try {



      return await sendWithTimeout({ type: 'AB_GET_CACHED_DATA' }, timeout);



    } catch (error) {



      lastError = error;



      const message = (error && error.message) || String(error || '');



      if (!/timeout/i.test(message) || attempt === retries) {



        throw error;



      }



      const nextTimeout = Math.min(timeout * 2, 30000);



      await new Promise(resolve => setTimeout(resolve, 1000 * (attempt + 1)));



      timeout = nextTimeout;



      attempt += 1;



    }



  }







  throw lastError || new Error('Failed to fetch cached data');



}







function highlightMatch(text, query) {



  if (!query) return escapeHtml(text);



  const regex = new RegExp(`(${escapeRegex(query)})`, 'ig');



  return text.split(regex).map(part => {



    if (!part) return '';



    if (part.toLowerCase() === query.toLowerCase()) {



      return `<span class="filter-highlight">${escapeHtml(part)}</span>`;



    }



    return escapeHtml(part);



  }).join('');



}







function buildCompanyDivisionIndex(contactsByDivision = {}) {



  const index = new Map();







  for (const [divisionKey, contacts] of Object.entries(contactsByDivision || {})) {



    if (!Array.isArray(contacts)) continue;







    contacts.forEach(contact => {



      const companyKey = contact.companyNormalized || normalizeText(contact.company);



      if (!companyKey) return;







      if (!index.has(companyKey)) {



        index.set(companyKey, {



          company: contact.company || '',



          divisions: new Set(),



          emails: new Set(),



          phones: new Set(),



          entries: []



        });



      }







      const bucket = index.get(companyKey);



      const divisionName = contact.division || contact.sheetName || divisionKey;



      if (divisionName) bucket.divisions.add(divisionName);



      if (contact.email) bucket.emails.add(contact.email.toLowerCase());



      const contactPhoneDigits = contact.phoneDigits || normalizePhone(contact.phone);



      if (contactPhoneDigits) bucket.phones.add(contactPhoneDigits);



      bucket.entries.push({ divisionKey, contact });



    });



  }







  return index;



}











function flattenProjectContacts(contactsByDivision = {}, project = null) {



  const list = [];



  for (const [divisionKey, contacts] of Object.entries(contactsByDivision || {})) {



    if (!Array.isArray(contacts)) continue;



    contacts.forEach(contact => {



      const baseDivisionName = contact.division || contact.sheetName || divisionKey;



      const mappedDivision = getDivisionDisplayLabel({



        displayName: baseDivisionName,



        name: baseDivisionName,



        sheetName: contact.sheetName || baseDivisionName,



        key: contact.divisionKey || divisionKey



      }, divisionKey, project);



      const divisionName = mappedDivision || baseDivisionName;



      const normalizedName = contact.name || contact.contactName || '';



      const companyNormalized = contact.companyNormalized || normalizeText(contact.company);



      const phoneDigits = contact.phoneDigits || normalizePhone(contact.phone || contact.phoneDigits);







      list.push({



        ...contact,



        divisionKey: contact.divisionKey || divisionKey,



        division: divisionName,



        name: contact.name || normalizedName,



        contactName: contact.contactName || normalizedName,



        companyNormalized,



        phoneDigits



      });



    });



  }



  return list;



}











function resolveSiteKeyFromProject(project) {



  if (!project || typeof project !== 'object') return null;



  const hints = [



    project.site,



    project.siteUrl,



    project.siteURL,



    project.sourceSite,



    project.folderSite,



    project.company,



    project.client



  ];



  for (const hint of hints) {



    if (typeof hint !== 'string') continue;



    const lower = hint.toLowerCase();



    if (lower.includes('quest')) return 'Quest';



    if (lower.includes('autobuilder')) return 'AutoBuilders';



  }



  return null;



}







function getCurrentSiteKey(projectOverride = null) {



  const overrideKey = resolveSiteKeyFromProject(projectOverride);



  if (overrideKey) return overrideKey;



  const contextKey = resolveSiteKeyFromProject(currentProjectContext?.project || null);



  if (contextKey) return contextKey;



  return 'AutoBuilders';



}











function getDivisionDisplayLabel(division, fallbackKey, projectOverride = null) {



  const candidates = [];



  if (division && typeof division === 'object') {



    if (division.displayName) candidates.push(division.displayName);



    if (division.name) candidates.push(division.name);



    if (division.sheetName) candidates.push(division.sheetName);



    if (division.key) candidates.push(division.key);



  } else if (typeof division === 'string') {



    candidates.push(division);



  }



  if (fallbackKey) candidates.push(fallbackKey);







  const siteKey = getCurrentSiteKey(projectOverride);



  const seen = new Set();



  const unique = [];



  for (const candidate of candidates) {



    if (!candidate) continue;



    const text = String(candidate).trim();



    if (!text) continue;



    const key = text.toLowerCase();



    if (seen.has(key)) continue;



    seen.add(key);



    unique.push(text);



  }







  if (typeof resolveDivisionFolder === 'function') {



    for (const candidate of unique) {



      try {



        const mapped = resolveDivisionFolder(candidate, siteKey);



        if (mapped) return mapped;



      } catch (error) {



        console.warn('[Panel] resolveDivisionFolder failed for', candidate, error);



      }



    }



  }







  return unique[0] || fallbackKey || '';



}







function normalizeContactsByDivision(contactsByDivision = {}, project = null) {



  const normalized = {};







  Object.entries(contactsByDivision || {}).forEach(([divisionKey, contacts]) => {



    if (!Array.isArray(contacts)) return;



    normalized[divisionKey] = contacts.map(contact => {



      const baseDivisionName = contact.division || contact.sheetName || divisionKey;



      const mappedDivision = getDivisionDisplayLabel({



        displayName: baseDivisionName,



        name: baseDivisionName,



        sheetName: contact.sheetName || baseDivisionName,



        key: contact.divisionKey || divisionKey



      }, divisionKey, project);



      const divisionName = mappedDivision || baseDivisionName;



      const companyNormalized = contact.companyNormalized || normalizeText(contact.company);



      const phoneDigits = contact.phoneDigits || normalizePhone(contact.phone || contact.phoneDigits);



      return {



        ...contact,



        division: divisionName,



        divisionKey: contact.divisionKey || divisionKey,



        companyNormalized,



        phoneDigits



      };



    });



  });







  const existingDivisionKeys = new Set(Object.keys(normalized || {}).map(normalizeDivisionKey));



  FALLBACK_DIVISION_ORDER.forEach(fallbackName => {



    const fallbackKey = normalizeDivisionKey(fallbackName);



    if (!existingDivisionKeys.has(fallbackKey)) {



      normalized[fallbackName] = [];



      existingDivisionKeys.add(fallbackKey);



    }



  });



  return normalized;



}







function applyProjectResponse(projectName, response) {



  const trimmedName = (projectName || '').trim();



  const contactsByDivisionRaw = response?.contactsByDivision || {};



  const contactsByDivision = normalizeContactsByDivision(contactsByDivisionRaw, response?.project);



  const divisions = normalizeDivisionList(response?.divisions || [], response?.project);







  const projectDisplayName = getProjectDisplayName(response?.project);



  const effectiveName = projectDisplayName || trimmedName || currentProject || '';



    currentProject = effectiveName;
  try { if (typeof localStorage !== "undefined") localStorage.setItem("ab_last_project", currentProject); } catch (_) {}







  const flattenedContacts = Array.isArray(response?.contacts) && response.contacts.length



    ? response.contacts.map(contact => {



        const divisionKey = contact.divisionKey || contact.sheetName || contact.division || '';



        const baseDivision = contact.division || contact.sheetName || divisionKey;



        const mappedDivision = getDivisionDisplayLabel({



          displayName: baseDivision,



          name: baseDivision,



          sheetName: contact.sheetName || baseDivision,



          key: divisionKey



        }, divisionKey, response?.project);



        return {



          ...contact,



          division: mappedDivision || baseDivision,



          divisionKey,



          companyNormalized: contact.companyNormalized || normalizeText(contact.company),



          phoneDigits: contact.phoneDigits || normalizePhone(contact.phone || contact.phoneDigits)



        };



      })



    : flattenProjectContacts(contactsByDivision, response?.project);







  currentProjectContacts = flattenedContacts;



  const sortedDivisions = Array.isArray(divisions) ? [...divisions].sort(compareDivisionEntries) : [];



  const { active: activeDivisions, notUsed: notUsedDivisions } = partitionDivisionsByUsage(sortedDivisions);



  const orderedDivisions = [...activeDivisions, ...notUsedDivisions];



  availableDivisions = orderedDivisions;







  let resolvedProject = response?.project ? { ...response.project } : {};



  if (effectiveName) {



    resolvedProject.name = response?.project?.name || effectiveName;



    resolvedProject.displayName = getProjectDisplayName(resolvedProject) || effectiveName;



  } else if (resolvedProject.name) {



    resolvedProject.displayName = getProjectDisplayName(resolvedProject);



  }



  resolvedProject.contactsByDivision = contactsByDivision;







  if (effectiveName) {



    const normalizedKey = normalizeText(effectiveName);



    let existingIndex = allProjects.findIndex(project => normalizeText(project?.name || project?.projectName || '') === normalizedKey);







    if (existingIndex >= 0) {



      allProjects[existingIndex] = { ...allProjects[existingIndex], ...resolvedProject };



    } else {



      allProjects.push({ ...resolvedProject });



      existingIndex = allProjects.length - 1;



    }







    allProjects.sort((a, b) => compareProjectNamesDesc(getProjectDisplayName(a), getProjectDisplayName(b)));



    const sortedIndex = allProjects.findIndex(project => normalizeText(project?.name || project?.projectName || '') === normalizedKey);



    if (sortedIndex >= 0) {



      resolvedProject = allProjects[sortedIndex];



    }



  } else if (resolvedProject.name) {



    resolvedProject.displayName = resolvedProject.displayName || resolvedProject.name;



    const normalizedKey = normalizeText(resolvedProject.name);



    if (!allProjects.some(project => normalizeText(project?.name || project?.projectName || '') === normalizedKey)) {



      allProjects.push({ ...resolvedProject });



      allProjects.sort((a, b) => compareProjectNamesDesc(getProjectDisplayName(a), getProjectDisplayName(b)));



    }



  }







  allProjects = prepareProjectList(allProjects);







  currentProjectContext = {



    name: currentProject,



    project: resolvedProject,



    divisions: orderedDivisions,



    contactsByDivision,



    companyIndex: buildCompanyDivisionIndex(contactsByDivision)



  };







  refreshProjectDropdowns();



  renderAddBidderDivisionOptions();



  renderBidDivisionChecklist();



  updateDivisionDuplicateWarnings();







  const bidderProjectInput = $('bidderProjectInput');



  if (bidderProjectInput) bidderProjectInput.value = currentProject;







  const bidProjectInput = $('bidProject');



  if (bidProjectInput) {



    bidProjectInput.value = currentProject;



  }







  return { contactsByDivision, contacts: flattenedContacts };



}











function updateProjectSummary(projectName) {



  const loadingMsg = $('projectLoadingMsg');



  const selectedInfo = $('selectedProjectInfo');



  const contactsContainer = $('projectContactsContainer');



  const displayName = projectName || currentProject || '';







  if (loadingMsg) loadingMsg.style.display = 'none';



  if (selectedInfo) selectedInfo.style.display = 'block';



  if (contactsContainer) contactsContainer.style.display = 'block';







  const selectedNameEl = $('selectedProjectName');



  const contactCountEl = $('projectContactCount');



  const divisionCountEl = $('projectDivisionCount');







  if (selectedNameEl) selectedNameEl.textContent = displayName;



  if (contactCountEl) {



    contactCountEl.textContent = `${currentProjectContacts.length} contact${currentProjectContacts.length !== 1 ? 's' : ''}`;



  }







  const divisions = currentProjectContext?.divisions || [];



  if (divisionCountEl) {



    divisionCountEl.textContent = `${divisions.length} division${divisions.length !== 1 ? 's' : ''}`;



  }







  displayProjectContacts();



}







async function loadProject(projectName, options = {}) {



  const { forceRefresh = false, forceContactRefresh = false } = options || {};



  const trimmedName = (projectName || '').trim();







  if (!trimmedName) {



    showStatus('Enter a project name to load', 'error');



    return;



  }







  const normalizedInput = normalizeText(trimmedName);



  const matchedProject = allProjects.find(project => normalizeText(getProjectDisplayName(project)) === normalizedInput || normalizeText(project?.name) === normalizedInput);



  const requestName = matchedProject?.name || matchedProject?.projectFolderName || trimmedName;



  const displayName = getProjectDisplayName(matchedProject) || trimmedName;







  try {



    showStatus(`Loading project: ${displayName}`, 'info');



    const response = await sendWithTimeout({



      type: 'AB_LOAD_PROJECT_WORKBOOK',



      projectName: requestName,



      includeContacts: true,



      forceRefresh: !!forceRefresh,



      forceContactRefresh: !!forceContactRefresh



    }, 300000);







    if (!response || !response.success) {



      throw new Error(response?.error || 'Failed to load project');



    }







    applyProjectResponse(trimmedName, response);



    const effectiveName = currentProject || trimmedName;



    updateProjectSummary(effectiveName);



    showStatus(`Loaded project: ${effectiveName}`, 'success');



    return response;



  } catch (error) {



    console.error('[Panel] Error loading project:', error);



    showStatus(`Error loading project: ${error.message}`, 'error');



    throw error;



  }



}











function normalizeDivisionList(divisions = [], project = null) {



  return divisions



    .map(division => {



      if (!division) return null;







      if (typeof division === 'string') {



        const key = division.trim();



        if (!key) return null;



        const displayName = getDivisionDisplayLabel(division, key, project);



        return {



          key,



          name: displayName,



          displayName,



          sheetName: key,



          startsOnRow4: false,



          codeValue: null,



          order: null



        };



      }







      if (typeof division === 'object') {



        const key = division.key || division.sheetName || division.name;



        if (!key) return null;



        const displayName = getDivisionDisplayLabel(division, key, project);



        return {



          key,



          name: displayName,



          displayName,



          sheetName: division.sheetName || key,



          startsOnRow4: !!division.startsOnRow4,



          codeValue: typeof division.codeValue === 'number' ? division.codeValue : null,



          order: typeof division.order === 'number' ? division.order : null



        };



      }







      return null;



    })



    .filter(Boolean);



}











function buildCompanyLabelWithDivisions(value) {



  if (!value) return '';



  const base = escapeHtml(value);



  if (!currentProjectContext?.companyIndex) return base;



  const entry = currentProjectContext.companyIndex.get(normalizeText(value));



  if (!entry || entry.divisions.size === 0) return base;



  const divisions = Array.from(entry.divisions).join(', ');



  return `${base}<span class="company-division-badge">Already on: ${escapeHtml(divisions)}</span>`;



}







function getZipCenterCoords(zip) {



  if (!zip) return null;



  const normalized = zip.trim().slice(0, 5);



  return zipCoordCache[normalized] || null;



}







async function lookupZipCoordinates(zip) {
  if (!zip) return null;
  const normalized = String(zip).trim().slice(0, 5);
  if (!normalized) return null;
  if (zipCoordCache[normalized]) {
    return zipCoordCache[normalized];
  }
  try {
    const response = await sendWithTimeout({ type: 'AB_LOOKUP_ZIP', zip: normalized }, 20000);
    if (response && response.success && response.location) {
      zipCoordCache[normalized] = response.location;
      if (typeof window !== 'undefined') {
        if (!window.zipCoordCache || typeof window.zipCoordCache !== 'object') {
          window.zipCoordCache = {};
        }
        window.zipCoordCache[normalized] = response.location;
      }
      return response.location;
    }
  } catch (error) {
    console.error('[Panel] Zip lookup failed:', error);
  }
  return null;
}








function computeDistanceMiles(coordA, coordB) {



  if (!coordA || !coordB) return Infinity;



  const toRadians = (deg) => deg * (Math.PI / 180);



  const earthRadiusMiles = 3958.8;



  const deltaLat = toRadians(coordB.lat - coordA.lat);



  const deltaLon = toRadians(coordB.lon - coordA.lon);







  const a = Math.sin(deltaLat / 2) ** 2 +



    Math.cos(toRadians(coordA.lat)) * Math.cos(toRadians(coordB.lat)) *



    Math.sin(deltaLon / 2) ** 2;



  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));



  return earthRadiusMiles * c;



}







async function filterContactsByRadius(contacts, centerZip, radiusMiles) {



  if (!radiusMiles || !centerZip) return contacts;







  const centerCoords = await lookupZipCoordinates(centerZip);



  if (!centerCoords) {



    showStatus('Unable to locate center zip code for radius filter', 'warning');



    return contacts;



  }







  const results = [];



  for (const contact of contacts) {



    const contactZip = contact.zip || contact.postal || '';



    if (!contactZip) continue;







    const contactCoords = await lookupZipCoordinates(contactZip);



    if (!contactCoords) continue;







    const distance = computeDistanceMiles(centerCoords, contactCoords);



    if (!isNaN(distance) && distance <= radiusMiles) {



      results.push(contact);



    }



  }







  return results;



}







function formatConflictPrompt(conflictReport, header = '') {



  if (!conflictReport) {



    return header || 'Potential conflicts detected.';



  }







  const lines = [];



  if (conflictReport.existingDivisions && conflictReport.existingDivisions.length) {



    lines.push(`Already on: ${conflictReport.existingDivisions.join(', ')}`);



  }







  if (conflictReport.divisionResults) {



    Object.values(conflictReport.divisionResults).forEach(result => {



      if (!result || !result.matches) return;



      const reasons = result.matches.map(match => match.reasons.join(', ')).join('; ');



      lines.push(`${result.existingContact?.sheetName || result.divisionKey}: ${reasons}`);



    });



  }







  if (conflictReport.otherMatches && conflictReport.otherMatches.length) {



    const grouped = new Map();



    conflictReport.otherMatches.forEach(match => {



      const list = grouped.get(match.divisionName) || [];



      list.push(match.reasons.join(', '));



      grouped.set(match.divisionName, list);



    });







    grouped.forEach((reasons, divisionName) => {



      lines.push(`${divisionName} (other project divisions): ${reasons.join('; ')}`);



    });



  }







  if (!lines.length) {



    lines.push('Potential duplicate detected.');



  }







  const prefix = header || 'Potential duplicates found:';



  return `${prefix}\n${lines.join('\n')}`;



}







function normalizeText(value) {



  if (!value) return '';



  return String(value)



    .toLowerCase()



    .replace(/[^a-z0-9]+/g, ' ')



    .replace(/\s+/g, ' ')



    .trim();



}







const PROJECT_NAME_BLOCKLIST = new Set(



  ['29-30 2930 Okeechobee Blvd', '2930 Okeechobee Blvd']



    .map(name => normalizeText(name))



    .filter(Boolean)



);



const PROJECT_CODE_EXCLUDE_PREFIXES = ['29-'];







function isBlockedProjectName(name) {



  const normalized = normalizeText(name);



  return normalized ? PROJECT_NAME_BLOCKLIST.has(normalized) : false;



}







function hasBlockedProjectName(project) {



  if (!project) return false;



  if (typeof project === 'string') {



    return isBlockedProjectName(project);



  }



  if (typeof project === 'object') {



    const candidates = [



      project.displayName,



      project.name,



      project.projectName,



      project.projectFolderName,



      project.fileName



    ];



    return candidates.some(isBlockedProjectName);



  }



  return false;



}







function isExcludedProjectLabel(label) {



  if (!label) return false;



  const code = extractProjectCode(label);



  if (!code) return false;



  return PROJECT_CODE_EXCLUDE_PREFIXES.some(prefix => code.startsWith(prefix));



}







function hasExcludedProjectCode(project) {



  if (!project) return false;



  if (typeof project === 'string') {



    return isExcludedProjectLabel(project);



  }



  if (typeof project === 'object') {



    const candidates = [



      project.projectNumber,



      project.displayName,



      project.name,



      project.projectName,



      project.projectFolderName,



      project.fileName



    ];



    for (const candidate of candidates) {



      if (isExcludedProjectLabel(candidate)) {



        return true;



      }



    }



    const fallback = getProjectDisplayName(project);



    return isExcludedProjectLabel(fallback);



  }



  return false;



}







function filterBlockedProjects(projects = []) {



  return (projects || []).filter(project => !hasBlockedProjectName(project) && !hasExcludedProjectCode(project));



}







function prepareProjectList(projects = []) {



  const filtered = filterBlockedProjects(projects || []);



  const sorted = [...filtered].sort((a, b) => compareProjectNamesDesc(



    getProjectDisplayName(a),



    getProjectDisplayName(b)



  ));



  sorted.forEach(project => {



    if (project && typeof project === 'object') {



      const label = getProjectDisplayName(project);



      if (label && project.displayName !== label) {



        project.displayName = label;



      }



    }



  });



  return sorted;



}







function normalizePhone(value) {



  if (!value) return '';



  return String(value).replace(/[^0-9]/g, '');



}







function formatPhoneNumber(value) {



  const digits = normalizePhone(value);



  if (digits.length === 10) {



    return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;



  }



  if (digits.length === 11 && digits.startsWith('1')) {



    return `+1 (${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7)}`;



  }



  return value ? String(value).trim() : '';



}

function formatDateInputValue(date = new Date()) {
  try {
    if (typeof date === 'string') {
      const match = date.match(/^\d{4}-\d{2}-\d{2}/);
      if (match) {
        return match[0];
      }
    }
    const d = new Date(date);
    if (!Number.isFinite(d.getTime())) return '';
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  } catch (error) {
    console.warn('[Panel] Failed to format date input value:', error);
    return '';
  }
}


function formatDateTime(value) {
  if (!value) return '';
  try {
    const date = new Date(value);
    if (!Number.isFinite(date.getTime())) {
      return String(value);
    }
    return date.toLocaleString();
  } catch (error) {
    console.warn('[Panel] Failed to format date time:', error);
    return String(value);
  }
}








function inferCompanyFromEmail(email) {



  if (!email || !email.includes('@')) return '';



  const domain = email.split('@')[1] || '';



  const domainParts = domain.split('.');



  if (domainParts.length === 0) return '';



  const core = domainParts[0].replace(/[^a-z0-9]/gi, ' ').trim();



  if (!core) return '';



  const normalizedCore = core.toLowerCase();



  const genericDomains = new Set(['gmail', 'yahoo', 'outlook', 'hotmail', 'icloud', 'me', 'msn', 'live', 'aol', 'googlemail']);



  if (genericDomains.has(normalizedCore)) return '';



  return core.split(' ').map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');



}







function getUniqueContactValues(field) {



  return sanitizeDataArray(allContacts.map(contact => contact[field]).filter(Boolean));



}











function pickFirstValue(values = []) {



  for (const value of values) {



    if (value === undefined || value === null) continue;



    const trimmed = String(value).trim();



    if (!trimmed) continue;



    return trimmed;



  }



  return '';



}







function setInputValue(inputId, value, options = {}) {



  if (value === undefined || value === null) return;



  const trimmed = String(value).trim();



  if (!trimmed) return;



  const input = $(inputId);



  if (!input) return;



  if (options.onlyIfEmpty && input.value) return;



  input.value = trimmed;



  input.dataset.selectedValues = trimmed;



}







function applyCacheResponse(cacheResponse = {}) {
  const contacts = Array.isArray(cacheResponse.contacts) ? cacheResponse.contacts : [];
  const projects = Array.isArray(cacheResponse.projects) ? cacheResponse.projects : [];
  allContacts = contacts;
  allProjects = prepareProjectList(projects);
  if (cacheResponse.zipCache && typeof cacheResponse.zipCache === 'object') {
    zipCoordCache = cacheResponse.zipCache;
  }
  if (typeof window !== 'undefined') {
    window.allContacts = allContacts;
    window.zipCoordCache = zipCoordCache;
  }
  const hasData = contacts.length > 0 || projects.length > 0;
  if (hasData && !initialDataLoaded) {
    initialDataLoaded = true;
    const refs = getLoadingOverlayRefs();
    if (refs?.root && !refs.root.classList.contains('app-hidden')) {
      hideLoadingOverlay();
    }
  }
  console.log('[Panel] Loaded from cache:', allContacts.length, 'contacts,', allProjects.length, 'projects');

  try {
    const runDropdownRefresh = () => {
      populateFilterDropdowns();
      const projectInput = $('bidderProjectInput');
      if (emailData?.subject && projectInput && !projectInput.value) {
        autoDetectProject(emailData.subject);
      }
    };

    if (document?.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => {
        try {
          runDropdownRefresh();
        } catch (err) {
          console.warn('[Panel] Deferred dropdown refresh failed:', err);
        }
      }, { once: true });
    } else {
      runDropdownRefresh();
    }
  } catch (error) {
    console.warn('[Panel] Failed to refresh dropdowns after cache load:', error);
  }
}








async function loadCachedDataOnStartup(options = {}) {



  const {



    showOverlay = false,



    requireData = false,



    maxAttempts = requireData ? 20 : 1,



    retryDelay = 1500,



    reason = 'default'



  } = options || {};







  let attempt = 0;



  let lastError = null;



  let response = null;



  let hasData = false;







  if (showOverlay) {



    showLoadingOverlay({



      title: 'Loading company data...',



      message: requireData ? 'Please wait while we prepare your workspace.' : '',



      showSpinner: true



    });



  }







  while (attempt < maxAttempts) {



    try {



      response = await fetchCachedData({ timeout: Math.min(8000 + attempt * 2000, 20000), retries: 1 });



      if (response && response.success) {
        if (typeof response.staticProjectsCount === 'number') {
          console.log('[Panel] Cache snapshot', {
            staticProjects: response.staticProjectsCount,
            dynamicProjects: response.dynamicProjectsCount,
            staticContacts: response.staticContactsCount,
            dynamicContacts: response.dynamicContactsCount
          });
        }



        applyCacheResponse(response);



        hasData = (Array.isArray(response.contacts) && response.contacts.length > 0) ||



          (Array.isArray(response.projects) && response.projects.length > 0);







        if (hasData || !requireData) {



          if (showOverlay) {



            hideLoadingOverlay();



          }



          if (hasData) {



            initialDataLoaded = true;



          }



          return { success: true, hasData, response };



        }







        lastError = new Error('Company data not available yet');



      } else {



        lastError = new Error(response?.error || 'Failed to load cached data');



      }



    } catch (error) {



      lastError = error;



    }







    attempt += 1;



    if (attempt < maxAttempts) {



      if (showOverlay) {



        const progressMessage = 'Still working... (attempt ' + (attempt + 1) + '/' + maxAttempts + ')';



        showLoadingOverlay({



          title: 'Loading company data...',



          message: progressMessage,



          showSpinner: true



        });



      }



      await delay(retryDelay * (attempt + 1));



    }



  }







  if (showOverlay) {



    const connected = isConnected;



    const message = lastError?.message || (connected



      ? 'Data is still syncing from SharePoint. Please try again shortly.'



      : 'Sign in to SharePoint to download the latest data.');



    showLoadingOverlay({



      title: connected ? 'Company data not ready yet' : 'Connect to download company data',



      message,



      actionLabel: connected ? 'Retry' : 'Connect',



      actionType: connected ? 'retry' : 'connect',



      showSpinner: false



    });



  }







  return { success: !!(response && response.success), hasData, error: lastError?.message || null };



}







async function checkConnection() {



  try {



    const response = await sendWithTimeout({ type: 'AB_GET_USER_PROFILE' }, 5000);



    if (response && response.success) {



      isConnected = true;



      const connectBtn = $('connectBtn');



      const connectionText = $('connectionText');



      const statusDot = $('statusDot');



      



      if (connectBtn) connectBtn.textContent = 'Disconnect';



      if (connectionText) connectionText.textContent = 'Connected';



      if (statusDot) statusDot.classList.add('connected');



      



      await loadStatesList();



      



      const cacheResponse = await fetchCachedData({ timeout: 5000, retries: 2 });



      if (cacheResponse && cacheResponse.success) {



        applyCacheResponse(cacheResponse);



      }



    }



  } catch (error) {



    console.log('[Panel] Not connected');



    isConnected = false;



    const connectBtn = $('connectBtn');



    const connectionText = $('connectionText');



    const statusDot = $('statusDot');



    



    if (connectBtn) connectBtn.textContent = 'Connect';



    if (connectionText) connectionText.textContent = 'Disconnected';



    if (statusDot) statusDot.classList.remove('connected');



  }



}







function sendWithTimeout(message, timeout = 30000) {



  return new Promise((resolve, reject) => {



    let finished = false;



    let timer = null;







    const heartbeatOps = new Set(['AB_SCAN_STATE','AB_LOAD_PROJECT','AB_LOAD_PROJECT_WORKBOOK','AB_EXPORT_EXCEL','AB_SUBMIT_BID']);



    const progressTypes = new Set(['AB_SCAN_PROGRESS','AB_LOAD_PROGRESS','AB_WORKBOOK_PROGRESS','AB_EXPORT_PROGRESS']);







    let rtListener = null;



    let winListener = null;



    let iframeResponseHandler = null;







    const cleanup = (err, response) => {



      if (finished) return;



      finished = true;



      if (timer) {



        try { clearTimeout(timer); } catch (_) {}



        timer = null;



      }



      if (rtListener) {



        try { chrome.runtime.onMessage.removeListener(rtListener); } catch (_) {}



        rtListener = null;



      }



      if (winListener) {



        try { window.removeEventListener('message', winListener); } catch (_) {}



        winListener = null;



      }



      if (iframeResponseHandler) {



        try { window.removeEventListener('message', iframeResponseHandler); } catch (_) {}



        iframeResponseHandler = null;



      }



      if (err) {



        reject(err);



      } else {



        resolve(response);



      }



    };







    const bump = () => {



      try { clearTimeout(timer); } catch (_) {}



      timer = setTimeout(() => cleanup(new Error('Message timeout')), timeout);



    };







    bump();







    if (heartbeatOps.has(message && message.type)) {



      rtListener = (msg) => {



        if (msg && progressTypes.has(msg.type)) bump();



      };



      try { chrome.runtime.onMessage.addListener(rtListener); } catch (_) {}







      winListener = (evt) => {



        const data = evt && evt.data;



        if (data && progressTypes.has(data.type)) bump();



      };



      try { window.addEventListener('message', winListener); } catch (_) {}



    }







    try {



      if (isInIframe) {



        const messageId = `${Date.now()}_${Math.random()}`;



        iframeResponseHandler = (event) => {



          const data = event && event.data;



          if (data && data.type === 'AB_RESPONSE' && data.messageId === messageId) {



            cleanup(null, data.response);



          }



        };



        window.addEventListener('message', iframeResponseHandler);



        window.parent.postMessage({ type: 'AB_FORWARD_TO_BACKGROUND', message, messageId }, '*');



      } else {



        chrome.runtime.sendMessage(message, (response) => {



          if (chrome.runtime.lastError) {



            cleanup(chrome.runtime.lastError, null);



          } else {



            cleanup(null, response);



          }



        });



      }



    } catch (err) {



      cleanup(err, null);



    }



  });



}



chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {



  if (message.type === 'AB_SCAN_PROGRESS') {



    updateProgressUI(message.progress);



  }



});







function updateProgressUI(progress = {}) {



  if (!initialDataLoaded) {



    const statusText = progress.status || `Processed ${progress.current || 0}/${progress.total || 0}`;



    showLoadingOverlay({



      title: progress.baseline ? 'Scanning SharePoint...' : 'Updating company data...',



      message: statusText,



      showSpinner: true



    });



  }







  const container = $('scanProgressContainer');



  const progressBar = $('progressBar');



  const progressStatus = $('progressStatus');



  const fileCount = $('fileCount');



  const timeRemaining = $('timeRemaining');







  if (!container || !progressBar) return;







  container.style.display = 'block';







  const percent = Number.isFinite(progress.percentage) ? progress.percentage : 0;



  progressBar.style.width = percent + '%';



  progressBar.textContent = percent + '%';







  if (progressStatus) progressStatus.textContent = progress.status || 'Processing...';







  if (fileCount) fileCount.textContent = `Files: ${progress.current || 0}/${progress.total || 0}`;







  if (timeRemaining) {



    if (progress.status && progress.status.includes('Est.')) {



      const timeMatch = progress.status.match(/Est. (.+?) remaining/);



      if (timeMatch) {



        timeRemaining.textContent = 'Time remaining: ' + timeMatch[1];



      }



    } else if (progress.current && progress.total && progress.current === progress.total) {



      timeRemaining.textContent = 'Complete!';



      setTimeout(() => {



        container.style.display = 'none';



        resetProgressUI();



      }, 3000);



    }



  }



}







function handleCacheReadyMessage(message = {}) {



  const baseline = message?.baseline === true;



  const result = message?.result || null;



  const errorText = message?.error || result?.error || null;



  const uploadError = result?.cacheUpload && result.cacheUpload.success === false ? result.cacheUpload.error : null;







  const shouldRequireData = !initialDataLoaded;



  const attempts = shouldRequireData ? 20 : 3;







  loadCachedDataOnStartup({



    showOverlay: shouldRequireData,



    requireData: shouldRequireData,



    maxAttempts: attempts,



    reason: baseline ? 'baseline-complete' : 'incremental-complete'



  }).catch(error => {



    console.warn('[Panel] Cache refresh after scan failed:', error);



  });







  if (uploadError) {



    showStatus('SharePoint upload warning: ' + uploadError, 'warning');



  } else if (errorText) {



    showStatus('Automatic scan issue: ' + errorText, 'warning');



  } else if (!shouldRequireData) {



    showStatus('Company data refreshed from SharePoint.', 'success');



  }



}



function resetProgressUI() {



  const progressBar = $('progressBar');



  const progressStatus = $('progressStatus');



  const fileCount = $('fileCount');



  const timeRemaining = $('timeRemaining');



  



  if (progressBar) {



    progressBar.style.width = '0%';



    progressBar.textContent = '0%';



  }



  if (progressStatus) progressStatus.textContent = 'Initializing...';



  if (fileCount) fileCount.textContent = 'Files: 0/0';



  if (timeRemaining) timeRemaining.textContent = 'Time remaining: calculating...';



}







window.addEventListener('message', async (event) => {







  console.log('[Panel] Iframe bridge - RAW message received:', event.data);



  



  if (event.data?.type === 'AB_EMAIL_DATA_RESPONSE') {



    console.log('[Panel] Received email data:', event.data.emailData);



    



    if (event.data.emailData) {



      const email = event.data.emailData;



      console.log('[Panel] Auto-populating Add Bidder form with:', email);



      



      emailData.sender = email.from || '';



      emailData.senderName = email.fromName || '';



      emailData.subject = email.subject || '';



      emailData.body = email.body || '';



      emailData.bodyHtml = email.bodyHtml || '';



      emailData.attachments = Array.isArray(email.attachments) ? email.attachments : [];



      emailData.sentAt = email.sentAt || '';



      emailData.messageId = email.messageId || '';



      emailData.signature.company = email.signature?.company || '';



      emailData.signature.phone = email.signature?.phone || '';



      emailData.forceOverride = false;



      



      const addBidderTab = document.getElementById('add-bidder-tab');



      const isAddBidderActive = addBidderTab?.classList.contains('active');



      



      if (isAddBidderActive) {



        populateContactFromEmail();



        autofillBidContactFromEmail();



        



        if (email.subject && allProjects.length > 0) {



          autoDetectProject(email.subject);



        }



      } else {



        const addBidderTabButton = document.querySelector('[data-tab="add-bidder"]');



        if (addBidderTabButton) {



          // addBidderTabButton.click(); // DISABLED: Don't auto-switch tabs



        }



        



        setTimeout(() => {



          populateContactFromEmail();



          autofillBidContactFromEmail();



          



          if (email.subject && allProjects.length > 0) {



            autoDetectProject(email.subject);



          }



        }, 100);



      }



    }



  }



});







let requestAttempts = 0;



const maxAttempts = 3;







function requestEmailData() {



  requestAttempts++;



  console.log(`[Panel] Requesting email data (attempt ${requestAttempts}/${maxAttempts})`);



  window.parent?.postMessage?.({ type: 'AB_REQUEST_EMAIL_DATA' }, '*');



  



  if (requestAttempts < maxAttempts) {



    setTimeout(requestEmailData, 1000);



  }



}







setTimeout(requestEmailData, 500);











function setupMoreMenu() {



  const moreBtn = document.getElementById('moreTabBtn');



  const menu = document.getElementById('moreMenu');



  if (!moreBtn || !menu) return;



  function openMenu() {



    menu.hidden = false;



    moreBtn.setAttribute('aria-expanded', 'true');



    document.addEventListener('click', onDoc, { once: true });



  }



  function closeMenu() {



    menu.hidden = true;



    moreBtn.setAttribute('aria-expanded', 'false');



  }



  function onDoc(e) {



    if (!menu.contains(e.target) && e.target !== moreBtn) closeMenu();



  }



  moreBtn.addEventListener('click', function(e){



    e.stopPropagation();



    if (menu.hidden) openMenu(); else closeMenu();



  });



  Array.prototype.forEach.call(menu.querySelectorAll('.more-item'), function(item){



    item.addEventListener('click', function(){



      const target = item.getAttribute('data-target-tab');



      const hiddenBtn = document.querySelector('.tab.hidden[data-tab="' + target + '"]');



      closeMenu();



      if (hiddenBtn) hiddenBtn.click();



    });



  });



}







function setupTabs() {



  const tabs = document.querySelectorAll('.tab');



  const tabContents = document.querySelectorAll('.tab-content');



  



  if (!tabs || tabs.length === 0) return;



  



  tabs.forEach(tab => {



    tab.addEventListener('click', () => {



      tabs.forEach(t => t.classList.remove('active'));



      tabContents.forEach(content => content.classList.remove('active'));



      tab.classList.add('active');



      const tabName = tab.dataset.tab;



      const content = document.getElementById(tabName + '-tab');



      if (content) content.classList.add('active');



      



      console.log('[Panel] Switched to tab:', tabName);



      



      if (tabName === 'debug') {



        refreshDebugStats();



      }



      if (tabName === 'audit') {
        if (typeof initTrackedEntriesUI === 'function') {
          Promise.resolve(initTrackedEntriesUI())
            .then(() => {
              if (typeof renderTrackedEntriesUI === 'function') {
                renderTrackedEntriesUI();
              }
            })
            .catch(error => {
              console.error('[Panel] Failed to initialize tracked entries UI', error);
            });
        } else {
          console.warn('[Panel] initTrackedEntriesUI function not found');
        }
      }



      if (tabName === 'invite') {



        console.log('[Panel] Invite tab activated, setting up filters');



        setTimeout(() => {



          populateFilterDropdowns();



        }, 100);



      }



      if (tabName === 'add-bidder') {



        setTimeout(() => {



          setupAddBidderTab();



        }, 100);



      }



      if (tabName === 'add-bid') {



        setupAddBidTab();

      }



      if (tabName === 'projects') {



        initProjectsTab();



      }



    });



  });



}







function setupConnectButton() {



  const connectBtn = $('connectBtn');



  if (connectBtn) {



    connectBtn.addEventListener('click', async () => {



      if (!isConnected) {



        await connectToSharePoint();



      } else {



        await disconnectFromSharePoint();



      }



    });



  }



}







async function connectToSharePoint(options = {}) {



  const { initiatedByOverlay = false } = options || {};



  const connectBtn = $('connectBtn');



  const connectionText = $('connectionText');



  const statusDot = $('statusDot');







  try {



    if (connectBtn) {



      connectBtn.textContent = 'Connecting...';



      connectBtn.disabled = true;



    }



    if (connectionText) connectionText.textContent = 'Connecting...';







    if (initiatedByOverlay || !initialDataLoaded) {



      showLoadingOverlay({



        title: 'Connecting to SharePoint...',



        message: 'Complete the Microsoft sign-in prompt if it appears.',



        showSpinner: true



      });



    }







    const response = await sendWithTimeout({ type: 'AB_AUTHENTICATE' }, 60000);



    if (!response || !response.success) {



      throw new Error(response?.error || 'Authentication failed');



    }







    const baselineQueued = response?.baselineQueued === true;



    const incrementalQueued = response?.incrementalQueued === true;







    if ((baselineQueued || !response.cacheHydrated) && (initiatedByOverlay || !initialDataLoaded)) {



      showLoadingOverlay({



        title: 'Scanning SharePoint...',



        message: 'Building company cache from SharePoint...',



        showSpinner: true



      });



    } else if (incrementalQueued && initiatedByOverlay) {



      showLoadingOverlay({



        title: 'Updating company data...',



        message: 'Refreshing cached data...',



        showSpinner: true



      });



    }







    isConnected = true;







    const profileResponse = await sendWithTimeout({ type: 'AB_GET_USER_PROFILE' }, 15000);



    if (profileResponse && profileResponse.success) {



      const user = profileResponse.profile || {};



      const label = user.displayName || user.userPrincipalName || '';



      if (connectionText) connectionText.textContent = label ? 'Connected as ' + label : 'Connected';



    } else if (connectionText) {



      connectionText.textContent = 'Connected';



    }







    if (connectBtn) {



      connectBtn.textContent = 'Disconnect';



      connectBtn.disabled = false;



    }



    if (statusDot) statusDot.classList.add('connected');







    showStatus('Connected successfully!', 'success');







    await loadStatesList();







    if (response && response.cacheError) {



      console.warn('[Panel] Remote cache hydration warning:', response.cacheError);



      showStatus('Cache sync warning: ' + response.cacheError, 'warning');



    }







    if (baselineQueued) {



      showStatus('Baseline scan queued automatically.', 'info');



    } else if (incrementalQueued) {



      showStatus('Incremental scan queued automatically.', 'info');



    }







    const shouldShowOverlay = initiatedByOverlay || !initialDataLoaded;



    if (shouldShowOverlay || (response && response.cacheHydrated)) {



      await loadCachedDataOnStartup({



        showOverlay: shouldShowOverlay,



        requireData: true,



        maxAttempts: 20,



        reason: 'post-connect'



      });



    } else {



      await loadCachedDataOnStartup({ showOverlay: false, requireData: false, reason: 'post-connect' });



    }



  } catch (error) {



    console.error('[Panel] Connection error:', error);



    if (connectBtn) {



      connectBtn.textContent = 'Connect';



      connectBtn.disabled = false;



    }



    if (connectionText) connectionText.textContent = 'Disconnected';



    if (statusDot) statusDot.classList.remove('connected');



    showStatus('Connection failed: ' + (error?.message || 'Unknown error'), 'error');







    if (initiatedByOverlay || !initialDataLoaded) {



      showLoadingOverlay({



        title: 'Unable to connect',



        message: error?.message || 'Please try again.',



        actionLabel: 'Try Again',



        actionType: 'connect',



        showSpinner: false



      });



    }



  }



}



async function disconnectFromSharePoint() {



  isConnected = false;



  const connectBtn = $('connectBtn');



  const connectionText = $('connectionText');



  const statusDot = $('statusDot');



  



  if (connectBtn) connectBtn.textContent = 'Connect';



  if (connectionText) connectionText.textContent = 'Disconnected';



  if (statusDot) statusDot.classList.remove('connected');



  



  try {



    await sendWithTimeout({ type: 'AB_DISCONNECT' }, 5000);



    showStatus('Disconnected', 'info');



  } catch (error) {



    console.error('[Panel] Error disconnecting:', error);



    showStatus('Disconnected (local only)', 'info');



  }



}







async function loadStatesList() {



  try {



    console.log('[Panel] Loading states list...');



    const response = await sendWithTimeout({ type: 'AB_GET_STATES' }, 30000);



    if (response && response.success) {



      const states = Array.isArray(response.states)



        ? response.states.filter(state => state && state.name && (state.site || state.siteUrl || state.siteURL))



        : [];



      availableStatesList = states;



      const selector = $('stateSelector');



      if (selector) {



        selector.innerHTML = '<option value="">Select a state...</option>';



        states.forEach(state => {



          const option = document.createElement('option');



          const siteValue = state.site || state.siteUrl || state.siteURL;



          option.value = JSON.stringify({ name: state.name, site: siteValue });



          option.textContent = state.name;



          selector.appendChild(option);



        });



        console.log('[Panel] Loaded', states.length, 'states');



      }



    }



  } catch (error) {



    console.error('[Panel] Error loading states:', error);



  }



}







async function ensureRemoteCacheUpload() {



  try {



    const uploadResponse = await sendWithTimeout({ type: 'AB_UPLOAD_CACHE_REMOTE' }, 240000);



    if (!uploadResponse) {



      return { success: false, error: 'No response from upload' };



    }



    if (uploadResponse.success === false) {



      return { success: false, error: uploadResponse.error || 'Upload failed' };



    }



    return { success: true, manifest: uploadResponse.manifest || null };



  } catch (error) {



    console.error('[Panel] SharePoint upload failed:', error);



    return { success: false, error: error.message || String(error || 'Upload failed') };



  }



}







async function executeStateScan({ name, site, uploadAfter = true, refreshCache = true }) {



  if (!site) {



    throw new Error('Missing SharePoint site reference');



  }







  const response = await sendWithTimeout({



    type: 'AB_SCAN_STATE',



    siteUrl: site,



    stateName: name



  }, 1800000);







  if (!response || !response.success) {



    throw new Error(response?.error || 'State scan failed');



  }







  let uploadInfo = response.cacheUpload;



  if (uploadAfter !== false) {



    uploadInfo = await ensureRemoteCacheUpload();



  }



  if (uploadInfo) {



    response.cacheUpload = uploadInfo;



  }







  if (refreshCache) {



    try {



      await loadCachedDataOnStartup({ showOverlay: false, requireData: false, reason: 'post-scan' });



    } catch (error) {



      console.warn('[Panel] Cache refresh after state scan failed:', error);



    }



  }







  return response;



}











async function scanSelectedState() {



  if (bulkStateScanActive) {



    showStatus('Sequential scan already running. Please wait until it finishes.', 'warning');



    return;



  }







  const selector = $('stateSelector');



  const statusDiv = $('stateScanStatus');



  const progressContainer = $('scanProgressContainer');



  const scanStateBtn = $('scanStateBtn');







  if (!selector || !selector.value) {



    if (statusDiv) statusDiv.textContent = 'Please select a state';



    return;



  }







  const { name, site } = JSON.parse(selector.value);







  if (scanStateBtn) {



    scanStateBtn.disabled = true;



    scanStateBtn.textContent = 'Scanning...';



  }







  if (statusDiv) statusDiv.textContent = `Scanning ${name}...`;



  if (progressContainer) {



    progressContainer.style.display = 'block';



    resetProgressUI();



  }



  showStatus(`Scanning ${name}...`, 'info');







  try {



    const response = await executeStateScan({ name, site, uploadAfter: true, refreshCache: true });







    const completedState = response.state || name;



    const newProjects = response.newProjects || 0;



    const newContacts = response.newContacts || 0;



    const uploadInfo = response.cacheUpload;







    let statusMessage = 'Scan complete for ' + completedState + ' - ' + newProjects + ' new projects, ' + newContacts + ' contacts';







    if (uploadInfo && uploadInfo.success === true) {



      statusMessage += ' (SharePoint cache updated)';



    } else if (uploadInfo && uploadInfo.success === false) {



      const uploadError = uploadInfo.error || 'unknown error';



      statusMessage += ' (SharePoint upload failed: ' + uploadError + ')';



    }







    if (statusDiv) {



      statusDiv.textContent = statusMessage;



    }







    const statusType = uploadInfo && uploadInfo.success === false ? 'warning' : 'success';



    showStatus(statusMessage, statusType);



  } catch (error) {



    console.error('[Panel] Error scanning state:', error);



    if (statusDiv) statusDiv.textContent = `Scan failed: ${error.message}`;



    showStatus(`Error scanning state: ${error.message}`, 'error');



  } finally {



    if (scanStateBtn) {



      scanStateBtn.disabled = false;



      scanStateBtn.textContent = 'Scan Selected State';



    }



    if (progressContainer) {



      setTimeout(() => {



        progressContainer.style.display = 'none';



        resetProgressUI();



      }, 2000);



    }



  }



}











async function scanAllStatesSequentially() {

  if (bulkStateScanActive) {

    showStatus('Sequential state scan already in progress. Please wait.', 'warning');

    return;

  }



  if (!Array.isArray(availableStatesList) || availableStatesList.length === 0) {

    try {

      await loadStatesList();

    } catch (error) {

      console.warn('[Panel] Unable to refresh states list before bulk scan:', error);

    }

  }



  const states = Array.isArray(availableStatesList)

    ? availableStatesList.filter(state => state && state.name && (state.site || state.siteUrl || state.siteURL))

    : [];



  if (states.length === 0) {

    showStatus('No states available to scan.', 'warning');

    return;

  }



  const scanAllBtn = $('scanAllStatesBtn');

  const scanStateBtn = $('scanStateBtn');

  const statusDiv = $('stateScanStatus');

  const progressContainer = $('scanProgressContainer');



  bulkStateScanActive = true;



  if (scanAllBtn) {

    scanAllBtn.disabled = true;

    scanAllBtn.textContent = 'Scanning All...';

  }

  if (scanStateBtn) {

    scanStateBtn.disabled = true;

    scanStateBtn.textContent = 'Scan Selected State';

  }

  if (statusDiv) {

    statusDiv.textContent = `Scanning ${states.length} states sequentially...`;

  }

  if (progressContainer) {

    progressContainer.style.display = 'block';

    resetProgressUI();

  }



  const results = [];



  try {

    for (let index = 0; index < states.length; index++) {

      const stateEntry = states[index];

      const stateName = stateEntry.name || stateEntry.state || stateEntry.fullName || `State ${index + 1}`;

      const siteValue = stateEntry.site || stateEntry.siteUrl || stateEntry.siteURL;



      if (!siteValue) {

        const errorMessage = 'Missing SharePoint site reference';

        console.warn('[Panel] Skipping state due to missing site URL:', stateEntry);

        results.push({ success: false, state: stateName, error: errorMessage });

        showStatus(`Skipping ${stateName}: ${errorMessage}`, 'warning');

        continue;

      }



      if (statusDiv) {

        statusDiv.textContent = `Scanning ${stateName} (${index + 1}/${states.length})...`;

      }

      showStatus(`Scanning ${stateName} (${index + 1}/${states.length})`, 'info');



      try {

        const response = await executeStateScan({ name: stateName, site: siteValue, uploadAfter: true, refreshCache: false });

        results.push({ success: true, state: stateName, response });



        const uploadInfo = response.cacheUpload;

        let message = 'Scan complete for ' + stateName + ' - ' + (response.newProjects || 0) + ' new projects, ' + (response.newContacts || 0) + ' contacts';



        if (uploadInfo && uploadInfo.success === true) {

          message += ' (SharePoint cache updated)';

        } else if (uploadInfo && uploadInfo.success === false) {

          const uploadError = uploadInfo.error || 'unknown error';

          message += ' (SharePoint upload failed: ' + uploadError + ')';

        }



        showStatus(message, uploadInfo && uploadInfo.success === false ? 'warning' : 'success');

      } catch (error) {

        console.error('[Panel] Error during sequential state scan:', error);

        results.push({ success: false, state: stateName, error: error.message || String(error || 'Unknown error') });

        showStatus(`Scan failed for ${stateName}: ${error.message}`, 'error');

      }

    }



    const successCount = results.filter(result => result.success).length;

    const failCount = results.length - successCount;

    const summaryMessage = failCount === 0

      ? `Sequential scan complete: ${successCount} states updated successfully.`

      : `Sequential scan complete: ${successCount} succeeded, ${failCount} failed.`;



    if (statusDiv) {

      statusDiv.textContent = summaryMessage;

    }

    showStatus(summaryMessage, failCount === 0 ? 'success' : 'warning');



    try {

      await loadCachedDataOnStartup({ showOverlay: false, requireData: false, reason: 'post-scan-all' });

    } catch (error) {

      console.warn('[Panel] Cache refresh after sequential scan failed:', error);

    }

  } finally {

    if (progressContainer) {

      setTimeout(() => {

        progressContainer.style.display = 'none';

        resetProgressUI();

      }, 2000);

    }

    if (scanAllBtn) {

      scanAllBtn.disabled = false;

      scanAllBtn.textContent = 'Scan All States Sequentially';

    }

    if (scanStateBtn) {

      scanStateBtn.disabled = false;

      scanStateBtn.textContent = 'Scan Selected State';

    }

    bulkStateScanActive = false;

  }

}





function autofillBidContactFromEmail() {



  const emailInput = $('bidEmail');



  const companyInput = $('bidCompany');



  const contactInput = $('bidContact');



  const phoneInput = $('bidPhone');







  const emailValue = (emailInput?.value || '').trim().toLowerCase();



  if (!emailValue) return;







  const projectMatch = currentProjectContacts.find(c => c.email && c.email.toLowerCase() === emailValue);



  const globalMatch = allContacts.find(c => c.email && c.email.toLowerCase() === emailValue);



  const match = projectMatch || globalMatch;







  if (!match) return;







  if (match.company && companyInput && !companyInput.value) {



    companyInput.value = match.company;



  }



  if (match.name && contactInput && !contactInput.value) {



    contactInput.value = match.name;



  }



  if (match.phone && phoneInput && !phoneInput.value) {



    phoneInput.value = formatPhoneNumber(match.phone);



  }



}







function renderAddBidderDivisionOptions() {
  const divisionsContainer = $('divisionsContainer');
  if (!divisionsContainer) return;

  divisionsContainer.innerHTML = '';
  if (!currentProjectContext.divisions.length) {
    divisionsContainer.innerHTML = '<p class="empty-state">Load a project to see available divisions</p>';
    return;
  }

  const header = document.createElement('h4');
  header.textContent = 'Select Divisions:';
  divisionsContainer.appendChild(header);

  const { active: activeDivisions, notUsed: notUsedDivisions } = partitionDivisionsByUsage(currentProjectContext.divisions);

  const appendLabel = (division, target) => {
    const key = division.key || division.sheetName || division.name;
    const labelText = division.displayName || division.name || division.sheetName || key;
    const existing = currentProjectContext.contactsByDivision[key] || [];

    const label = document.createElement('label');
    label.className = 'division-checkbox';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = key;
    checkbox.className = 'division-check';
    label.appendChild(checkbox);

    const nameSpan = document.createElement('span');
    nameSpan.textContent = labelText;
    label.appendChild(nameSpan);

    if (existing.length) {
      const countSpan = document.createElement('span');
      countSpan.className = 'division-count';
      countSpan.textContent = `(${existing.length} contacts)`;
      label.appendChild(countSpan);
    }

    checkbox.addEventListener('change', () => {
      updateDivisionDuplicateWarnings();
    });

    target.appendChild(label);
  };

  activeDivisions.forEach(division => appendLabel(division, divisionsContainer));

  if (notUsedDivisions.length) {
    const details = document.createElement('details');
    details.className = 'division-not-used-group';
    const summary = document.createElement('summary');
    summary.textContent = `Not Used (${notUsedDivisions.length})`;
    details.appendChild(summary);

    const list = document.createElement('div');
    list.className = 'division-not-used-list';
    notUsedDivisions.forEach(division => appendLabel(division, list));
    details.appendChild(list);
    divisionsContainer.appendChild(details);
  }

  updateDivisionDuplicateWarnings();
}









function renderBidDivisionChecklist() {
  const container = $('bidDivisionChecklist');
  if (!container) return;

  container.innerHTML = '';
  if (!currentProjectContext.divisions.length) {
    container.innerHTML = '<p class="form-hint">Load a project to choose divisions.</p>';
    return;
  }

  const { active: activeDivisions, notUsed: notUsedDivisions } = partitionDivisionsByUsage(currentProjectContext.divisions);

  const appendOption = (division, target) => {
    const key = division.key || division.sheetName || division.name;
    const displayName = division.displayName || division.name || key;

    const option = document.createElement('label');
    option.className = 'division-bid-option';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'bid-division-check';
    checkbox.dataset.divisionKey = key;
    option.appendChild(checkbox);

    const nameSpan = document.createElement('span');
    nameSpan.className = 'division-name';
    nameSpan.textContent = displayName;
    option.appendChild(nameSpan);

    const input = document.createElement('input');
    input.type = 'number';
    input.className = 'bid-amount-input';
    input.dataset.divisionKey = key;
    input.placeholder = 'Bid Amount';
    input.min = '0';
    input.step = '0.01';
    option.appendChild(input);

    target.appendChild(option);
  };

  activeDivisions.forEach(division => appendOption(division, container));

  if (notUsedDivisions.length) {
    const details = document.createElement('details');
    details.className = 'division-not-used-group';
    const summary = document.createElement('summary');
    summary.textContent = `Not Used (${notUsedDivisions.length})`;
    details.appendChild(summary);

    const list = document.createElement('div');
    list.className = 'division-not-used-list';
    notUsedDivisions.forEach(division => appendOption(division, list));
    details.appendChild(list);
    container.appendChild(details);
  }
}
function updateDivisionDuplicateWarnings() {



  const companyInput = $('bidderCompany');



  const emailInput = $('bidderEmail');



  const phoneInput = $('bidderPhone');







  const companyKey = normalizeText(companyInput?.value || '');



  const emailLower = (emailInput?.value || '').trim().toLowerCase();



  const phoneDigits = normalizePhone(phoneInput?.value || '');



  const contactsByDivision = currentProjectContext?.contactsByDivision || {};







  const checkboxes = document.querySelectorAll('.division-check');



  checkboxes.forEach(checkbox => {



    const wrapper = checkbox.closest('.division-checkbox');



    if (!wrapper) return;



    wrapper.classList.remove('duplicate-division');







    const divisionKey = checkbox.value;



    const contacts = contactsByDivision[divisionKey] || [];



    const hasMatch = contacts.some(contact => {



      if (emailLower && contact.email && contact.email.toLowerCase() === emailLower) return true;



      if (phoneDigits && contact.phoneDigits === phoneDigits) return true;



      if (companyKey && contact.companyNormalized === companyKey) return true;



      return false;



    });







    if (hasMatch) {



      wrapper.classList.add('duplicate-division');



      checkbox.checked = false;



      checkbox.disabled = true;



    } else {



      wrapper.classList.remove('duplicate-division');



      checkbox.disabled = false;



    }



  });



}







async function submitBidder(options = {}) {



  const { override = false } = options;



  const submitBtn = $('submitBidderBtn');







  try {



    if (!currentProject) {



      showStatus('Please load a project first', 'error');



      return;



    }







    const contactData = {



      company: $('bidderCompany')?.value.trim() || '',



      name: $('bidderName')?.value.trim() || '',



      phone: $('bidderPhone')?.value.trim() || '',



      email: $('bidderEmail')?.value.trim() || '',



      notes: $('bidderNotes')?.value.trim() || ''



    };







    if (!contactData.email) {



      showStatus('Email is required', 'error');



      return;



    }







    const selectedDivisions = Array.from(



      document.querySelectorAll('.division-check:checked')



    ).map(cb => cb.value);







    if (selectedDivisions.length === 0) {



      showStatus('Please select at least one division', 'error');



      return;



    }







    if (submitBtn) {



      submitBtn.textContent = 'Adding Bidder...';



      submitBtn.disabled = true;



    }







    showStatus('Adding bidder to project...', 'info');







    const response = await sendWithTimeout({



      type: 'AB_ADD_BIDDER',



      projectName: currentProject,



      contactData,



      selectedDivisions,



      overrideConflicts: override



    }, 120000);







    if (response?.requiresOverride && !override) {



      const promptMessage = formatConflictPrompt(response.conflicts, 'This contact already appears on the following divisions\n');



      const proceed = window.confirm(`${promptMessage}\nAdd anyway?`);



      if (proceed) {



        return await submitBidder({ override: true });



      }



      showStatus('Bidder add cancelled by user', 'warning');



      return;



    }







    if (!response || !response.success) {



      throw new Error(response?.error || 'Failed to add bidder');



    }







    const { addedCount, failedCount, results } = response;







    if (failedCount > 0) {



      console.warn('[Panel] Some divisions failed:', results.filter(r => !r.success));



      showStatus(`Added to ${addedCount} divisions, ${failedCount} failed`, 'warning');



    } else {



      showStatus(`Successfully added bidder to ${addedCount} divisions!`, 'success');



    }







    clearBidderForm();



    await loadProject(currentProject);







  } catch (error) {



    console.error('[Panel] Error submitting bidder:', error);



    showStatus(`Error: ${error.message}`, 'error');



  } finally {



    if (submitBtn) {



      submitBtn.textContent = 'Add Bidder';



      submitBtn.disabled = false;



    }



  }



}







async function handleBidFormSubmit(event) {



  event.preventDefault();







  const bidForm = $('bidForm');



  const submitBtn = bidForm?.querySelector('button[type="submit"]');







  try {



    const projectInputValue = $('bidProject')?.value.trim();



    const projectName = projectInputValue || currentProject;







    if (!projectName) {



      showStatus('Enter a project name before submitting a bid', 'error');



      return;



    }







    if (!currentProject || currentProject.toLowerCase() !== projectName.toLowerCase()) {



      await loadProject(projectName);



    }







    const contactData = {



      company: $('bidCompany')?.value.trim() || '',



      name: $('bidContact')?.value.trim() || '',



      phone: $('bidPhone')?.value.trim() || '',



      email: $('bidEmail')?.value.trim() || '',



      notes: $('bidNotes')?.value.trim() || ''



    };







    if (!contactData.company) {



      showStatus('Company name is required for bid submission', 'error');



      return;



    }







    if (!contactData.email) {



      showStatus('Email is required for bid submission', 'error');



      return;



    }







    const divisionSelections = Array.from(document.querySelectorAll('.bid-division-check:checked'))



      .map(checkbox => {



        const key = checkbox.dataset.divisionKey;



        const amountInput = document.querySelector(`.bid-amount-input[data-division-key="${CSS.escape(key)}"]`);



        const bidAmount = amountInput ? amountInput.value.trim() : '';



        return {



          key,



          label: checkbox.closest('.division-bid-option')?.querySelector('.division-name')?.textContent || key,



          bidAmount



        };



      });







    if (divisionSelections.length === 0) {



      showStatus('Select at least one division to submit a bid', 'error');



      return;



    }







    if (submitBtn) {



      submitBtn.disabled = true;



      submitBtn.textContent = 'Submitting Bid...';



    }







    showStatus('Submitting bid to project...', 'info');







    const emailMeta = {



      subject: emailData.subject || '',



      from: emailData.sender || '',



      sentAt: emailData.sentAt || '',



      bodyHtml: emailData.bodyHtml || '',



      bodyText: emailData.body || '',



      messageId: emailData.messageId || ''



    };







    const attachments = Array.isArray(emailData.attachments) ? emailData.attachments : [];







    const payload = {



      type: 'AB_SUBMIT_BID',



      projectName: projectName,



      contactData,



      divisions: divisionSelections,



      attachments,



      emailMeta,



      overrideConflicts: !!emailData.forceOverride



    };







    const response = await sendWithTimeout(payload, 180000);







    if (response?.requiresOverride && !payload.overrideConflicts) {



      const message = formatConflictPrompt(response.conflicts, 'This bid conflicts with existing contacts:');



      const proceed = window.confirm(`${message}\nSubmit anyway?`);



      if (proceed) {



        emailData.forceOverride = true;



        return await handleBidFormSubmit(event);



      }



      showStatus('Bid submission cancelled by user', 'warning');



      return;



    }







    if (!response || !response.success) {



      throw new Error(response?.error || 'Failed to submit bid');



    }







    const { results = [], artifactResults = [] } = response;



    const failed = results.filter(r => !r.success);







    if (failed.length > 0) {



      showStatus(`Bid submitted with warnings. ${failed.length} divisions reported errors.`, 'warning');



      console.warn('[Panel] Bid submission failures:', failed);



    } else {



      showStatus('Bid submitted successfully!', 'success');



    }







    if (artifactResults.some(artifact => artifact.success === false)) {



      console.warn('[Panel] Some bid artifacts failed to upload:', artifactResults);



    }







    emailData.forceOverride = false;



    resetBidForm();



    await loadProject(projectName);







  } catch (error) {



    console.error('[Panel] Error submitting bid:', error);



    showStatus(`Bid submission error: ${error.message}`, 'error');



  } finally {



    if (submitBtn) {



      submitBtn.disabled = false;



      submitBtn.textContent = 'Submit Bid';



    }



  }



}







// alias for legacy callers



const submitBid = handleBidFormSubmit; window.submitBid = handleBidFormSubmit;















function resetBidForm() {



  const bidForm = $('bidForm');



  if (bidForm) {



    bidForm.reset();



  }







  document.querySelectorAll('.bid-division-check').forEach(cb => {



    cb.checked = false;



  });







  document.querySelectorAll('.bid-amount-input').forEach(input => {



    input.value = '';



  });







  emailData.forceOverride = false;



}







function clearBidderForm() {



  const companyInput = $('bidderCompany');



  const nameInput = $('bidderName');



  const phoneInput = $('bidderPhone');



  const emailInput = $('bidderEmail');



  



  if (companyInput) companyInput.value = '';



  if (nameInput) nameInput.value = '';



  if (phoneInput) phoneInput.value = '';



  if (emailInput) emailInput.value = '';



  



  document.querySelectorAll('.division-check').forEach(cb => cb.checked = false);



  emailData.forceOverride = false;



  updateDivisionDuplicateWarnings();



}







async function initProjectsTab() {



  if (allProjects.length === 0) {



    const response = await sendWithTimeout({ type: 'AB_GET_CACHED_DATA' }, 5000);



    if (response && response.success) {



      allProjects = prepareProjectList(response.projects || []);



      allProjects.sort((a, b) => compareProjectNamesDesc(getProjectDisplayName(a), getProjectDisplayName(b)));



    }



  }



  



  setupProjectSearch();



}











function setupProjectSearch() {



  refreshProjectDropdowns();







  const searchInput = $('projectSearchInput');



  if (!searchInput) return;







  if (!searchInput.dataset.enhanced) {



    searchInput.addEventListener('keydown', (event) => {



      if (event.key === 'Enter') {



        event.preventDefault();



        const projectName = searchInput.value.trim();



        if (projectName) {



          selectProject(projectName);



        }



      }



    });



    searchInput.dataset.enhanced = 'true';



  }



}







async function selectProject(projectName, options = {}) {



  const { prefetched = null, forceRefresh = false, forceContactRefresh = false } = options || {};



  const loadingMsg = $('projectLoadingMsg');



  const selectedInfo = $('selectedProjectInfo');



  const contactsContainer = $('projectContactsContainer');







  if (loadingMsg) loadingMsg.style.display = 'block';



  if (selectedInfo) selectedInfo.style.display = 'none';



  if (contactsContainer) contactsContainer.style.display = 'none';







  try {



    if (prefetched) {



      applyProjectResponse(projectName, prefetched);



      const effectiveName = currentProject || projectName;



      updateProjectSummary(effectiveName);



      showStatus(`Loaded project: ${effectiveName}`, 'success');



      return prefetched;



    }







    const response = await loadProject(projectName, { forceRefresh, forceContactRefresh });



    return response;



  } catch (error) {



    console.error('Error loading project:', error);



    if (loadingMsg) {



      loadingMsg.innerHTML = `<div style="color: #dc3545;">Error: ${error.message}</div>`;



    }



    throw error;



  }



}







function displayProjectContacts() {



  const container = $('projectContactsList');



  const contactsContainer = $('projectContactsContainer');







  if (!container) return;







  const context = currentProjectContext || {};



  const divisions = Array.isArray(context.divisions) ? context.divisions : [];



  const contactsByDivision = context.contactsByDivision || {};



  const flattenedContacts = Array.isArray(currentProjectContacts) ? currentProjectContacts : [];







  flattenedContacts.forEach((contact, index) => {



    if (contact && typeof contact === 'object') {



      contact.__flattenIndex = index;



    }



  });







  const useFallback = divisions.length === 0 || Object.keys(contactsByDivision).length === 0;



  let html = '';







  const renderContacts = (contactsArray, divisionKey, displayName) => {



    const sortedContacts = [...contactsArray].sort((a, b) => {



      const companyA = (a?.company || '').toString().toLowerCase();



      const companyB = (b?.company || '').toString().toLowerCase();



      if (companyA === companyB) {



        return (a?.name || '').toString().localeCompare(b?.name || '', undefined, { sensitivity: 'base' });



      }



      return companyA.localeCompare(companyB, undefined, { sensitivity: 'base' });



    });







    sortedContacts.forEach((contact, idx) => {



      if (!contact) return;



      const flattenIndex = Number.isFinite(contact.__flattenIndex) ? contact.__flattenIndex : -1;



      html += `



        <div class="contact-row">



          <input type="checkbox" class="contact-checkbox" data-division="${escapeHtml(displayName)}" data-division-key="${escapeHtml(divisionKey)}" data-index="${idx}" data-flatten-index="${flattenIndex}">



          <div class="contact-info">



            <div class="contact-field">



              <span class="contact-label">Company</span>



              <span class="contact-value">${escapeHtml(contact.company || '')}</span>



            </div>



            <div class="contact-field">



              <span class="contact-label">Contact</span>



              <span class="contact-value">${escapeHtml(contact.name || '-')}</span>



            </div>



            <div class="contact-field">



              <span class="contact-label">Email</span>



              <span class="contact-value">${escapeHtml(contact.email || '-')}</span>



            </div>



            <div class="contact-field">



              <span class="contact-label">Phone</span>



              <span class="contact-value">${escapeHtml(contact.phone || contact.phoneDigits || '-')}</span>



            </div>



          </div>



        </div>`;



    });



  };







  if (useFallback) {



    const fallbackMap = new Map();



    flattenedContacts.forEach(contact => {



      if (!contact) return;



      const divisionKey = contact.divisionKey || contact.division || 'Unknown';



      if (!fallbackMap.has(divisionKey)) {



        fallbackMap.set(divisionKey, []);



      }



      fallbackMap.get(divisionKey).push(contact);



    });







    const fallbackDivisions = Array.from(fallbackMap.keys())



      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));







    fallbackDivisions.forEach((divisionKey, divisionIndex) => {



      const contacts = fallbackMap.get(divisionKey) || [];



      if (!contacts.length) return;







      const displayName = contacts[0].division || divisionKey || `Division ${divisionIndex + 1}`;



      html += `<div class="division-section">`;



      html += `<div class="division-header">${escapeHtml(displayName)}</div>`;



      renderContacts(contacts, divisionKey, displayName);



      html += '</div>';



    });



  } else {



    divisions.forEach((division, divisionIndex) => {



      const divisionKey = division.key || division.sheetName || division.displayName || division.name || `division-${divisionIndex}`;



      const displayName = division.displayName || division.name || division.sheetName || divisionKey;



      const contacts = Array.isArray(contactsByDivision[divisionKey]) ? contactsByDivision[divisionKey] : [];



      if (!contacts.length) return;







      html += `<div class="division-section">`;



      html += `<div class="division-header">${escapeHtml(displayName)}</div>`;



      renderContacts(contacts, divisionKey, displayName);



      html += '</div>';



    });



  }







  if (!html.trim()) {



    html = '<div style="padding: 20px; text-align: center; color: #666;">No contacts found</div>';



  }







  flattenedContacts.forEach(contact => {



    if (contact && typeof contact === 'object' && Object.prototype.hasOwnProperty.call(contact, '__flattenIndex')) {



      delete contact.__flattenIndex;



    }



  });







  container.innerHTML = html;



  if (contactsContainer) contactsContainer.style.display = 'block';



}



const selectAllProjectContacts = $('selectAllProjectContacts');



if (selectAllProjectContacts) {



  selectAllProjectContacts.addEventListener('click', () => {



    const checkboxes = document.querySelectorAll('#projectContactsList .contact-checkbox');



    const allChecked = Array.from(checkboxes).every(cb => cb.checked);



    checkboxes.forEach(cb => {



      cb.checked = !allChecked;



    });



  });



}



const exportProjectContacts = $('exportProjectContacts');



if (exportProjectContacts) {



  exportProjectContacts.addEventListener('click', async () => {



    const checkboxes = document.querySelectorAll('#projectContactsList .contact-checkbox:checked');







    if (checkboxes.length === 0) {



      alert('Please select at least one contact to export');



      return;



    }







    const contactsByDivision = currentProjectContext?.contactsByDivision || {};



    const selected = [];







    checkboxes.forEach(cb => {



      let resolvedContact = null;







      const flattenIndex = parseInt(cb.getAttribute('data-flatten-index'), 10);



      if (Number.isInteger(flattenIndex) && currentProjectContacts[flattenIndex]) {



        resolvedContact = currentProjectContacts[flattenIndex];



      }







      if (!resolvedContact) {



        const divisionKey = cb.getAttribute('data-division-key');



        const divisionIndex = parseInt(cb.getAttribute('data-index'), 10);



        if (divisionKey && Number.isInteger(divisionIndex)) {



          const divisionContacts = contactsByDivision[divisionKey];



          if (Array.isArray(divisionContacts) && divisionContacts[divisionIndex]) {



            resolvedContact = divisionContacts[divisionIndex];



          }



        }



      }







      if (!resolvedContact) {



        const fallbackDivision = cb.getAttribute('data-division');



        const fallbackIndex = parseInt(cb.getAttribute('data-index'), 10);



        if (fallbackDivision && Number.isInteger(fallbackIndex)) {



          const fallbackContacts = currentProjectContacts.filter(contact => (contact.divisionKey || contact.division || 'Unknown') === fallbackDivision);



          if (fallbackContacts[fallbackIndex]) {



            resolvedContact = fallbackContacts[fallbackIndex];



          }



        }



      }







      if (resolvedContact) {



        selected.push(resolvedContact);



      }



    });







    if (selected.length === 0) {



      alert('Unable to resolve the selected contacts. Please try again.');



      return;



    }







    const response = await sendWithTimeout({



      type: 'AB_EXPORT_EXCEL',



      contacts: selected



    }, 60000);







    if (response && response.success) {



      alert(`Exported ${selected.length} contact${selected.length !== 1 ? 's' : ''}`);



    } else {



      alert('Export failed: ' + (response?.error || 'Unknown error'));



    }



  });



}



function setupUniversalFilterDropdown(inputId, dataArray, allowMultipleOrOptions = true, extraOptions = {}) {



  const input = $(inputId);



  if (!input) {



    console.log('[Panel] Input not found:', inputId);



    return;



  }







  let options = {};







  if (typeof allowMultipleOrOptions === 'object') {



    options = allowMultipleOrOptions || {};



  } else {



    options = typeof extraOptions === 'object' ? extraOptions : {};



    options.allowMultiple = allowMultipleOrOptions;



  }







  const mergedConfig = {



    ...DROPDOWN_DEFAULT_OPTIONS,



    ...options



  };







  if (typeof mergedConfig.allowMultiple !== 'boolean') {



    mergedConfig.allowMultiple = true;



  }







  mergedConfig.minChars = Math.max(0, Number(mergedConfig.minChars ?? DROPDOWN_DEFAULT_OPTIONS.minChars));



  mergedConfig.maxResults = Math.max(5, Number(mergedConfig.maxResults ?? DROPDOWN_DEFAULT_OPTIONS.maxResults));



  mergedConfig.showAllOnFocus = mergedConfig.showAllOnFocus !== false;



  mergedConfig.filterMode = mergedConfig.filterMode === 'startsWith' ? 'startsWith' : 'contains';



  mergedConfig.onSelect = typeof mergedConfig.onSelect === 'function' ? mergedConfig.onSelect : null;



  mergedConfig.labelBuilder = typeof mergedConfig.labelBuilder === 'function' ? mergedConfig.labelBuilder : null;







  const preserveOrder = mergedConfig.preserveOrder === true;



  const dropdownConfig = { ...mergedConfig };



  delete dropdownConfig.preserveOrder;







  const sanitizedData = preserveOrder



    ? uniquePreserveOrder(dataArray)



    : sanitizeDataArray(dataArray);







  let entry = dropdownDataMap.get(inputId);







  if (!entry) {



    const wrapper = document.createElement('div');



    wrapper.className = 'filter-wrapper';







    const inputContainer = document.createElement('div');



    inputContainer.className = 'filter-input-container';







    input.parentElement.insertBefore(wrapper, input);



    wrapper.appendChild(inputContainer);



    inputContainer.appendChild(input);







    const arrow = document.createElement('button');



    arrow.className = 'filter-dropdown-arrow';



    arrow.type = 'button';



    arrow.setAttribute('aria-label', 'Toggle suggestions');



    arrow.innerHTML = '&#9662;';



    inputContainer.appendChild(arrow);







    const menu = document.createElement('div');



    menu.className = 'filter-dropdown-menu';



    menu.id = `${inputId}-dropdown`;



    wrapper.appendChild(menu);







    entry = {



      inputId,



      input,



      arrow,



      menu,



      wrapper,



      data: sanitizedData,



      lastQuery: '',



      ...dropdownConfig



    };







    dropdownDataMap.set(inputId, entry);







    arrow.addEventListener('click', (event) => {



      event.preventDefault();



      event.stopPropagation();



      toggleDropdown(entry);



    });







    input.addEventListener('focus', () => {



      if (entry.showAllOnFocus) {



        toggleDropdown(entry, true);



      }



    });







    input.addEventListener('click', () => {



      if (entry.showAllOnFocus && !entry.menu.classList.contains('show')) {



        toggleDropdown(entry, true);



      }



    });







    input.addEventListener('input', () => handleDropdownInput(entry));



    input.addEventListener('keydown', (event) => handleDropdownKeydown(event, entry));



  } else {



    Object.assign(entry, dropdownConfig);



    entry.data = sanitizedData;



  }







  entry.data = sanitizedData;







  if (entry.menu.classList.contains('show')) {



    renderDropdownMenu(entry, entry.lastQuery);



  }



}







function toggleDropdown(entry, forceOpen = false) {



  if (!entry || !entry.menu) return;







  document.querySelectorAll('.filter-dropdown-menu.show').forEach(menu => {



    if (menu !== entry.menu) {



      menu.classList.remove('show');



    }



  });







  const query = entry.input.value.trim();







  if (forceOpen) {



    renderDropdownMenu(entry, query);



    if (entry.menu.dataset.itemsCount !== '0') {



      entry.menu.classList.add('show');



    }



    return;



  }







  if (entry.menu.classList.contains('show')) {



    entry.menu.classList.remove('show');



    return;



  }







  renderDropdownMenu(entry, query);



  if (entry.menu.dataset.itemsCount !== '0') {



    entry.menu.classList.add('show');



  }



}







function handleDropdownInput(entry) {



  const query = entry.input.value.trim();



  entry.lastQuery = query;







  if (!query && !entry.showAllOnFocus) {



    entry.menu.classList.remove('show');



    return;



  }







  if (!query && entry.showAllOnFocus) {



    renderDropdownMenu(entry, '');



    if (entry.menu.dataset.itemsCount !== '0') {



      entry.menu.classList.add('show');



    }



    return;



  }







  if (query.length < entry.minChars) {



    entry.menu.classList.remove('show');



    return;



  }







  renderDropdownMenu(entry, query);



  if (entry.menu.dataset.itemsCount !== '0') {



    entry.menu.classList.add('show');



  } else {



    entry.menu.classList.remove('show');



  }



}







function handleDropdownKeydown(event, entry) {



  if (event.key === 'Enter' && !entry.allowMultiple) {



    const value = entry.input.value.trim();



    if (value) {



      entry.menu.classList.remove('show');



      if (entry.onSelect) {



        entry.onSelect(value);



      }



    }



  }



}







function getDropdownSuggestions(entry, query = '') {



  const data = entry.data || [];



  if (!query) {



    return data.slice(0, entry.maxResults);



  }







  const normalizedQuery = query.toLowerCase();



  const tokens = normalizedQuery.split(/\s+/).filter(Boolean);



  if (tokens.length === 0) {



    return data.slice(0, entry.maxResults);



  }







  const startsWithMatches = [];



  const otherMatches = [];







  data.forEach(item => {



    const lower = item.toLowerCase();



    const matches = entry.filterMode === 'startsWith'



      ? lower.startsWith(tokens[0])



      : tokens.every(token => lower.includes(token));







    if (!matches) return;







    if (lower.startsWith(tokens[0])) {



      startsWithMatches.push(item);



    } else {



      otherMatches.push(item);



    }



  });







  const combined = [...startsWithMatches, ...otherMatches.filter(item => !startsWithMatches.includes(item))];



  return combined.slice(0, entry.maxResults);



}







function renderDropdownMenu(entry, query = '') {



  const { menu, input, allowMultiple, labelBuilder } = entry;



  if (!menu) return;







  const suggestions = getDropdownSuggestions(entry, query);



  entry.menu.dataset.itemsCount = String(suggestions.length);



  entry.lastQuery = query;







  if (suggestions.length === 0) {



    menu.innerHTML = '<div class="filter-dropdown-empty">No matches found</div>';



    return;



  }







  const selectedValues = input.dataset.selectedValues



    ? input.dataset.selectedValues.split(',').filter(Boolean)



    : [];







  const totalItems = entry.data.length;



  const truncated = suggestions.length === entry.maxResults && entry.maxResults < totalItems;







  let html = `<div class="filter-dropdown-count">${suggestions.length} match${suggestions.length === 1 ? '' : 'es'}${totalItems > suggestions.length ? ` of ${totalItems}` : ''}</div>`;







  if (allowMultiple && entry.data.length > 1) {



    html += `



      <div class="filter-dropdown-controls">



        <button type="button" class="select-all-btn">Select All</button>



        <button type="button" class="deselect-all-btn">Clear</button>



      </div>



    `;



  }







  suggestions.forEach((item, index) => {



    const value = String(item);



    const labelText = labelBuilder ? labelBuilder(item, query) : value;



    const highlightedLabel = highlightMatch(labelText, query);



    const isChecked = selectedValues.includes(value);



    const checkboxId = `${entry.inputId}-item-${index}`;







    if (allowMultiple) {



      html += `



        <div class="filter-dropdown-item">



          <input type="checkbox" id="${checkboxId}" value="${escapeHtml(value)}" ${isChecked ? 'checked' : ''}>



          <label for="${checkboxId}">${highlightedLabel}</label>



        </div>



      `;



    } else {



      html += `



        <div class="filter-dropdown-item" data-value="${escapeHtml(value)}">



          <label>${highlightedLabel}</label>



        </div>



      `;



    }



  });







  if (truncated) {



    html += `<div class="filter-dropdown-footnote">Showing first ${entry.maxResults} matches</div>`;



  }







  menu.innerHTML = html;







  if (allowMultiple) {



    const selectAllBtn = menu.querySelector('.select-all-btn');



    if (selectAllBtn) {



      selectAllBtn.addEventListener('click', (event) => {



        event.preventDefault();



        menu.querySelectorAll('input[type="checkbox"]').forEach(cb => {



          cb.checked = true;



        });



        updateInputFromSelection(entry);



      });



    }







    const deselectAllBtn = menu.querySelector('.deselect-all-btn');



    if (deselectAllBtn) {



      deselectAllBtn.addEventListener('click', (event) => {



        event.preventDefault();



        menu.querySelectorAll('input[type="checkbox"]').forEach(cb => {



          cb.checked = false;



        });



        updateInputFromSelection(entry);



      });



    }







    menu.querySelectorAll('input[type="checkbox"]').forEach(cb => {



      cb.addEventListener('change', () => {



        updateInputFromSelection(entry);



      });



    });



  } else {



    menu.querySelectorAll('.filter-dropdown-item').forEach(item => {



      item.addEventListener('click', () => {



        const value = item.dataset.value;



        input.value = value;



        input.dataset.selectedValues = value;



        menu.classList.remove('show');



        if (entry.onSelect) {



          entry.onSelect(value);



        }



      });



    });



  }



}







function updateInputFromSelection(entry) {



  if (!entry.allowMultiple) return;







  const checked = Array.from(entry.menu.querySelectorAll('input[type="checkbox"]:checked'))



    .map(cb => cb.value);







  entry.input.dataset.selectedValues = checked.join(',');







  if (checked.length === 0) {



    entry.input.value = '';



  } else if (checked.length === 1) {



    entry.input.value = checked[0];



  } else {



    entry.input.value = `${checked.length} items selected`;



  }







  if (entry.onSelect) {



    entry.onSelect([...checked]);



  }



}











function populateFilterDropdowns() {



  console.log('[Panel] Setting up filter dropdowns with', allContacts.length, 'contacts');







  const companies = getUniqueContactValues('company');



  const contactNames = getUniqueContactValues('name');







  const emails = sanitizeDataArray(allContacts



    .map(c => c.email)



    .filter(email => email && email.includes('@') && email.includes('.'))



  );







  const cities = sanitizeDataArray(allContacts



    .map(c => {



      const city = c.city;



      if (!city) return null;



      const cityStr = String(city).trim();



      if (/^\d/.test(cityStr)) return null;



      if (cityStr.length > 30) return null;



      if (/\b(hwy|highway|street|st|road|rd|drive|dr|avenue|ave|blvd)\b/i.test(cityStr)) return null;



      return cityStr;



    })



    .filter(Boolean)



  );







  const projects = sanitizeDataArray(allContacts.map(c => c.project).filter(Boolean));



  const states = sanitizeDataArray(allContacts.map(c => c.state).filter(Boolean));



  const divisions = sanitizeDataArray(allContacts.map(c => c.division).filter(Boolean));



  const zips = sanitizeDataArray(allContacts.map(c => c.zip).filter(Boolean));



  const phones = sanitizeDataArray(allContacts.map(c => formatPhoneNumber(c.phone)).filter(Boolean));







  setupUniversalFilterDropdown('companyFilter', companies, { allowMultiple: true, maxResults: 150 });



  setupUniversalFilterDropdown('contactNameFilter', contactNames, { allowMultiple: true, maxResults: 150 });



  setupUniversalFilterDropdown('emailFilter', emails, { allowMultiple: true, maxResults: 150, filterMode: 'contains', showAllOnFocus: false });



  setupUniversalFilterDropdown('cityFilter', cities, { allowMultiple: true, maxResults: 150 });



  setupUniversalFilterDropdown('projectFilter', projects, { allowMultiple: true, maxResults: 150 });







  const stateFilter = $('stateFilter');



  if (stateFilter) {



    stateFilter.innerHTML = '<option value="">All States</option>' +



      states.map(s => `<option value="${s}">${s}</option>`).join('');



  }







  const divisionFilter = $('divisionFilter');



  if (divisionFilter) {



    divisionFilter.innerHTML = '<option value="">All Divisions</option>' +



      divisions.map(d => `<option value="${d}">${d}</option>`).join('');



  }







  if ($('zipFilter')) {



    setupUniversalFilterDropdown('zipFilter', zips, { allowMultiple: false, maxResults: 100, showAllOnFocus: false, filterMode: 'startsWith' });



  }







  refreshContactDrivenDropdowns({



    companies,



    contactNames,



    emails,



    phones



  });







  refreshProjectDropdowns();



}







function refreshContactDrivenDropdowns(precomputed = {}, extraOptions = {}) {



  const companies = precomputed.companies || getUniqueContactValues('company');



  const contactNames = precomputed.contactNames || getUniqueContactValues('name');



  const emails = precomputed.emails || sanitizeDataArray(allContacts



    .map(c => c.email)



    .filter(email => email && email.includes('@') && email.includes('.'))



  );



  const phones = precomputed.phones || sanitizeDataArray(allContacts



    .map(c => formatPhoneNumber(c.phone))



    .filter(Boolean)



  );







  const companyDropdownOptions = {



    allowMultiple: false,



    maxResults: 120,



    ...(extraOptions.companyOptions || {})



  };







  if (extraOptions.companyLabelBuilder) {



    companyDropdownOptions.labelBuilder = extraOptions.companyLabelBuilder;



  }







  setupUniversalFilterDropdown('bidderCompany', companies, companyDropdownOptions);



  setupUniversalFilterDropdown('bidCompany', companies, { allowMultiple: false, maxResults: 120 });







  setupUniversalFilterDropdown('bidderName', contactNames, { allowMultiple: false, maxResults: 120 });



  setupUniversalFilterDropdown('bidContact', contactNames, { allowMultiple: false, maxResults: 120 });







  setupUniversalFilterDropdown('bidderEmail', emails, { allowMultiple: false, maxResults: 120, filterMode: 'contains', showAllOnFocus: false });



  setupUniversalFilterDropdown('bidEmail', emails, { allowMultiple: false, maxResults: 120, filterMode: 'contains', showAllOnFocus: false });







  setupUniversalFilterDropdown('bidderPhone', phones, { allowMultiple: false, maxResults: 120, filterMode: 'startsWith', showAllOnFocus: false });



  setupUniversalFilterDropdown('bidPhone', phones, { allowMultiple: false, maxResults: 120, filterMode: 'startsWith', showAllOnFocus: false });



}







function refreshProjectDropdowns(projectNames) {



  const projects = Array.isArray(projectNames)



    ? normalizeProjectNameList(projectNames)



    : buildProjectNameList(allProjects);







  const baseOptions = { allowMultiple: false, maxResults: 200, preserveOrder: true, labelBuilder: labelProjectCodeFirst };



  setupUniversalFilterDropdown('bidderProjectInput', projects, { ...baseOptions });



  setupUniversalFilterDropdown('bidProject', projects, { ...baseOptions });



  setupUniversalFilterDropdown('projectSearchInput', projects, { ...baseOptions, showAllOnFocus: true, onSelect: (value) => selectProject(value) });



}











function findBestContactMatch(signature = {}) {



  if (!allContacts || allContacts.length === 0) return null;







  const normalizedSender = (emailData.sender || '').toLowerCase();



  const signatureEmails = Array.isArray(signature.emails) ? signature.emails : [];



  const emailCandidates = sanitizeDataArray([



    normalizedSender,



    ...signatureEmails.map(email => email.toLowerCase())



  ].filter(Boolean));







  const senderDomain = normalizedSender.includes('@') ? normalizedSender.split('@')[1] : '';







  const signatureName = normalizeText(signature.name || emailData.senderName);



  const signatureCompany = normalizeText(signature.company);







  const signaturePhones = Array.isArray(signature.phones)



    ? signature.phones.map(normalizePhone).filter(Boolean)



    : [];







  const signaturePhonePrimary = normalizePhone(signature.phone);



  if (signaturePhonePrimary) {



    signaturePhones.push(signaturePhonePrimary);



  }







  let bestMatch = null;



  let bestScore = 0;







  allContacts.forEach(contact => {



    let score = 0;



    const contactEmail = (contact.email || '').toLowerCase();



    const contactDomain = contactEmail.includes('@') ? contactEmail.split('@')[1] : '';



    const contactName = normalizeText(contact.name);



    const contactCompany = normalizeText(contact.company);



    const contactPhoneDigits = normalizePhone(contact.phone);







    if (contactEmail && emailCandidates.includes(contactEmail)) {



      score += 140;



    } else if (contactEmail && senderDomain && contactDomain === senderDomain) {



      score += 55;



    }







    if (signatureName && contactName) {



      if (contactName === signatureName) {



        score += 70;



      } else if (contactName.includes(signatureName) || signatureName.includes(contactName)) {



        score += 40;



      }



    }







    if (signatureCompany && contactCompany) {



      if (contactCompany === signatureCompany) {



        score += 55;



      } else if (contactCompany.includes(signatureCompany) || signatureCompany.includes(contactCompany)) {



        score += 30;



      }



    }







    if (contactPhoneDigits && signaturePhones.includes(contactPhoneDigits)) {



      score += 45;



    }







    if (score > bestScore) {



      bestScore = score;



      bestMatch = contact;



    }



  });







  if (!bestMatch) return null;







  return { contact: bestMatch, score: bestScore };



}







function setupInviteTab() {



  const applyBtn = $('applyFiltersBtn');



  if (applyBtn && !applyBtn.dataset.enhanced) {



    applyBtn.addEventListener('click', async () => {



      await applyInviteFilters();



    });



    applyBtn.dataset.enhanced = 'true';



  }



  



  const clearBtn = $('clearFiltersBtn');



  if (clearBtn) {



    clearBtn.addEventListener('click', clearInviteFilters);



  }



  



  const sendEmailBtn = $('sendEmailBtn');



  if (sendEmailBtn) {



    sendEmailBtn.addEventListener('click', openEmailCompose);



  }



  



  const copyBtn = $('copyContactsBtn');



  if (copyBtn) {



    copyBtn.addEventListener('click', copySelectedContacts);



  }



  



  const exportBtn = $('exportInvitesBtn');



  if (exportBtn) {



    exportBtn.addEventListener('click', exportSelectedContacts);



  }



}







async function applyInviteFilters() {



  let filtered = [...allContacts];







  const companyInput = $('companyFilter');



  const contactInput = $('contactNameFilter');



  const emailInput = $('emailFilter');



  const cityInput = $('cityFilter');



  const projectInput = $('projectFilter');







  const companyFilter = companyInput?.dataset.selectedValues?.split(',').filter(Boolean) ||



    (companyInput?.value ? [companyInput.value] : []);



  const contactFilter = contactInput?.dataset.selectedValues?.split(',').filter(Boolean) ||



    (contactInput?.value ? [contactInput.value] : []);



  const emailFilter = emailInput?.dataset.selectedValues?.split(',').filter(Boolean) ||



    (emailInput?.value ? [emailInput.value] : []);



  const cityFilter = cityInput?.dataset.selectedValues?.split(',').filter(Boolean) ||



    (cityInput?.value ? [cityInput.value] : []);



  const projectFilter = projectInput?.dataset.selectedValues?.split(',').filter(Boolean) ||



    (projectInput?.value ? [projectInput.value] : []);







  const stateFilter = $('stateFilter')?.value;



  const divisionFilter = $('divisionFilter')?.value;



  const zipFilterValue = $('zipFilter')?.value?.trim();



  const radiusValue = parseFloat($('radiusFilter')?.value || '');







  if (companyFilter.length > 0) {



    filtered = filtered.filter(c => companyFilter.includes(c.company));



  }



  if (contactFilter.length > 0) {



    filtered = filtered.filter(c => contactFilter.includes(c.name));



  }



  if (emailFilter.length > 0) {



    filtered = filtered.filter(c => emailFilter.includes(c.email));



  }



  if (cityFilter.length > 0) {



    filtered = filtered.filter(c => cityFilter.includes(c.city));



  }



  if (projectFilter.length > 0) {



    filtered = filtered.filter(c => projectFilter.includes(c.project));



  }



  if (stateFilter) {



    filtered = filtered.filter(c => c.state === stateFilter);



  }



  if (divisionFilter) {



    filtered = filtered.filter(c => c.division === divisionFilter);



  }







  if (radiusValue && !zipFilterValue) {



    showStatus('Enter a zip code to use radius filtering', 'warning');



  }







  if (radiusValue && zipFilterValue) {



    filtered = await filterContactsByRadius(filtered, zipFilterValue, radiusValue);



  } else if (zipFilterValue) {



    filtered = filtered.filter(c => (c.zip || '').startsWith(zipFilterValue));



  }







  displayInviteResults(filtered);



}







function clearInviteFilters() {



  const inputs = ['companyFilter', 'contactNameFilter', 'emailFilter', 'cityFilter', 'projectFilter'];



  



  inputs.forEach(id => {



    const input = $(id);



    if (input) {



      input.value = '';



      input.dataset.selectedValues = '';



    }



  });



  



  const stateFilter = $('stateFilter');



  const divisionFilter = $('divisionFilter');



  const zipFilter = $('zipFilter');



  const radiusFilter = $('radiusFilter');



  



  if (stateFilter) stateFilter.value = '';



  if (divisionFilter) divisionFilter.value = '';



  if (zipFilter) zipFilter.value = '';



  if (radiusFilter) radiusFilter.value = '';



  



  displayInviteResults(allContacts);



}







function displayInviteResults(contacts) {



  const container = $('inviteResultsContainer');



  const countSpan = $('inviteResultCount');



  



  if (!container) {



    console.log('[Panel] inviteResultsContainer not found!');



    return;



  }



  



  if (contacts.length === 0) {



    container.innerHTML = '<p class="no-results">No contacts match your filters.</p>';



    if (countSpan) countSpan.textContent = '0 contacts found';



    return;



  }



  



  if (countSpan) {



    countSpan.textContent = `${contacts.length} contacts found`;



  }



  



  const tableWrapper = document.createElement('div');



  tableWrapper.className = 'table-wrapper';



  



  const table = document.createElement('table');



  table.className = 'contacts-table';



  



  const thead = document.createElement('thead');



  thead.innerHTML = `



    <tr>



      <th class="checkbox-col"><input type="checkbox" id="selectAllInvites"></th>



      <th class="company-col">Company Name</th>



      <th class="name-col">Contact Name</th>



      <th class="phone-col">Phone</th>



      <th class="email-col">Email</th>



      <th class="city-col">City</th>



      <th class="state-col">State</th>



    </tr>



  `;



  table.appendChild(thead);



  



  const tbody = document.createElement('tbody');



  contacts.forEach(contact => {



    const tr = document.createElement('tr');



    tr.innerHTML = `



      <td class="checkbox-col"><input type="checkbox" class="contact-checkbox" data-email="${contact.email || ''}" data-name="${contact.name || ''}" data-company="${contact.company || ''}"></td>



      <td class="company-col" title="${contact.company || ''}">${contact.company || 'Ã¢â¬â'}</td>



      <td class="name-col" title="${contact.name || ''}">${contact.name || 'Ã¢â¬â'}</td>



      <td class="phone-col" title="${contact.phone || ''}">${contact.phone || 'Ã¢â¬â'}</td>



      <td class="email-col" title="${contact.email || ''}">${contact.email || 'Ã¢â¬â'}</td>



      <td class="city-col" title="${contact.city || ''}">${contact.city || 'Ã¢â¬â'}</td>



      <td class="state-col" title="${contact.state || ''}">${contact.state || 'Ã¢â¬â'}</td>



    `;



    tbody.appendChild(tr);



  });



  table.appendChild(tbody);



  



  tableWrapper.appendChild(table);



  container.innerHTML = '';



  container.appendChild(tableWrapper);



  



  const selectAllCheckbox = document.getElementById('selectAllInvites');



  if (selectAllCheckbox) {



    selectAllCheckbox.addEventListener('change', (e) => {



      const checkboxes = container.querySelectorAll('.contact-checkbox');



      checkboxes.forEach(cb => cb.checked = e.target.checked);



      updateSelectedContacts();



    });



  }



  



  const checkboxes = container.querySelectorAll('.contact-checkbox');



  checkboxes.forEach(cb => {



    cb.addEventListener('change', updateSelectedContacts);



  });



}







function updateSelectedContacts() {



  const checkboxes = document.querySelectorAll('.contact-checkbox:checked');



  const emails = Array.from(checkboxes).map(cb => cb.dataset.email).filter(Boolean);



  



  selectedContacts = allContacts.filter(c => emails.includes(c.email));



  



  const sendBtn = $('sendEmailBtn');



  const copyBtn = $('copyContactsBtn');



  const exportBtn = $('exportInvitesBtn');



  



  const hasSelection = selectedContacts.length > 0;



  if (sendBtn) sendBtn.disabled = !hasSelection;



  if (copyBtn) copyBtn.disabled = !hasSelection;



  if (exportBtn) exportBtn.disabled = !hasSelection;



}







function openEmailCompose() {



  if (selectedContacts.length === 0) return;



  



  const emails = ['***SELECT MAILMERGE***', ...selectedContacts.map(c => c.email).filter(Boolean)];



  const emailList = emails.join(';');



  



  const composeUrl = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(emailList)}`;



  window.open(composeUrl, '_blank');



  



  showStatus(`Opening email with ${selectedContacts.length} contacts`, 'success');



}







function copySelectedContacts() {



  if (selectedContacts.length === 0) return;



  



  const text = selectedContacts.map(c => 



    `${c.company}\t${c.name}\t${c.email}\t${c.phone}\t${c.division}\t${c.project}\t${c.state}`



  ).join('\n');



  



  const header = 'Company\tName\tEmail\tPhone\tDivision\tProject\tState\n';



  



  navigator.clipboard.writeText(header + text).then(() => {



    showStatus(`Copied ${selectedContacts.length} contacts to clipboard`, 'success');



  }).catch(err => {



    showStatus('Failed to copy to clipboard', 'error');



  });



}







async function exportSelectedContacts() {



  if (selectedContacts.length === 0) return;



  



  try {



    const response = await sendWithTimeout({ 



      type: 'AB_EXPORT_EXCEL',



      contacts: selectedContacts



    }, 60000);



    



    if (response && response.success) {



      showStatus(`Exported ${selectedContacts.length} contacts to Excel`, 'success');



    } else {



      showStatus('Export failed', 'error');



    }



  } catch (error) {



    showStatus('Export error', 'error');



  }



}







function setupDebugButtons() {



  console.log('[Panel] Setting up debug buttons');



  



  const refreshStatsBtn = $('refreshStatsBtn');



  if (refreshStatsBtn) {



    refreshStatsBtn.addEventListener('click', () => {



      refreshDebugStats();



    });



  }



  



  const clearCacheBtn = $('clearCacheBtn');



  if (clearCacheBtn) {



    clearCacheBtn.addEventListener('click', async () => {



      try {



        const response = await sendWithTimeout({ type: 'AB_CLEAR_CACHE' }, 5000);



        if (response && response.success) {



          allProjects = [];



          allContacts = [];



          showStatus('Cache cleared', 'info');



        }



      } catch (error) {



        showStatus('Error clearing cache', 'error');



      }



    });



  }



  



  const scanStateBtn = $('scanStateBtn');



  if (scanStateBtn) {



    scanStateBtn.addEventListener('click', scanSelectedState);



  }







  const scanAllStatesBtn = $('scanAllStatesBtn');



  if (scanAllStatesBtn) {



    scanAllStatesBtn.addEventListener('click', scanAllStatesSequentially);



  }







  const manualBackupBtn = $('manualBackupBtn');



  if (manualBackupBtn && !manualBackupBtn.dataset.bound) {



    manualBackupBtn.addEventListener('click', async () => {



      console.log('[Panel] BACKUP BUTTON CLICKED');



      try {



        const response = await sendWithTimeout({ type: 'AB_SAVE_CACHE_BACKUP' }, 5000);



        if (response && response.success) {



          showStatus('Backup saved successfully', 'success');



        }



      } catch (error) {



        showStatus('Backup failed', 'error');



      }



    });



    manualBackupBtn.dataset.bound = 'true';



  }



  



  const uploadRemoteBtn = $('uploadRemoteCacheBtn');



  if (uploadRemoteBtn && !uploadRemoteBtn.dataset.bound) {



    uploadRemoteBtn.addEventListener('click', async () => {



      try {



        uploadRemoteBtn.disabled = true;



        uploadRemoteBtn.textContent = 'Uploading...';



        showStatus('Uploading cache to SharePoint...', 'info');



        const response = await sendWithTimeout({ type: 'AB_UPLOAD_CACHE_REMOTE' }, 180000);



        if (response && response.success) {



          showStatus('Cache uploaded to SharePoint', 'success');



        } else {



          throw new Error(response?.error || 'Upload failed');



        }



      } catch (error) {



        console.error('[Panel] Remote upload failed:', error);



        showStatus(`SharePoint upload failed: ${error.message}`, 'error');



      } finally {



        uploadRemoteBtn.disabled = false;



        uploadRemoteBtn.textContent = 'Upload Cache to SharePoint';



      }



    });



    uploadRemoteBtn.dataset.bound = 'true';



  }



  



  const downloadRemoteBtn = $('downloadRemoteCacheBtn');



  if (downloadRemoteBtn && !downloadRemoteBtn.dataset.bound) {



    downloadRemoteBtn.addEventListener('click', async () => {



      try {



        downloadRemoteBtn.disabled = true;



        downloadRemoteBtn.textContent = 'Downloading...';



        showStatus('Downloading cache from SharePoint...', 'info');



        const response = await sendWithTimeout({ type: 'AB_DOWNLOAD_CACHE_REMOTE' }, 180000);



        if (response && response.success) {



          showStatus('Cache downloaded from SharePoint', 'success');



          await loadCachedDataOnStartup({ showOverlay: true, requireData: true, maxAttempts: 20, reason: 'manual-download' });



        } else {



          throw new Error(response?.error || 'Download failed');



        }



      } catch (error) {



        console.error('[Panel] Remote download failed:', error);



        showStatus(`SharePoint download failed: ${error.message}`, 'error');



      } finally {



        downloadRemoteBtn.disabled = false;



        downloadRemoteBtn.textContent = 'Download Cache from SharePoint';



      }



    });



    downloadRemoteBtn.dataset.bound = 'true';



  }



  



  const restoreBackupBtn = $('restoreBackupBtn');



  const restoreFileInput = $('restoreFileInput');



  



  if (restoreBackupBtn && restoreFileInput && !restoreBackupBtn.dataset.bound) {



    restoreBackupBtn.addEventListener('click', () => {



      restoreFileInput.click();



    });



    



    restoreFileInput.addEventListener('change', async (e) => {



      const file = e.target.files[0];



      if (!file) return;



      



      const reader = new FileReader();



      reader.onload = async (event) => {



        try {



          const fileContent = event.target.result;



          const response = await sendWithTimeout({ 



            type: 'AB_LOAD_CACHE_FROM_FILE',



            fileContent: fileContent



          }, 10000);



          



          if (response && response.success) {



            allContacts = response.contacts || [];



            allProjects = prepareProjectList(response.projects || []);



            showStatus(`Restored ${response.contactsLoaded} contacts!`, 'success');



            displayContactsList();



            populateFilterDropdowns();



            refreshDebugStats();



            



            const addBidderTab = document.getElementById('add-bidder-tab');



            if (addBidderTab && addBidderTab.classList.contains('active')) {



              setTimeout(() => {



                setupAddBidderTab();



              }, 100);



            }



          }



        } catch (error) {



          showStatus('Restore failed: ' + error.message, 'error');



        }



      };



      reader.readAsText(file);



    });



  }



}







function refreshDebugStats() {



  (async () => {



    try {



      const resp = await sendWithTimeout({ type: 'AB_GET_STATISTICS' }, 10000);



      if (!resp || !resp.success) return;



      const s = resp.stats || {};







      const set = (id, v) => { 



        const el = $(id); 



        if (el) el.textContent = String(v ?? 0); 



      };



      



      set('debugProjects', s.totalProjects);



      set('debugContacts', s.totalContacts);



      set('debugCompanies', s.uniqueCompanies);



      set('debugDivisions', s.uniqueDivisions);



      set('debugConnection', isConnected ? 'Connected' : 'Disconnected');



      



      const last = $('debugLastScan');



      if (last) last.textContent = s.lastScanTime ? new Date(s.lastScanTime).toLocaleString() : 'Never';







      const renderList = (map, containerId) => {



        const el = $(containerId);



        if (!el) return;



        const entries = Object.entries(map || {}).sort((a,b) => b[1]-a[1]);



        el.innerHTML = entries.length



          ? entries.map(([k,v]) => `<div>${k}: ${v}</div>`).join('')



          : '<div>None</div>';



      };



      





      renderList(s.contactsByState, 'debugStateBreakdown');



      renderList(s.contactsByDivision, 'debugDivisionBreakdown');



    } catch (e) {



      console.error('[Panel] Stats error:', e);



    }



  })();



}








function setAuditLogStatus(message = '', level = 'info') {
  const statusEl = $('auditLogStatus');
  if (!statusEl) return;
  statusEl.textContent = message || '';
  statusEl.dataset.level = level;
}

function renderAuditLogEntries(entries = []) {
  const container = $('auditLogTable');
  if (!container) return;

  if (!entries.length) {
    container.innerHTML = '<p class="form-hint">No audit entries recorded yet.</p>';
    return;
  }

  const rows = entries.map(entry => {
    const detected = entry.detectedAt ? formatDateTime(entry.detectedAt) : '';
    const workbook = entry.workbookName || entry.workbookPath || entry.projectName || '';
    const addedBy = entry.lastModifiedBy || 'Unknown';
    return `<tr><td>${escapeHtml(entry.email || '')}</td><td>${escapeHtml(workbook)}</td><td>${escapeHtml(addedBy)}</td><td>${escapeHtml(detected)}</td></tr>`;
  }).join('');

  container.innerHTML = `
    <div class="audit-project-card">
      <table class="audit-table">
        <thead>
          <tr>
            <th>Email</th>
            <th>Workbook</th>
            <th>Added By</th>
            <th>When Added</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

async function refreshAuditLog(options = {}) {
  const { silent = false } = options;
  if (!silent) {
    setAuditLogStatus('Loading audit log...', 'info');
  }

  try {
    const response = await sendWithTimeout({ type: 'AB_GET_AUDIT_LOG', limit: AUDIT_LOG_FETCH_LIMIT }, 60000);
    if (!response || !response.success) {
      throw new Error(response?.error || 'Audit log request failed');
    }

    auditLogEntries = Array.isArray(response.entries) ? response.entries : [];
    renderAuditLogEntries(auditLogEntries);

    const count = auditLogEntries.length;
    setAuditLogStatus(count ? `Showing ${count} audit entr${count === 1 ? 'y' : 'ies'}.` : 'No audit entries recorded yet.', count ? 'success' : 'info');
  } catch (error) {
    console.error('[Panel] Audit log fetch failed:', error);
    setAuditLogStatus('Failed to load audit log: ' + (error?.message || error), 'error');
  }
}

async function revealAuditTab() {
  const hiddenTabButton = document.querySelector('.tab.hidden[data-tab="audit"]');
  if (!hiddenTabButton) {
    console.warn('[Panel] Audit tab button not found');
    return;
  }

  if (typeof initTrackedEntriesUI === 'function') {
    try {
      await Promise.resolve(initTrackedEntriesUI());
      if (typeof renderTrackedEntriesUI === 'function') {
        renderTrackedEntriesUI();
      }
    } catch (error) {
      console.warn('[Panel] Failed to initialize tracked entries UI before reveal:', error);
    }
  }

  hiddenTabButton.click();
}

window.revealAuditTab = revealAuditTab;
window.__AB_DEBUG = window.__AB_DEBUG || {};
window.__AB_DEBUG.showAuditLog = revealAuditTab;
function showStatus(message, type = 'info') {



  console.log(`[Panel] [${type}] ${message}`);



  const statusText = $('statusText');



  if (statusText) {



    statusText.textContent = message;



  }



}







document.addEventListener('click', (e) => {



  if (!e.target.closest('.filter-wrapper')) {



    document.querySelectorAll('.filter-dropdown-menu.show').forEach(menu => {



      menu.classList.remove('show');



    });



  }



});







document.addEventListener('DOMContentLoaded', async () => {



  console.log('[Panel] DOM loaded, initializing...');







  getLoadingOverlayRefs();







  setupTabs();



  setupMoreMenu();



  setupConnectButton();



  setupDebugButtons();



  setupInviteTab();



  setupAddBidderTab();



  setupAddBidTab();

  setAuditLogStatus('Select the Audit tab to load data.', 'info');






  const loadProjectBtn = $('loadProjectBtn');



  if (loadProjectBtn && !loadProjectBtn.dataset.bound) {



    loadProjectBtn.dataset.bound = 'true';



    loadProjectBtn.addEventListener('click', async () => {



      const projectInput = $('bidderProjectInput');



      const projectName = projectInput?.value.trim();



      if (!projectName) {



        showStatus('Enter a project name to load', 'error');



        return;



      }







      const originalLabel = loadProjectBtn.textContent;



      loadProjectBtn.disabled = true;



      loadProjectBtn.textContent = 'Loading...';







      try {



        await loadProject(projectName);



      } catch (error) {



        console.warn('[Panel] Load project button error:', error);



      } finally {



        loadProjectBtn.disabled = false;



        loadProjectBtn.textContent = originalLabel || 'Load Project';



      }



    });



  }







  const initialLoad = await loadCachedDataOnStartup({



    showOverlay: true,



    requireData: true,



    maxAttempts: 20,



    reason: 'initial'



  });







  if (!initialLoad.success && initialLoad.error) {



    console.warn('[Panel] Initial cache load warning:', initialLoad.error);



  }







  if (!initialLoad.hasData) {



    showStatus('Connect to SharePoint to download company data.', 'info');



  }







  await checkConnection();







  if (isInIframe) {



    setTimeout(() => {



      window.parent.postMessage({ type: 'AB_REQUEST_EMAIL_DATA' }, '*');



    }, 1000);



  }







  console.log('[Panel] Panel initialized');

  autoLoadLastProject();



});















function setupAddBidderTab() {
  // Safe-guard: wire behaviors if fields exist. Avoid errors if tab not present.
  const nameEl = $('bidderName');
  const emailEl = $('bidderEmail');
  const phoneEl = $('bidderPhone');
  const projectInput = $('bidderProjectInput');

  // Prefill contact fields from latest email data
  try {
    if (typeof emailData !== 'undefined') {
      if (emailEl && !emailEl.value) emailEl.value = (emailData.sender || '').toLowerCase();
      if (nameEl && !nameEl.value) nameEl.value = (emailData.senderName || '').trim();
      if (phoneEl && !phoneEl.value && emailData.signature && emailData.signature.phone) {
        phoneEl.value = emailData.signature.phone;
      }
    }
  } catch (e) { console.warn('setupAddBidderTab prefill skipped:', e); }

  // Run broader email-to-form population (shared with Add Bid tab)
  try { populateContactFromEmail(); } catch (err) { console.warn('populateContactFromEmail failed in setupAddBidderTab:', err); }

  // Auto-detect project from subject if we have data and the field is empty
  try {
    if (emailData?.subject && projectInput && !projectInput.value && Array.isArray(allProjects) && allProjects.length > 0) {
      autoDetectProject(emailData.subject);
    }
  } catch (err) { console.warn('autoDetectProject in setupAddBidderTab failed:', err); }

  // Wire the Add Bidder button
  const submitBtn = $('submitBidderBtn');
  if (submitBtn && !submitBtn.dataset.bound) {
    submitBtn.dataset.bound = 'true';
    submitBtn.addEventListener('click', (event) => {
      event?.preventDefault?.();
      submitBidder().catch(error => {
        console.error('[Panel] submitBidder failed:', error);
        showStatus('Failed to add bidder: ' + (error?.message || error), 'error');
      });
    });
  }
}

function setupAddBidTab() {



  // Minimal safe initialization for Add Bid tab.



  const cEl = $('bidContact');



  const eEl = $('bidEmail');



  const pEl = $('bidPhone');



  try {



    if (typeof emailData !== 'undefined') {



      if (eEl && !eEl.value) eEl.value = (emailData.sender || '').toLowerCase();



      if (cEl && !cEl.value) cEl.value = (emailData.senderName || '').trim();



      if (pEl && !pEl.value && emailData.signature && emailData.signature.phone) {



        pEl.value = emailData.signature.phone;



      }



    }



  } catch (e) { console.warn('setupAddBidTab prefill skipped:', e); }



}



















function populateContactFromEmail() {



  try {



    if (typeof emailData === 'undefined') return;



    var name = (emailData.senderName || '').trim();



    var email = (emailData.sender || '').trim();



    var phone = (emailData.signature && emailData.signature.phone) ? emailData.signature.phone : '';







    var fields = [



      $('bidderName'), $('bidderEmail'), $('bidderPhone'),



      $('bidContact'), $('bidEmail'), $('bidPhone')



    ];







    if (fields[0] && !fields[0].value && name) fields[0].value = name;



    if (fields[1] && !fields[1].value && email) fields[1].value = email;



    if (fields[2] && !fields[2].value && phone) fields[2].value = phone;



    if (fields[3] && !fields[3].value && name) fields[3].value = name;



    if (fields[4] && !fields[4].value && email) fields[4].value = email;



    if (fields[5] && !fields[5].value && phone) fields[5].value = phone;



  } catch (e) {



    console.warn('populateContactFromEmail failed:', e);



  }



}



















function autoDetectProject(subject) {



  try {



    if (!subject || !Array.isArray(allProjects) || allProjects.length === 0) return;



    var sub = String(subject).toLowerCase();



    var match = null;



    for (var i = 0; i < allProjects.length; i++) {



      var p = allProjects[i];



      var projectDisplay = getProjectDisplayName(p).toLowerCase();



      if (projectDisplay && sub.indexOf(projectDisplay) !== -1) { match = p; break; }



    }



    if (!match) return;



    var display = getProjectDisplayName(match);



    if ($('bidderProjectInput')) $('bidderProjectInput').value = display || '';



    if ($('bidProject')) $('bidProject').value = display || '';



  } catch (e) { console.warn('autoDetectProject failed:', e); }



}























































































