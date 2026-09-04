// AB Estimating website bridge
// Limited bridge for the protected AB Estimating Cloudflare Pages site.
// It reuses the extension's existing SharePoint/Graph project scan logic without
// exposing general chrome.runtime messaging to arbitrary websites.

(() => {
  'use strict';

  const ALLOWED_ORIGIN = 'https://ab-estimating.pages.dev';
  const REQUEST_TYPE = 'AB_ESTIMATING_EXTENSION_REQUEST';
  const RESPONSE_TYPE = 'AB_ESTIMATING_EXTENSION_RESPONSE';
  const READY_TYPE = 'AB_ESTIMATING_EXTENSION_READY';
  const PROGRESS_TYPE = 'AB_ESTIMATING_SCAN_PROGRESS';

  const ALLOWED_MESSAGE_TYPES = new Set([
    'AB_AUTHENTICATE',
    'AB_GET_STATES',
    'AB_SCAN_STATE',
    'AB_GET_CACHED_DATA'
  ]);

  if (window.location.origin !== ALLOWED_ORIGIN) return;

  function postToPage(payload) {
    window.postMessage(payload, ALLOWED_ORIGIN);
  }

  window.addEventListener('message', async (event) => {
    if (event.source !== window || event.origin !== ALLOWED_ORIGIN) return;

    const data = event.data;
    if (!data || data.type !== REQUEST_TYPE) return;

    const requestId = typeof data.requestId === 'string' ? data.requestId : '';
    const message = data.message;
    const messageType = message && typeof message.type === 'string' ? message.type : '';

    if (!requestId || !ALLOWED_MESSAGE_TYPES.has(messageType)) {
      if (requestId) {
        postToPage({
          type: RESPONSE_TYPE,
          requestId,
          response: { success: false, error: 'Unsupported AB Estimating extension request.' }
        });
      }
      return;
    }

    try {
      const response = await chrome.runtime.sendMessage(message);
      postToPage({ type: RESPONSE_TYPE, requestId, response });
    } catch (error) {
      postToPage({
        type: RESPONSE_TYPE,
        requestId,
        response: { success: false, error: String(error?.message || error || 'Extension request failed.') }
      });
    }
  });

  chrome.runtime.onMessage.addListener((message) => {
    if (!message || message.type !== 'AB_SCAN_PROGRESS') return;
    postToPage({ type: PROGRESS_TYPE, progress: message.progress || null });
  });

  postToPage({ type: READY_TYPE });
})();
