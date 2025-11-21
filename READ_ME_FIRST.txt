Campaign Tab Follow-Up (next session)
===================================

Context / Current Issues
------------------------
- The campaigns worker (campaigns-bg.js) currently lacks top-level normalizeSchedule/normalizeFollowUps helpers. The restored copy still has normalizeFollowUps defined inside normalizeSchedule, so AB_CMP_CREATE crashes with "normalizeFollowUps is not defined". First task is to reintroduce standalone normalizeSchedule + normalizeFollowUps functions and reapply the message guard so only AB_CAMPAIGN / AB_CMP / campaign:* messages stay in the worker. Everything else must fall through to background.js.
- The HTML for the contact search filters is still plain <input> fields. We need chip-based multi-select/typeahead components for City, State, Project, and Division. These should pull from allContacts (city/state data), allProjects (project names matching the Audit tab’s “Select Project to Scan” list), and division_mapping_helper.js (divisions sorted ascending). runContactSearch() currently expects single strings; it must be refactored to accept arrays of selected values and match any selection.
- There’s no per-recipient delivery log. We only store lastSentAt at the campaign level, and remove recipients once they respond. To display the requested summary (counts of sent/pending/responses plus detailed table), we need to extend runSendCycle/sendEmails to log delivery timestamps per recipient, keep pending recipients with metadata, persist that data to storage, and expose it to the UI so the campaign tab can render a summary table whenever a campaign loads.
- The “Review & Send” modal’s “Send Selected” button calls AB_CMP_SEND_SELECTED, which bypasses scheduling and sends immediately. Fixing the “should wait until later in the day” issue requires redesigning that flow (e.g., disable/send through scheduler when future send times exist) and ensuring applySchedule/reschedule logic re-arms alarms after manual sends. The current worker also fails to renormalize follow-ups after manual sends, which is why RFI/Bid reminders never triggered.

Next-Session Checklist
----------------------
1. Update campaigns-bg.js
   - Restore/define normalizeSchedule + normalizeFollowUps at top level.
   - Gate chrome.runtime.onMessage so only campaign messages are handled; wrap sendResponse via safeRespond.
   - After fixes, run node --check campaigns-bg.js to ensure syntax is valid.
   - Verify follow-up schedules persist after manual sends.

2. Multi-select Search Filters
   - Replace City/State/Project/Division inputs in panel.html with chip-based multi-select components.
   - Add required CSS (panel.css) and JS plumbing (campaigns-ui.js) to populate dropdowns from allContacts/allProjects/division mapping.
   - Update runContactSearch to accept multiple selected values per field and adjust how filters are applied.

3. Campaign Summary
   - Extend campaigns-bg.js to track per-recipient send metadata (delivery timestamps, status) and keep recipients until they respond.
   - Expose summary info in campaign payloads sent to the UI.
   - Update campaigns-ui.js to render a table showing sent/pending counts and per-recipient status when a campaign loads.

4. Scheduler / Follow-Ups
   - Ensure scheduled sends respect future nextSendAt values (no immediate sends after review unless intended).
   - Re-test follow-up types (RFI/Bid reminders) after manual sends; confirm applySchedule/reschedule re-queues follow-ups correctly.

Completion Criteria
-------------------
- campaigns-bg.js loads without runtime errors; AB_CMP_CREATE works.
- City/State/Project/Division filters support multi-select with typeahead dropdowns populated from live data.
- Selecting a campaign shows a summary (counts + table) of sent/pending/responses.
- Manual sends do not bypass scheduling when a future time is set; follow-up emails send at their configured times.
