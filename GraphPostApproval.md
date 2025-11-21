Post-Approval Graph Integration Steps
=====================================

1. Update background.js to use the new Azure AD application client ID:
   - Replace the current AZURE_CONFIG.clientId value ('a9be13b9-b750-4dd6-a28b-9bf0169c44bb') with the new app id ('ff9b3bfb-d466-4c67-864a-65e9860dd9f3').
   - Ensure tenantId remains 73403434-959f-42fd-990a-9b774abe1489.

2. Confirm the Azure app has a SPA redirect URI registered in Azure AD:
   - URI format: https://<extension-id>.chromiumapp.org/
   - <extension-id> is the Chrome extension ID shown in chrome://extensions once the unpacked build is loaded.

3. Retest Connect to SharePoint in panel.js:
   - Use the Connect button, complete Microsoft login, and verify cached data loads.
   - Run a single-state scan and confirm cache upload succeeds.

4. Optional hardening once the robot flow is live:
   - Consider storing the client secret outside the bundle (e.g., proxy service or build-time injection) if the app is ever shipped outside trusted machines.

File created 2025-10-29 to track the cutover steps.
