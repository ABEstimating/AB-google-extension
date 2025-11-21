// Optional bundled cache for completed projects.
// Replace the placeholder arrays with your frozen project data.
// Each project entry should match the structure returned by the service worker's cachedProjects values.
// Contacts should use the same shape as cachedContacts entries.
// Example shapes:
// projects: [
//   {
//     projectKey: '25073kiaofboerne',
//     displayName: '25-073 Kia of Boerne',
//     divisions: [],
//     contactsByDivision: {},
//     // ...any additional fields you rely on in the UI
//   }
// ],
// contacts: [
//   {
//     projectKey: '25073kiaofboerne',
//     project: '25-073 Kia of Boerne',
//     division: 'Electrical',
//     company: 'Sample Electric',
//     email: 'estimator@example.com'
//   }
// ]
self.STATIC_PROJECT_CACHE = Object.freeze({
  generatedAt: null,
  projects: [],
  contacts: [],
  zipCache: {},
  activeProjectKeys: []
});
