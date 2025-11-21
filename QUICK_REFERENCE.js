/**
 * QUICK REFERENCE CARD
 * Main Functions for Merged Contacts Feature
 */

// ============================================================================
// CORE WORKFLOW
// ============================================================================

// 1. Prepare companies from raw contacts
const companies = prepareCompaniesForDisplay(contactsArray);

// 2. Render companies in UI
renderCompanies(companies, containerElement);

// 3. Get selected contacts
const selections = getSelectedContactsFromCompanies(containerElement);

// 4. Map selections back to contact objects
const selectedContacts = selections.map(ref => {
  const company = companies[ref.companyIndex];
  return ref.contactIndex === 0 
    ? company.mainContact 
    : company.additionalContacts[ref.contactIndex - 1];
}).filter(Boolean);


// ============================================================================
// FUNCTION SIGNATURES
// ============================================================================

/**
 * Prepare contacts for display - does all merging and grouping
 * @param {Array} contacts - Raw contact array
 * @returns {Array} Array of company objects ready for display
 */
prepareCompaniesForDisplay(contacts)

/**
 * Render companies in the UI
 * @param {Array} companies - Array from prepareCompaniesForDisplay()
 * @param {HTMLElement} containerElement - DOM element to render into
 */
renderCompanies(companies, containerElement)

/**
 * Get selected contacts
 * @param {HTMLElement} containerElement - Same element used in renderCompanies()
 * @returns {Array} Array of {companyKey, companyIndex, contactIndex, email}
 */
getSelectedContactsFromCompanies(containerElement)

/**
 * Toggle company expansion (called automatically by click)
 * @param {HTMLElement} companyRowElement - The company row DOM element
 * @param {Object} companyData - The company data object
 */
toggleCompanyExpansion(companyRowElement, companyData)


// ============================================================================
// DATA STRUCTURES
// ============================================================================

// Company Object (output of prepareCompaniesForDisplay)
{
  companyKey: "domain-company.com",           // Unique key
  company: "ABC Construction",                 // Display name
  domain: "company.com",                       // Email domain
  mainContact: {                              // Primary contact
    email: "john@company.com",
    name: "John Doe",
    phone: "555-1234",
    divisions: ["Concrete", "Masonry"],
    occurrenceCount: 15                       // Times seen in workbooks
  },
  additionalContacts: [                       // Other contacts
    {
      email: "jane@company.com",
      name: "Jane Smith",
      divisions: ["Concrete"],
      occurrenceCount: 8
    }
  ],
  allDivisions: ["Concrete", "Masonry"],     // All unique divisions
  allContacts: [...],                         // All contacts combined
  isExpanded: false,                          // UI state
  totalContacts: 2                            // Total count
}

// Selection Object (output of getSelectedContactsFromCompanies)
{
  companyKey: "domain-company.com",
  companyIndex: 42,                           // Index in companies array
  contactIndex: 0,                            // 0 = main, 1+ = additional
  email: "john@company.com"
}


// ============================================================================
// HELPER FUNCTIONS (also available if needed)
// ============================================================================

normalizeEmail(email)                         // "John@Company.COM" → "john@company.com"
getEmailDomain(email)                         // "john@company.com" → "company.com"
isGenericDomain(domain)                       // "gmail.com" → true
normalizeCompanyName(name)                    // "ABC Corp." → "abc corp"
mergeContactsByEmail(contacts)                // Combines duplicates
groupContactsByCompany(contacts)              // Groups by company
determineMainContact(contacts)                // Picks primary contact


// ============================================================================
// USAGE EXAMPLES
// ============================================================================

// Example 1: Display contacts in a container
const myContacts = [...]; // Your contacts array
const companies = prepareCompaniesForDisplay(myContacts);
const container = document.getElementById('my-container');
renderCompanies(companies, container);

// Example 2: Get selected contacts
const selections = getSelectedContactsFromCompanies(container);
const selectedContacts = selections.map(ref => {
  const company = companies[ref.companyIndex];
  if (!company) return null;
  return ref.contactIndex === 0 
    ? company.mainContact 
    : company.additionalContacts[ref.contactIndex - 1];
}).filter(Boolean);

console.log('Selected:', selectedContacts);

// Example 3: Select all contacts programmatically
companies.forEach((company, idx) => {
  const row = container.querySelector(`[data-company-index="${idx}"]`);
  if (row) {
    const checkboxes = row.querySelectorAll('.contact-checkbox');
    checkboxes.forEach(cb => cb.checked = true);
  }
});

// Example 4: Filter companies before rendering
const companiesWithMultipleContacts = companies.filter(c => c.totalContacts > 1);
renderCompanies(companiesWithMultipleContacts, container);

// Example 5: Expand all companies programmatically
companies.forEach((company, idx) => {
  const row = container.querySelector(`[data-company-index="${idx}"]`);
  if (row && company.totalContacts > 1) {
    const additionalContainer = row.querySelector('.additional-contacts-container');
    if (additionalContainer) {
      additionalContainer.style.display = 'block';
      row.querySelector('.expand-icon').innerHTML = '▼';
      row.classList.add('expanded');
    }
  }
});


// ============================================================================
// CSS CLASSES (for custom styling)
// ============================================================================

.company-row                                  // Outer container
.company-main-row                             // Clickable row
.expand-icon                                  // ▶/▼ icon
.contact-checkbox                             // All checkboxes
.main-contact-checkbox                        // Main contact checkbox
.additional-contact-checkbox                  // Additional contact checkboxes
.company-info                                 // Company info grid
.company-field                                // Each field
.field-label                                  // Field labels
.field-value                                  // Field values
.additional-contacts-container                // Expandable section
.select-buttons                               // Button container
.select-all-company                           // Select all button
.clear-all-company                            // Clear all button
.additional-contact-row                       // Each additional contact
.additional-contact-info                      // Additional contact fields
.expanded                                     // Class added when expanded


// ============================================================================
// DEBUGGING
// ============================================================================

// Check prepared companies
const companies = prepareCompaniesForDisplay(contacts);
console.log('Total companies:', companies.length);
console.log('Sample:', companies[0]);
console.log('Multi-contact companies:', companies.filter(c => c.totalContacts > 1).length);

// Check rendering
renderCompanies(companies, container);
console.log('Rendered elements:', container.querySelectorAll('.company-row').length);

// Check selections
const selections = getSelectedContactsFromCompanies(container);
console.log('Selected count:', selections.length);
console.log('Selections:', selections);

// Verify merge
const originalCount = contacts.length;
const merged = mergeContactsByEmail(contacts);
console.log('Merged', originalCount - merged.length, 'duplicates');
console.log('Result:', merged.length, 'unique contacts');
