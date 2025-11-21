// New Rendering Function for Merged Companies
// Replace the existing renderContacts function in panel.js

/**
 * Render companies with expandable contact rows
 */
function renderCompanies(companies, containerElement) {
  if (!containerElement) {
    console.error('[renderCompanies] No container element provided');
    return;
  }
  
  containerElement.innerHTML = '';
  
  if (!companies || companies.length === 0) {
    containerElement.innerHTML = '<div class="no-results">No contacts found matching your criteria.</div>';
    return;
  }
  
  console.log(`[renderCompanies] Rendering ${companies.length} companies`);
  
  companies.forEach((companyData, companyIndex) => {
    const { company, mainContact, additionalContacts, allDivisions, totalContacts, companyKey } = companyData;
    
    // Create company row container
    const companyRow = document.createElement('div');
    companyRow.className = 'company-row';
    companyRow.dataset.companyKey = companyKey;
    companyRow.dataset.companyIndex = companyIndex;
    
    // Main company row (collapsed view)
    const mainRow = document.createElement('div');
    mainRow.className = 'company-main-row';
    
    // Expand/collapse icon
    const expandIcon = document.createElement('span');
    expandIcon.className = 'expand-icon';
    expandIcon.innerHTML = totalContacts > 1 ? '▶' : ''; // Only show if multiple contacts
    expandIcon.style.cursor = totalContacts > 1 ? 'pointer' : 'default';
    
    // Main contact checkbox
    const mainCheckbox = document.createElement('input');
    mainCheckbox.type = 'checkbox';
    mainCheckbox.className = 'contact-checkbox main-contact-checkbox';
    mainCheckbox.dataset.companyKey = companyKey;
    mainCheckbox.dataset.contactIndex = '0'; // Main contact is index 0
    mainCheckbox.dataset.email = mainContact?.email || '';
    
    // Company info
    const companyInfo = document.createElement('div');
    companyInfo.className = 'company-info';
    companyInfo.innerHTML = `
      <div class="company-field company-name">
        <span class="field-label">Company:</span>
        <span class="field-value">${escapeHtml(company || 'Unknown')}</span>
        ${totalContacts > 1 ? `<span class="contact-count">(${totalContacts} contacts)</span>` : ''}
      </div>
      <div class="company-field contact-name">
        <span class="field-label">Contact:</span>
        <span class="field-value">${escapeHtml(mainContact?.name || '-')}</span>
      </div>
      <div class="company-field contact-email">
        <span class="field-label">Email:</span>
        <span class="field-value">${escapeHtml(mainContact?.email || '-')}</span>
      </div>
      <div class="company-field contact-phone">
        <span class="field-label">Phone:</span>
        <span class="field-value">${escapeHtml(mainContact?.phone || mainContact?.phoneDigits || '-')}</span>
      </div>
      <div class="company-field contact-location">
        <span class="field-label">Location:</span>
        <span class="field-value">${escapeHtml(mainContact?.city || '')}, ${escapeHtml(mainContact?.state || '')}</span>
      </div>
      <div class="company-field company-divisions">
        <span class="field-label">Divisions:</span>
        <span class="field-value divisions-list">${allDivisions.map(d => escapeHtml(d)).join(', ')}</span>
      </div>
    `;
    
    // Append elements to main row
    mainRow.appendChild(expandIcon);
    mainRow.appendChild(mainCheckbox);
    mainRow.appendChild(companyInfo);
    
    // Add click handler for expanding (only if multiple contacts)
    if (totalContacts > 1) {
      mainRow.style.cursor = 'pointer';
      mainRow.addEventListener('click', (e) => {
        // Don't expand if clicking checkbox
        if (e.target.type === 'checkbox') return;
        
        toggleCompanyExpansion(companyRow, companyData);
      });
    }
    
    companyRow.appendChild(mainRow);
    
    // Additional contacts container (initially hidden)
    if (totalContacts > 1) {
      const additionalContainer = document.createElement('div');
      additionalContainer.className = 'additional-contacts-container';
      additionalContainer.style.display = 'none';
      
      // Select All / Clear All buttons
      const selectButtons = document.createElement('div');
      selectButtons.className = 'select-buttons';
      selectButtons.innerHTML = `
        <button class="select-all-company" data-company-key="${companyKey}">Select All</button>
        <button class="clear-all-company" data-company-key="${companyKey}">Clear All</button>
      `;
      
      additionalContainer.appendChild(selectButtons);
      
      // Render additional contacts
      additionalContacts.forEach((contact, contactIndex) => {
        const contactRow = document.createElement('div');
        contactRow.className = 'additional-contact-row';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'contact-checkbox additional-contact-checkbox';
        checkbox.dataset.companyKey = companyKey;
        checkbox.dataset.contactIndex = contactIndex + 1; // +1 because 0 is main contact
        checkbox.dataset.email = contact?.email || '';
        
        const contactInfo = document.createElement('div');
        contactInfo.className = 'additional-contact-info';
        contactInfo.innerHTML = `
          <div class="contact-field">
            <span class="field-label">Name:</span>
            <span class="field-value">${escapeHtml(contact?.name || '-')}</span>
          </div>
          <div class="contact-field">
            <span class="field-label">Email:</span>
            <span class="field-value">${escapeHtml(contact?.email || '-')}</span>
          </div>
          <div class="contact-field">
            <span class="field-label">Phone:</span>
            <span class="field-value">${escapeHtml(contact?.phone || contact?.phoneDigits || '-')}</span>
          </div>
          <div class="contact-field">
            <span class="field-label">Divisions:</span>
            <span class="field-value">${(contact.divisions || [contact.division]).filter(Boolean).map(d => escapeHtml(d)).join(', ')}</span>
          </div>
        `;
        
        contactRow.appendChild(checkbox);
        contactRow.appendChild(contactInfo);
        additionalContainer.appendChild(contactRow);
      });
      
      companyRow.appendChild(additionalContainer);
      
      // Add event listeners for Select All / Clear All
      selectButtons.querySelector('.select-all-company').addEventListener('click', () => {
        const checkboxes = companyRow.querySelectorAll('.contact-checkbox');
        checkboxes.forEach(cb => cb.checked = true);
      });
      
      selectButtons.querySelector('.clear-all-company').addEventListener('click', () => {
        const checkboxes = companyRow.querySelectorAll('.contact-checkbox');
        checkboxes.forEach(cb => cb.checked = false);
      });
    }
    
    containerElement.appendChild(companyRow);
  });
  
  console.log('[renderCompanies] Rendering complete');
}

/**
 * Toggle company expansion
 */
function toggleCompanyExpansion(companyRowElement, companyData) {
  const additionalContainer = companyRowElement.querySelector('.additional-contacts-container');
  const expandIcon = companyRowElement.querySelector('.expand-icon');
  
  if (!additionalContainer) return;
  
  const isCurrentlyExpanded = additionalContainer.style.display !== 'none';
  
  if (isCurrentlyExpanded) {
    // Collapse
    additionalContainer.style.display = 'none';
    expandIcon.innerHTML = '▶';
    companyRowElement.classList.remove('expanded');
  } else {
    // Expand
    additionalContainer.style.display = 'block';
    expandIcon.innerHTML = '▼';
    companyRowElement.classList.add('expanded');
  }
}

/**
 * Get selected contacts from rendered companies
 */
function getSelectedContactsFromCompanies(containerElement) {
  const selectedContacts = [];
  const checkboxes = containerElement.querySelectorAll('.contact-checkbox:checked');
  
  checkboxes.forEach(checkbox => {
    const companyKey = checkbox.dataset.companyKey;
    const email = checkbox.dataset.email;
    
    // Find the company row
    const companyRow = checkbox.closest('.company-row');
    if (companyRow) {
      const companyIndex = parseInt(companyRow.dataset.companyIndex);
      const contactIndex = parseInt(checkbox.dataset.contactIndex);
      
      selectedContacts.push({
        companyKey,
        companyIndex,
        contactIndex,
        email
      });
    }
  });
  
  return selectedContacts;
}

// Helper function to escape HTML
function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  const div = document.createElement('div');
  div.textContent = String(text);
  return div.innerHTML;
}
