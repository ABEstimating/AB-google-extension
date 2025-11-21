// Contact Merging Helper Functions
// Add these functions to panel.js

/**
 * Normalize email for comparison
 */
function normalizeEmail(email) {
  return (email || '').toString().trim().toLowerCase();
}

/**
 * Extract domain from email
 */
function getEmailDomain(email) {
  const normalized = normalizeEmail(email);
  const atIndex = normalized.lastIndexOf('@');
  return atIndex > 0 ? normalized.substring(atIndex + 1) : '';
}

/**
 * Check if domain is generic (gmail, yahoo, etc)
 */
function isGenericDomain(domain) {
  const genericDomains = [
    'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 
    'aol.com', 'icloud.com', 'mail.com', 'protonmail.com',
    'live.com', 'msn.com', 'ymail.com', 'comcast.net'
  ];
  return genericDomains.includes(domain.toLowerCase().trim());
}

/**
 * Normalize company name for matching
 */
function normalizeCompanyName(name) {
  return (name || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Merge duplicate contacts by email
 * Combines divisions, keeps most complete data
 */
function mergeContactsByEmail(contacts) {
  const emailMap = new Map();
  
  contacts.forEach(contact => {
    const email = normalizeEmail(contact.email);
    
    if (!email) {
      // No email - treat as unique contact
      emailMap.set(`no-email-${Math.random()}`, [contact]);
      return;
    }
    
    if (!emailMap.has(email)) {
      emailMap.set(email, []);
    }
    emailMap.get(email).push(contact);
  });
  
  const mergedContacts = [];
  
  emailMap.forEach((contactGroup, email) => {
    if (contactGroup.length === 1) {
      mergedContacts.push(contactGroup[0]);
      return;
    }
    
    // Merge multiple contacts with same email
    const merged = { ...contactGroup[0] };
    
    // Collect all unique divisions
    const divisionsSet = new Set();
    const divisionKeysSet = new Set();
    const projectsSet = new Set();
    
    contactGroup.forEach(c => {
      if (c.division) divisionsSet.add(c.division);
      if (c.divisionKey) divisionKeysSet.add(c.divisionKey);
      if (c.project) projectsSet.add(c.project);
      
      // Use most complete data
      if (!merged.name && c.name) merged.name = c.name;
      if (!merged.phone && c.phone) merged.phone = c.phone;
      if (!merged.company && c.company) merged.company = c.company;
      if (!merged.city && c.city) merged.city = c.city;
      if (!merged.state && c.state) merged.state = c.state;
      if (!merged.zip && c.zip) merged.zip = c.zip;
    });
    
    // Store as arrays for multiple divisions/projects
    merged.divisions = Array.from(divisionsSet);
    merged.divisionKeys = Array.from(divisionKeysSet);
    merged.projects = Array.from(projectsSet);
    merged.occurrenceCount = contactGroup.length;
    
    mergedContacts.push(merged);
  });
  
  return mergedContacts;
}

/**
 * Group contacts by company
 * Returns Map of companyKey -> {company, domain, contacts[]}
 */
function groupContactsByCompany(contacts) {
  const companyMap = new Map();
  
  contacts.forEach(contact => {
    const email = normalizeEmail(contact.email);
    const domain = getEmailDomain(email);
    const companyName = normalizeCompanyName(contact.company);
    
    let companyKey;
    
    if (!domain || isGenericDomain(domain)) {
      // Generic email - use company name as key
      companyKey = `generic-${companyName}`;
    } else {
      // Non-generic - use domain as key
      companyKey = `domain-${domain}`;
    }
    
    if (!companyMap.has(companyKey)) {
      companyMap.set(companyKey, {
        company: contact.company, // Use first occurrence
        companyNormalized: companyName,
        domain: domain,
        contacts: []
      });
    }
    
    companyMap.get(companyKey).contacts.push(contact);
  });
  
  return companyMap;
}

/**
 * Determine main contact for a company
 * Main contact = most occurrences in workbooks (highest occurrenceCount)
 */
function determineMainContact(contacts) {
  if (!contacts || contacts.length === 0) return null;
  if (contacts.length === 1) return contacts[0];
  
  // Sort by occurrence count (highest first)
  const sorted = [...contacts].sort((a, b) => {
    const countA = a.occurrenceCount || 1;
    const countB = b.occurrenceCount || 1;
    if (countB !== countA) return countB - countA;
    
    // Tie-breaker: prefer contacts with more complete info
    const scoreA = (a.name ? 1 : 0) + (a.phone ? 1 : 0) + (a.email ? 1 : 0);
    const scoreB = (b.name ? 1 : 0) + (b.phone ? 1 : 0) + (b.email ? 1 : 0);
    if (scoreB !== scoreA) return scoreB - scoreA;
    
    // Final tie-breaker: alphabetical by email
    return normalizeEmail(a.email).localeCompare(normalizeEmail(b.email));
  });
  
  return sorted[0];
}

/**
 * Prepare companies for display
 * Returns array of {company, mainContact, additionalContacts[], allDivisions[], isExpanded}
 */
function prepareCompaniesForDisplay(contacts) {
  // Step 1: Merge contacts by email
  const mergedContacts = mergeContactsByEmail(contacts);
  
  // Step 2: Group by company
  const companyMap = groupContactsByCompany(mergedContacts);
  
  // Step 3: Prepare display structure
  const companies = [];
  
  companyMap.forEach((companyData, companyKey) => {
    const mainContact = determineMainContact(companyData.contacts);
    const additionalContacts = companyData.contacts.filter(c => c !== mainContact);
    
    // Collect all unique divisions for the company
    const allDivisionsSet = new Set();
    companyData.contacts.forEach(contact => {
      if (contact.divisions && Array.isArray(contact.divisions)) {
        contact.divisions.forEach(d => allDivisionsSet.add(d));
      } else if (contact.division) {
        allDivisionsSet.add(contact.division);
      }
    });
    
    companies.push({
      companyKey: companyKey,
      company: companyData.company,
      domain: companyData.domain,
      mainContact: mainContact,
      additionalContacts: additionalContacts,
      allDivisions: Array.from(allDivisionsSet).sort(),
      allContacts: companyData.contacts,
      isExpanded: false,
      totalContacts: companyData.contacts.length
    });
  });
  
  // Sort companies alphabetically
  companies.sort((a, b) => {
    const nameA = (a.company || '').toLowerCase();
    const nameB = (b.company || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });
  
  return companies;
}
