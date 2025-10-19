document.addEventListener('DOMContentLoaded', function() {
    const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzp7EWEWddYYqQMfBRcA2eriHkbbvCllS-9IOo-panx2Ve04oJThBxUINCuSN_7DwEzFg/exec';
    let lastSavedCaseKey = null; // To track the last saved form for Sheets sync

    // ===== Load or initialize stored data =====
    const defaults = {
        investigators: ['Conrad', 'Delores', 'Lisa', 'Michael', 'Jeff'],
        surveillanceInvestigators: ['Freddy', 'Ricardo', 'Delores', 'Michael', 'Alma', 'Manuel', 'Jeff'],
        clients: [],
        attorneys: [],
        employers: [],
        claimants: [],
        pipelineStages: [],
        states: ["CA", "NY", "TX", "FL"],
        currentForm: {},
        caseIdCounter: 1,
    };

    let localData = {};

    function loadLocalData() {
        const stored = JSON.parse(localStorage.getItem('intakeData') || '{}');
        localData = { ...defaults, ...stored };
        for (const key in defaults) {
            if (Array.isArray(defaults[key]) && (!localData[key] || !Array.isArray(localData[key]))) {
                localData[key] = defaults[key];
            } else if (typeof defaults[key] === 'object' && defaults[key] !== null && !Array.isArray(defaults[key]) && (!localData[key] || typeof localData[key] !== 'object' || Array.isArray(localData[key]))) {
                localData[key] = defaults[key];
            }
        }
        localData.currentForm = { ...defaults.currentForm, ...(stored.currentForm || {}) };
    }

    loadLocalData();

    // New success message function
    function showSuccessMessage(message) {
        const successDiv = document.createElement('div');
        successDiv.textContent = message;
        Object.assign(successDiv.style, {
            position: 'fixed',
            top: '20px',
            right: '20px',
            backgroundColor: '#10b981',
            color: 'white',
            padding: '16px 24px',
            borderRadius: '8px',
            fontWeight: '600',
            zIndex: '10000',
            boxShadow: '0 4px 6px rgba(0,0,0,0.1)',
                      opacity: '0',
                      transition: 'opacity 0.5s ease-in-out'
        });
        document.body.appendChild(successDiv);
        setTimeout(() => successDiv.style.opacity = '1', 10);
        setTimeout(() => {
            successDiv.style.opacity = '0';
            successDiv.addEventListener('transitionend', () => successDiv.remove());
        }, 3000);
    }


    const namedEntitySections = ['client', 'attorney', 'employer', 'claimant'];
    namedEntitySections.forEach(setupNamedEntitySection);
    setupPipelineStage();
    setupInvestigatorDropdown();
    setupStateField();
    setupSsnFormatting();

    document.querySelectorAll('input, select, textarea').forEach(field => {
        if (localData.currentForm[field.id] !== undefined) field.value = localData.currentForm[field.id];
        field.addEventListener('blur', () => saveCurrentFormState(field));
        if (field.tagName === 'SELECT') field.addEventListener('change', () => saveCurrentFormState(field));
        if (field.hasAttribute('list')) {
            field.addEventListener('input', () => {
                const datalist = document.getElementById(field.getAttribute('list'));
                const options = Array.from(datalist.options).map(o => o.value);
                // Clear dependent fields only if the input doesn't match an existing option AND is not empty
                if (!options.includes(field.value) && field.value !== '') {
                    // Do nothing, let user type
                } else if (field.value === '') {
                    clearDependentFields(field.id);
                }
            });
            // On change (when an option is selected or focus leaves), ensure blur is triggered
            field.addEventListener('change', () => field.blur());
        }
    });

    document.getElementById('save-btn').addEventListener('click', saveFullForm);
    document.getElementById('sync-btn').addEventListener('click', syncToSheets);

    function setupNamedEntitySection(section) {
        const nameField = document.getElementById(section + 'Name');
        const companyField = document.getElementById(section + 'Company');
        const listKey = section + 's';

        if (nameField) {
            populateDatalist(section + 'NameList', localData[listKey].map(i => i.name).filter(Boolean));
            nameField.addEventListener('change', () => autofillSection(section, nameField.value, 'name'));
            nameField.addEventListener('blur', () => saveNamedEntity(section)); // Removed triggerType, will collect all
        }
        if (companyField) {
            populateDatalist(section + 'CompanyList', localData[listKey].map(i => i.company).filter(Boolean));
            companyField.addEventListener('change', () => autofillSection(section, companyField.value, 'company'));
            companyField.addEventListener('blur', () => saveNamedEntity(section)); // Removed triggerType, will collect all
        }
        // Add event listeners for all fields within the section to save on blur/change
        document.querySelectorAll(`[data-section="${section}"] input, [data-section="${section}"] select, [data-section="${section}"] textarea`)
        .forEach(f => {
            // Initial load of currentForm values
            if (localData.currentForm[f.id] !== undefined) f.value = localData.currentForm[f.id];

            f.addEventListener('blur', () => saveCurrentFormState(f)); // Save individual field state
            if (f.tagName === 'SELECT') f.addEventListener('change', () => saveCurrentFormState(f)); // Save individual field state

            // IMPORTANT: When a field in a named entity section changes, we should attempt to save the whole entity.
            // This ensures that if a user changes an address for an existing client, it updates the client object.
            f.addEventListener('blur', () => saveNamedEntity(section));
            if (f.tagName === 'SELECT') f.addEventListener('change', () => saveNamedEntity(section));
        });
    }

    /**
     * Autofills fields for a specific section based on a matching value (name or company).
     * This function is crucial for data isolation.
     * @param {string} section - The section identifier (e.g., 'client', 'employer', 'claimant').
     * @param {string} value - The value to search for (e.g., a client name, employer company).
     * @param {string} triggerType - 'name' or 'company' to specify which field triggered the autofill.
     */
    function autofillSection(section, value, triggerType) {
        if (!value) return clearNamedEntitySection(section);

        const listKey = section + 's';
        const list = localData[listKey] || [];
        // CRITICAL FIX: Ensure we only search within the specific section's list.
        let record = triggerType === 'name' ? list.find(r => r.name === value) : list.find(r => r.company === value);

        if (record) {
            // Fill all fields for this section from the found record
            document.querySelectorAll(`[data-section="${section}"] input, [data-section="${section}"] select, [data-section="${section}"] textarea`)
            .forEach(f => {
                let k = f.id.replace(section, '');
                if (k) {
                    k = k.charAt(0).toLowerCase() + k.slice(1);
                    if (record[k] !== undefined) { // Only update if the record has this specific key
                        f.value = record[k] || '';
                        saveCurrentFormState(f); // Save the individual field's state to currentForm
                    }
                }
            });
        } else {
            // If no record found, clear all fields in the section except the one being typed into
            clearNamedEntitySectionExceptTrigger(section, triggerType);
        }
    }

    function clearNamedEntitySectionExceptTrigger(section, triggerType) {
        const fields = document.querySelectorAll(`[data-section="${section}"] input,[data-section="${section}"] select,[data-section="${section}"] textarea`);
        fields.forEach(f => {
            // CRITICAL FIX: Ensure the ID check for the trigger field is correct
            const triggerFieldId = section + capitalize(triggerType);
            if (f.id !== triggerFieldId) {
                if (f.tagName === 'SELECT') f.selectedIndex = 0;
                else if (f.type === 'date') f.value = '';
                else if (f.id.includes('State')) f.value = 'CA'; // Default state for state fields
                else f.value = '';
                saveCurrentFormState(f);
            }
        });
    }

    function clearNamedEntitySection(section) {
        const fields = document.querySelectorAll(`[data-section="${section}"] input,[data-section="${section}"] select,[data-section="${section}"] textarea`);
        fields.forEach(f => {
            if (f.tagName === 'SELECT') f.selectedIndex = 0;
            else if (f.type === 'date') f.value = '';
            else if (f.id.includes('State')) f.value = 'CA';
            else f.value = '';
            saveCurrentFormState(f);
        });
    }

    // This function is intended to clear a *named entity section* when its primary autofill field (e.g., clientName) is cleared.
    function clearDependentFields(fieldId) {
        // Find which section the fieldId belongs to
        const section = namedEntitySections.find(s => fieldId.startsWith(s));
        // If it's a primary named entity field (Name or Company) and it was cleared, clear the whole section.
        if (section && (fieldId.endsWith('Name') || fieldId.endsWith('Company'))) {
            const currentFieldValue = document.getElementById(fieldId).value;
            if (currentFieldValue === '') {
                clearNamedEntitySection(section);
            }
        }
    }

    /**
     * Saves or updates a named entity (client, attorney, employer, claimant) to localData.
     * This function is crucial for data isolation.
     * @param {string} section - The section identifier (e.g., 'client', 'employer', 'claimant').
     */
    function saveNamedEntity(section) {
        const record = collectNamedEntityValues(section);
        const listKey = section + 's';
        let list = localData[listKey] || [];

        // Check for empty records to avoid saving blank entries
        // For client, attorney, employer: require either name or company
        if (['client', 'attorney', 'employer'].includes(section) && !record.name && !record.company) return;
        // For claimant: require at least one identifying field
        if (section === 'claimant' && !record.name && !record.jobTitle && !record.phone && !record.email && !record.ssn && !record.dob && !record.street && !record.city && !record.state && !record.zip) return;

        let existingIndex = -1;
        if (record.name) {
            existingIndex = list.findIndex(r => r.name === record.name);
        } else if (record.company && section !== 'claimant') { // Claimants don't usually have a 'company' in this context for unique identification
            existingIndex = list.findIndex(r => r.company === record.company);
        }

        if (existingIndex !== -1) {
            // Update existing record by merging new values
            list[existingIndex] = { ...list[existingIndex], ...record };
        } else {
            // Add new record
            list.push(record);
        }
        localData[listKey] = list; // CRITICAL FIX: Assign the updated list back to localData for the specific section.
        localStorage.setItem('intakeData', JSON.stringify(localData));

        // Update datalists for the current section
        const nameField = document.getElementById(section + 'Name');
        const companyField = document.getElementById(section + 'Company');
        if (nameField) populateDatalist(section + 'NameList', localData[listKey].map(i => i.name).filter(Boolean));
        if (companyField) populateDatalist(section + 'CompanyList', localData[listKey].map(i => i.company).filter(Boolean));
    }

    function saveCurrentFormState(field) {
        if (field.id) {
            localData.currentForm[field.id] = field.value;
            localStorage.setItem('intakeData', JSON.stringify(localData));
        }
    }

    /**
     * Collects all field values for a specific named entity section.
     * CRITICAL FIX: Uses data-section attribute to ensure isolation.
     * @param {string} section - The section identifier (e.g., 'client', 'employer').
     * @returns {object} An object containing all collected field-value pairs for the section.
     */
    function collectNamedEntityValues(section) {
        const r = {};
        // CRITICAL FIX: Use the data-section attribute to correctly target fields
        document.querySelectorAll(`[data-section="${section}"] input, [data-section="${section}"] select, [data-section="${section}"] textarea`)
        .forEach(f => {
            let k = f.id.replace(section, ''); // Remove section prefix to get the key (e.g., 'Name', 'Phone')
        if (k) {
            k = k.charAt(0).toLowerCase() + k.slice(1); // Convert to camelCase (e.g., 'name', 'phone')
        r[k] = f.value;
        }
        });
        return r;
    }

    function setupPipelineStage() {
        const f = document.getElementById('assignmentPipelineStage');
        if (f) { populateDatalist('pipelineStageList', localData.pipelineStages); f.addEventListener('blur', () => savePipelineStage(f.value)); }
    }

    function savePipelineStage(v) {
        if (!v || localData.pipelineStages.includes(v)) return;
        localData.pipelineStages.push(v);
        localStorage.setItem('intakeData', JSON.stringify(localData));
        populateDatalist('pipelineStageList', localData.pipelineStages);
    }

    function setupStateField() {
        const sFields = document.querySelectorAll('input[id$="State"]');
        sFields.forEach(f => {
            if (!f.value) f.value = 'CA'; // Default to CA if empty
            populateDatalist('stateList', localData.states);
            f.addEventListener('blur', () => saveState(f.value));
        });
    }

    function saveState(v) {
        if (!v || localData.states.includes(v)) return;
        localData.states.push(v);
        localStorage.setItem('intakeData', JSON.stringify(localData));
        populateDatalist('stateList', localData.states);
    }

    function populateDatalist(id, vals) {
        const dl = document.getElementById(id);
        if (dl) {
            dl.innerHTML = '';
            vals.filter(v => v).sort().forEach(v => {
                const o = document.createElement('option');
                o.value = v; dl.appendChild(o);
            });
        }
    }

    function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

    function formatSSN(i) {
        let v = i.value.replace(/\D/g, ''), f = '';
        if (v.length > 0) f += v.substring(0, 3);
        if (v.length > 3) f += '-' + v.substring(3, 5);
        if (v.length > 5) f += '-' + v.substring(5, 9);
        i.value = f; saveCurrentFormState(i);
    }
    window.formatSSN = formatSSN;
    function setupSsnFormatting() {
        const s = document.getElementById('claimantSSN');
        if (s) s.addEventListener('input', () => formatSSN(s));
    }

    function setupInvestigatorDropdown() {
        const sSel = document.getElementById('assignmentService');
        const iSel = document.getElementById('investigatorName');
        function update() {
            const service = sSel.value;
            const list = service === 'Surveillance' ? localData.surveillanceInvestigators : localData.investigators;
            const current = iSel.value; // Get current value before clearing
            iSel.innerHTML = '';
            const def = document.createElement('option');
            def.value = ''; def.textContent = '-- Select Investigator --';
            iSel.appendChild(def);
            list.forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; iSel.appendChild(o); });
            // Set value back, preferring currentForm, then trying to match current value, otherwise default
            iSel.value = localData.currentForm[iSel.id] || (list.includes(current) ? current : '');
            saveCurrentFormState(iSel);
        }
        sSel.addEventListener('change', update);
        update(); // Call once on load to set initial state
    }

    function saveFullForm() {
        const formData = {};
        document.querySelectorAll('input, select, textarea').forEach(f => { if (f.id) formData[f.id] = f.value; });
        if (!localData.currentForm.caseId) {
            localData.currentForm.caseId = 'case_' + localData.caseIdCounter++;
            localStorage.setItem('intakeData', JSON.stringify(localData));
        }
        formData.caseId = localData.currentForm.caseId;

        // Normalize field names for compatibility
        formData.claimNumber = formData.assignmentClaimNumber || '';
        formData.status = formData.assignmentStatus || 'Open';
        formData.clientCompany = formData.clientCompany || ''; // This might be redundant if clientName is used
        formData.investigatorAssigned = formData.investigatorName || '';
        formData.dateCreated = new Date().toISOString();

        const cName = (formData.claimantName || '').trim().split(' ');
        formData.claimantFirst = cName[0] || '';
        formData.claimantLast = cName.slice(1).join(' ') || '';

        lastSavedCaseKey = formData.caseId;
        localStorage.setItem(lastSavedCaseKey, JSON.stringify(formData)); // Save as separate key

        document.getElementById('sync-btn').disabled = false;
        document.getElementById('sync-status').innerHTML = '<span style="color: var(--warning);">⚠️ Not synced to Sheets yet</span>';
        alert('Intake form saved to localStorage!\n\nClick "Sync to Sheets" to back up to Google Sheets.');

        // Update currentForm to reflect the just-saved state
        localData.currentForm = { ...formData };
        localStorage.setItem('intakeData', JSON.stringify(localData));
    }

    async function syncToSheets() {
        if (!localData.currentForm.caseId) return alert('Please save first.');
        const key = localData.currentForm.caseId;
        if (!localStorage.getItem(key)) return alert('No local data found for this case.');

        const btn = document.getElementById('sync-btn');
        const status = document.getElementById('sync-status');
        btn.disabled = true; btn.textContent = 'Syncing...';
        status.innerHTML = '<span style="color: var(--accent);">⏳ Syncing to Google Sheets...</span>';

        try {
            const caseData = JSON.parse(localStorage.getItem(key));

            // Send the caseData to Google Sheets
            await fetch(WEB_APP_URL, {
                method: 'POST',
                // mode: 'no-cors' is typically used for simple requests. If the script expects JSON,
                // it might need a 'cors' mode and proper CORS headers on the server side.
                // Assuming 'no-cors' works for Google Apps Script for now.
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ action: 'syncCase', caseData })
            });

            const currentIntakeData = JSON.parse(localStorage.getItem('intakeData')) || defaults;
            const caseIdNum = parseInt(caseData.caseId.replace('case_', '')) || 0; // Handle non-numeric caseId gracefully

            const normalizedCase = {
                ...caseData, // All original fields
                id: caseData.caseId, // Use existing caseId
                claimNumber: caseData.assignmentClaimNumber || '',
                status: caseData.assignmentStatus || 'Open',
                dateReceived: caseData.assignmentDateReceived || '',
                dateAssigned: caseData.assignmentDateAssigned || '',
                investigatorName: caseData.investigatorName || '',
                clientName: caseData.clientName || '',
                claimantName: caseData.claimantName || '',
                timestamp: new Date().toISOString()
            };

            // Save the normalized case as a separate localStorage item using its caseId
            localStorage.setItem(caseData.caseId, JSON.stringify(normalizedCase));

            // Only increment caseIdCounter if the current caseIdNum is higher than the existing counter
            if (caseIdNum >= currentIntakeData.caseIdCounter -1) { // -1 because counter is already incremented for next new case
                currentIntakeData.caseIdCounter = caseIdNum + 1;
            }
            localStorage.setItem('intakeData', JSON.stringify(currentIntakeData)); // Update caseIdCounter in intakeData

            showSuccessMessage('Case synced and saved locally!');

            status.innerHTML = `<span style="color: var(--success);">✅ Synced at ${new Date().toLocaleTimeString()}</span>`;
            btn.textContent = 'Sync to Sheets'; btn.disabled = false;

            // Reset the form after successful sync and update of local storage
            const form = document.querySelector('form');
            if (form) form.reset();

            // Clear the current form state to prepare for a new entry
            localData.currentForm = {};
            localStorage.setItem('intakeData', JSON.stringify(localData));
            loadLocalData(); // Reload localData to ensure defaults and cleared currentForm are active

            // Re-setup dropdowns/fields that depend on localData
            namedEntitySections.forEach(setupNamedEntitySection);
            setupPipelineStage();
            setupInvestigatorDropdown(); // Important for the investigator dropdown to reset based on cleared form
            setupStateField();
            setupSsnFormatting();

            btn.disabled = true; status.innerHTML = ''; // Disable sync button and clear status for new blank form
        } catch (e) {
            console.error('Sync error:', e);
            status.innerHTML = `<span style="color: var(--danger);">❌ Sync failed: ${e.message}</span>`;
            btn.textContent = 'Retry Sync'; btn.disabled = false;
        }
    }

    // Initial population of form fields from localData.currentForm
    for (const id in localData.currentForm) {
        const f = document.getElementById(id);
        if (f) {
            f.value = localData.currentForm[id];
            // Special handling for dynamic dropdowns that need re-initialization
            if (id === 'assignmentService') {
                setupInvestigatorDropdown(); // Ensure investigator dropdown updates based on loaded service
            }
        }
    }

    // Determine initial state of sync button and status on page load
    if (localData.currentForm.caseId && localStorage.getItem(localData.currentForm.caseId)) {
        document.getElementById('sync-btn').disabled = false;
        document.getElementById('sync-status').innerHTML = '<span style="color: var(--warning);">⚠️ Not synced to Sheets yet</span>';
    } else {
        document.getElementById('sync-btn').disabled = true;
        document.getElementById('sync-status').innerHTML = '';
    }
});
