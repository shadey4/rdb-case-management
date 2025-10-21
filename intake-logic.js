document.addEventListener('DOMContentLoaded', function() {
    const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwS2Ah2qP19aGcnvBnNlBwb9RFeUeID-zCuxjhQcMfpfEY0ppS7eHxYxTXaBCbxWGMA/exec'; // Updated URL to avoid collision
    let lastSavedCaseKey = null; // To track the last saved form for Sheets sync

    // ===== Load or initialize stored data =====
    const defaults = {
        investigators: ['Conrad', 'Delores', 'Lisa', 'Michael', 'Jeff'],
        surveillanceInvestigators: ['Freddy', 'Ricardo', 'Delores', 'Michael', 'Alma', 'Manuel', 'Jeff'],
        clients: [], // Used for global client autofill options, but actual form data stored in currentCaseData
        attorneys: [], // Used for global attorney autofill options
        employers: [], // Used for global employer autofill options
        claimants: [], // Used for global claimant autofill options
        pipelineStages: [],
        states: ["CA", "NY", "TX", "FL"],
        currentForm: {}, // Stores the active form's field values for auto-save
        caseIdCounter: 1,
            savedCases: {} // New structure for entity-isolated case data
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
        // Ensure currentForm is initialized correctly, merging stored with defaults
        localData.currentForm = { ...defaults.currentForm, ...(stored.currentForm || {}) };
        // Ensure savedCases is initialized correctly
        localData.savedCases = { ...defaults.savedCases, ...(stored.savedCases || {}) };
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

    // Event listeners for general fields (not part of named entities)
    document.querySelectorAll('input:not([data-section]), select:not([data-section]), textarea:not([data-section])').forEach(field => {
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
        const listKey = section + 's'; // e.g., 'clients', 'attorneys'

        // Populate datalists from the global lists (for autofill suggestions)
        if (nameField) {
            populateDatalist(section + 'NameList', localData[listKey].map(i => i.name).filter(Boolean));
            nameField.addEventListener('change', () => autofillSection(section, nameField.value, 'name'));
        }
        if (companyField) {
            populateDatalist(section + 'CompanyList', localData[listKey].map(i => i.company).filter(Boolean));
            companyField.addEventListener('change', () => autofillSection(section, companyField.value, 'company'));
        }

        // Add event listeners for all fields within the section to save on blur/change
        document.querySelectorAll(`[data-section="${section}"] input, [data-section="${section}"] select, [data-section="${section}"] textarea`)
        .forEach(f => {
            // Initial load of currentForm values
            if (localData.currentForm[f.id] !== undefined) f.value = localData.currentForm[f.id];

            f.addEventListener('blur', () => {
                saveCurrentFormState(f); // Save individual field state to currentForm
                saveNamedEntity(section); // Also attempt to save the whole entity
            });
            if (f.tagName === 'SELECT') f.addEventListener('change', () => {
                saveCurrentFormState(f); // Save individual field state to currentForm
                saveNamedEntity(section); // Also attempt to save the whole entity
            });
        });
    }

    /**
     * Autofills fields for a specific section based on a matching value (name or company).
     * This function is crucial for data isolation, now reading from currentForm and global lists.
     * @param {string} section - The section identifier (e.g., 'client', 'employer', 'claimant').
     * @param {string} value - The value to search for (e.g., a client name, employer company).
     * @param {string} triggerType - 'name' or 'company' to specify which field triggered the autofill.
     */
    function autofillSection(section, value, triggerType) {
        if (!value) return clearNamedEntitySection(section);

        const listKey = section + 's';
        const globalList = localData[listKey] || []; // Global list for suggestions
        let record = null;

        if (triggerType === 'name') {
            record = globalList.find(r => r.name === value);
        } else { // triggerType === 'company'
            record = globalList.find(r => r.company === value);
        }

        if (record) {
            // Fill all fields for this section from the found record
            document.querySelectorAll(`[data-section="${section}"] input, [data-section="${section}"] select, [data-section="${section}"] textarea`)
            .forEach(f => {
                let k = f.id.replace(section, '');
                if (k) {
                    k = k.charAt(0).toLowerCase() + k.slice(1);
                    if (record[k] !== undefined) {
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
            const triggerFieldId = section + capitalize(triggerType);
            if (f.id !== triggerFieldId) {
                if (f.tagName === 'SELECT') f.selectedIndex = 0;
                else if (f.type === 'date') f.value = '';
                else if (f.id.includes('State')) f.value = 'CA'; // Default state for state fields
                else f.value = '';
                saveCurrentFormState(f); // Save the cleared state to currentForm
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
            saveCurrentFormState(f); // Save the cleared state to currentForm
        });
    }

    function clearDependentFields(fieldId) {
        const section = namedEntitySections.find(s => fieldId.startsWith(s));
        if (section && (fieldId.endsWith('Name') || fieldId.endsWith('Company'))) {
            const currentFieldValue = document.getElementById(fieldId).value;
            if (currentFieldValue === '') {
                clearNamedEntitySection(section);
            }
        }
    }

    /**
     * Saves or updates a named entity (client, attorney, employer, claimant) to localData.
     * This now updates both the global list for autofill suggestions and the current case's specific entity data.
     * @param {string} section - The section identifier (e.g., 'client', 'employer', 'claimant').
     */
    function saveNamedEntity(section) {
        const record = collectNamedEntityValues(section);
        const listKey = section + 's'; // e.g., 'clients'

        // Check for empty records to avoid saving blank entries
        if (['client', 'attorney', 'employer'].includes(section) && !record.name && !record.company) return;
        if (section === 'claimant' && !Object.values(record).some(val => val && val.trim() !== '')) return; // Require at least one non-empty field for claimant

        // --- Update global list for autofill suggestions ---
        let globalList = localData[listKey] || [];
        let existingGlobalIndex = -1;
        if (record.name) {
            existingGlobalIndex = globalList.findIndex(r => r.name === record.name);
        } else if (record.company && section !== 'claimant') {
            existingGlobalIndex = globalList.findIndex(r => r.company === record.company);
        }

        if (existingGlobalIndex !== -1) {
            globalList[existingGlobalIndex] = { ...globalList[existingGlobalIndex], ...record };
        } else {
            globalList.push(record);
        }
        localData[listKey] = globalList;

        // --- Update current case's isolated entity data ---
        const caseId = localData.currentForm.caseId;
        if (caseId) {
            if (!localData.savedCases[caseId]) localData.savedCases[caseId] = {};
            // Defensive checks for nested entity objects
            if (!localData.savedCases[caseId].claimant) localData.savedCases[caseId].claimant = {};
            if (!localData.savedCases[caseId].employer) localData.savedCases[caseId].employer = {};
            if (!localData.savedCases[caseId].client) localData.savedCases[caseId].client = {};
            if (!localData.savedCases[caseId].attorney) localData.savedCases[caseId].attorney = {};

            localData.savedCases[caseId][section] = { ...localData.savedCases[caseId][section], ...record };
            console.log(`Saved ${section}:`, localData.savedCases[caseId][section]);
        }

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
     * Uses data-section attribute to ensure isolation.
     * @param {string} section - The section identifier (e.g., 'client', 'employer').
     * @returns {object} An object containing all collected field-value pairs for the section.
     */
    function collectNamedEntityValues(section) {
        const r = {};
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
                o.value = v;
                dl.appendChild(o);
            });
        }
    }

    function capitalize(s) {
        return s.charAt(0).toUpperCase() + s.slice(1);
    }

    function formatSSN(i) {
        let v = i.value.replace(/\D/g, ''),
                          f = '';
                          if (v.length > 0) f += v.substring(0, 3);
                          if (v.length > 3) f += '-' + v.substring(3, 5);
                          if (v.length > 5) f += '-' + v.substring(5, 9);
                          i.value = f;
        saveCurrentFormState(i); // Save the formatted SSN
        saveNamedEntity('claimant'); // Trigger save for the claimant section
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
            def.value = '';
            def.textContent = '-- Select Investigator --';
            iSel.appendChild(def);
            list.forEach(n => {
                const o = document.createElement('option');
                o.value = n;
                o.textContent = n;
                iSel.appendChild(o);
            });
            // Set value back, preferring currentForm, then trying to match current value, otherwise default
            iSel.value = localData.currentForm[iSel.id] || (list.includes(current) ? current : '');
            saveCurrentFormState(iSel);
        }
        sSel.addEventListener('change', update);
        update(); // Call once on load to set initial state
    }

    function saveFullForm() {
        // Ensure a caseId exists for the current form
        if (!localData.currentForm.caseId) {
            localData.currentForm.caseId = 'case_' + localData.caseIdCounter++;
            localStorage.setItem('intakeData', JSON.stringify(localData));
        }
        const caseId = localData.currentForm.caseId;

        // Collect all general (non-entity) form data
        const formData = {};
        document.querySelectorAll('input:not([data-section]), select:not([data-section]), textarea:not([data-section])')
        .forEach(f => {
            if (f.id) formData[f.id] = f.value;
        });

            // Normalize core metadata fields
            formData.caseId = caseId;
            formData.claimNumber = formData.assignmentClaimNumber || '';
            formData.status = formData.assignmentStatus || 'Open';
            formData.investigatorAssigned = formData.investigatorName || '';
            formData.dateCreated = new Date().toISOString();

            // Split claimant full name into first and last
            const cName = (formData.claimantName || '').trim().split(' ');
            formData.claimantFirst = cName[0] || '';
            formData.claimantLast = cName.slice(1).join(' ') || '';

            // --- Preserve isolated entity objects (do not flatten them) ---
            if (!localData.savedCases[caseId]) localData.savedCases[caseId] = {};
            namedEntitySections.forEach(section => {
                if (!localData.savedCases[caseId][section]) {
                    localData.savedCases[caseId][section] = {};
                }
            });

            // Merge current saved case with updated general form data
            localData.savedCases[caseId] = {
                ...localData.savedCases[caseId],
                ...formData
            };

            // Save back to localStorage
            localStorage.setItem('intakeData', JSON.stringify(localData));
            lastSavedCaseKey = caseId;

            // Enable sync button and update status
            document.getElementById('sync-btn').disabled = false;
            document.getElementById('sync-status').innerHTML =
            '<span style="color: var(--warning);">⚠️ Not synced to Sheets yet</span>';

            alert('Intake form saved to localStorage!\n\nClick "Sync to Sheets" to back up to Google Sheets.');

            // Update currentForm to reflect just-saved state
            localData.currentForm = { ...localData.currentForm, ...formData };
            localStorage.setItem('intakeData', JSON.stringify(localData));
    }
}

    async function syncToSheets() {
        if (!localData.currentForm.caseId) return alert('Please save first.');
        const key = localData.currentForm.caseId;

        // Retrieve the complete case data from savedCases
        const caseData = localData.savedCases[key];
        if (!caseData) return alert('No local data found for this case.');

        const btn = document.getElementById('sync-btn');
        const status = document.getElementById('sync-status');
        btn.disabled = true;
        btn.textContent = 'Syncing...';
        status.innerHTML = '<span style="color: var(--accent);">⏳ Syncing to Google Sheets...</span>';

        try {
            // Send the complete caseData to Google Sheets (including nested entities)
            await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ action: 'syncCase', caseData })
            });

            const currentIntakeData = JSON.parse(localStorage.getItem('intakeData')) || defaults;
            const caseIdNum = parseInt(key.replace('case_', '')) || 0;

            // Only increment caseIdCounter if the current caseIdNum is higher than or equal to the existing counter
            if (caseIdNum >= currentIntakeData.caseIdCounter - 1) { // -1 because counter is already incremented for next new case
                currentIntakeData.caseIdCounter = caseIdNum + 1;
            }
            localStorage.setItem('intakeData', JSON.stringify(currentIntakeData)); // Update caseIdCounter in intakeData

            showSuccessMessage('Case synced and saved locally!');

            status.innerHTML = `<span style="color: var(--success);">✅ Synced at ${new Date().toLocaleTimeString()}</span>`;
            btn.textContent = 'Sync to Sheets';
            btn.disabled = false;

            // --- Reset the form and localData for a new entry ---
            const form = document.querySelector('form');
            if (form) form.reset();

            // Clear the current form state and load defaults for a new entry
            localData.currentForm = {};
            localStorage.setItem('intakeData', JSON.stringify(localData)); // Save cleared currentForm
            loadLocalData(); // Reload localData to ensure defaults and cleared currentForm are active

            // Re-setup dropdowns/fields that depend on localData and now need to reflect a blank form
            namedEntitySections.forEach(setupNamedEntitySection);
            setupPipelineStage();
            setupInvestigatorDropdown();
            setupStateField();
            setupSsnFormatting();

            // Clear general fields that are not part of named entities
            document.querySelectorAll('input:not([data-section]), select:not([data-section]), textarea:not([data-section])').forEach(f => {
                if (f.tagName === 'SELECT') f.selectedIndex = 0;
                else if (f.type === 'date') f.value = '';
                else f.value = '';
            });

                btn.disabled = true;
                status.innerHTML = ''; // Disable sync button and clear status for new blank form
        } catch (e) {
            console.error('Sync error:', e);
            status.innerHTML = `<span style="color: var(--danger);">❌ Sync failed: ${e.message}</span>`;
            btn.textContent = 'Retry Sync';
            btn.disabled = false;
        }
    }

    // Initial population of form fields from localData.currentForm on page load
    // This will reconstruct the form from the last saved state.
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
    if (localData.currentForm.caseId && localData.savedCases[localData.currentForm.caseId]) {
        // If there's a current caseId and it exists in savedCases, enable sync
        document.getElementById('sync-btn').disabled = false;
        document.getElementById('sync-status').innerHTML = '<span style="color: var(--warning);">⚠️ Not synced to Sheets yet</span>';
    } else {
        // Otherwise, disable sync
        document.getElementById('sync-btn').disabled = true;
        document.getElementById('sync-status').innerHTML = '';
    }
});
