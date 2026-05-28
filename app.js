// AME Engineering Latur - Application Logic

let collegesDb = [];
let rankPercMapping = [];
let cleanRankPercMapping = [];
const topCollegeCodes = [3012, 16006, 3036, 3215, 6271, 6007, 6276, 6273, 3014, 3199];

// Active State for Multi-Select Cities
let allCities = [];
let selectedCities = [];

// Active State for Multi-Select Branches
let allBranches = [];
let selectedBranches = [];

// Active State for current prediction output (Excel generation)
let currentTopCollegesMatches = [];
let currentFinalMatches = [];
let currentRank = 0;
let currentPercentile = 0.0;
let currentCategory = "";
let currentGender = "";

// Cache DOM Elements
const predictorForm = document.getElementById('predictorForm');
const studentRankInput = document.getElementById('studentRank');
const studentPercentileInput = document.getElementById('studentPercentile');
const studentGenderSelect = document.getElementById('studentGender');
const studentCategorySelect = document.getElementById('studentCategory');
const seatHomeState = document.getElementById('seatHomeState');
const seatOtherState = document.getElementById('seatOtherState');
const seatStateLevel = document.getElementById('seatStateLevel');
const useCustomRange = document.getElementById('useCustomRange');
const minRankRange = document.getElementById('minRankRange');
const maxRankRange = document.getElementById('maxRankRange');
const advancedSettingsBody = document.getElementById('advancedSettingsBody');
const loader = document.getElementById('loader');
const resultsSection = document.getElementById('resultsSection');
const totalMatchesCount = document.getElementById('totalMatchesCount');
const topCollegesGrid = document.getElementById('topCollegesGrid');
const matchesTableBody = document.getElementById('matchesTableBody');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');
const printOptionFormBtn = document.getElementById('printOptionFormBtn');
const limitSelectorWrapper = document.getElementById('limitSelectorWrapper');
const listLimitSelect = document.getElementById('listLimitSelect');

// Custom Multi-Select Cities DOM Elements
const citySelectTrigger = document.getElementById('citySelectTrigger');
const cityDropdown = document.getElementById('cityDropdown');
const citySearchInput = document.getElementById('citySearchInput');
const cityOptionsContainer = document.getElementById('cityOptionsContainer');
const cityPlaceholder = document.getElementById('cityPlaceholder');
const cityTags = document.getElementById('cityTags');

// Custom Multi-Select Branches DOM Elements
const branchSelectTrigger = document.getElementById('branchSelectTrigger');
const branchDropdown = document.getElementById('branchDropdown');
const branchSearchInput = document.getElementById('branchSearchInput');
const branchOptionsContainer = document.getElementById('branchOptionsContainer');
const branchPlaceholder = document.getElementById('branchPlaceholder');
const branchTags = document.getElementById('branchTags');

// Modal Elements
const collegeModal = document.getElementById('collegeModal');
const modalCollegeCode = document.getElementById('modalCollegeCode');
const modalCollegeName = document.getElementById('modalCollegeName');
const modalCollegeCity = document.getElementById('modalCollegeCity');
const modalStudentRank = document.getElementById('modalStudentRank');
const modalStudentPercentile = document.getElementById('modalStudentPercentile');
const modalSelectedCategory = document.getElementById('modalSelectedCategory');
const modalCutoffsBody = document.getElementById('modalCutoffsBody');
const modalCloseBtn = document.getElementById('modalCloseBtn');

// Student Name Prompt Modal Elements
const namePromptModal = document.getElementById('namePromptModal');
const studentNameInput = document.getElementById('studentNameInput');
const namePromptCloseBtn = document.getElementById('namePromptCloseBtn');
const namePromptCancelBtn = document.getElementById('namePromptCancelBtn');
const namePromptSubmitBtn = document.getElementById('namePromptSubmitBtn');

// Initialize App: Load data
window.addEventListener('DOMContentLoaded', () => {
    loadCollegesData();
    loadRankPercentileMapping();
    setupEventListeners();

    // Ensure "Top 40 Ranks" is selected by default on initial load
    if (listLimitSelect) {
        listLimitSelect.value = '40';
    }
});

// Event Listeners Setup
function setupEventListeners() {
    predictorForm.addEventListener('submit', handleFormSubmit);
    
    // Close modal triggers
    modalCloseBtn.addEventListener('click', closeModal);
    collegeModal.addEventListener('click', (e) => {
        if (e.target === collegeModal) closeModal();
    });
    
    // Escape key close modal
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            if (!collegeModal.classList.contains('hidden')) closeModal();
            if (!namePromptModal.classList.contains('hidden')) closeNamePromptModal();
            closeCityDropdown();
            closeBranchDropdown();
        }
    });

    // Custom Multi-Select Cities Interactions
    citySelectTrigger.addEventListener('click', (e) => {
        e.stopPropagation();
        closeBranchDropdown(); // Close branch dropdown if open
        toggleCityDropdown();
    });

    citySearchInput.addEventListener('click', (e) => {
        e.stopPropagation(); 
    });

    citySearchInput.addEventListener('input', (e) => {
        filterCityCheckboxes(e.target.value);
    });

    // Custom Multi-Select Branches Interactions
    branchSelectTrigger.addEventListener('click', (e) => {
        e.stopPropagation();
        closeCityDropdown(); // Close city dropdown if open
        toggleBranchDropdown();
    });

    branchSearchInput.addEventListener('click', (e) => {
        e.stopPropagation(); 
    });

    branchSearchInput.addEventListener('input', (e) => {
        filterBranchCheckboxes(e.target.value);
    });

    // Close dropdowns on click outside
    document.addEventListener('click', (e) => {
        if (!e.target.closest('#cityMultiselect')) {
            closeCityDropdown();
        }
        if (!e.target.closest('#branchMultiselect')) {
            closeBranchDropdown();
        }
    });

    // Download Excel Button Interaction (Opens Name Modal)
    if (downloadExcelBtn) {
        downloadExcelBtn.addEventListener('click', (e) => {
            e.preventDefault();
            openNamePromptModal();
        });
    }

    if (printOptionFormBtn) {
        printOptionFormBtn.addEventListener('click', (e) => {
            e.preventDefault();
            printOptionForm();
        });
    }

    // Name Prompt Modal Close triggers
    if (namePromptCloseBtn) namePromptCloseBtn.addEventListener('click', closeNamePromptModal);
    if (namePromptCancelBtn) namePromptCancelBtn.addEventListener('click', closeNamePromptModal);
    if (namePromptModal) {
        namePromptModal.addEventListener('click', (e) => {
            if (e.target === namePromptModal) closeNamePromptModal();
        });
    }

    // Name Prompt Modal Submissions
    if (studentNameInput) {
        studentNameInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                handleNamePromptSubmit();
            }
        });
    }
    if (namePromptSubmitBtn) {
        namePromptSubmitBtn.addEventListener('click', (e) => {
            e.preventDefault();
            handleNamePromptSubmit();
        });
    }

    // Form reset handler to clear custom multi-select UI and results
    predictorForm.addEventListener('reset', () => {
        // Clear selection arrays
        selectedCities = [];
        selectedBranches = [];
        
        // Uncheck all checkboxes in custom multiselect dropdowns
        const multiSelectCheckboxes = predictorForm.querySelectorAll('.multiselect-dropdown input[type="checkbox"]');
        multiSelectCheckboxes.forEach(chk => {
            chk.checked = false;
        });
        
        // Remove 'selected' class from option divs
        const optionDivs = predictorForm.querySelectorAll('.multiselect-option');
        optionDivs.forEach(div => {
            div.classList.remove('selected');
        });
        
        // Reset Seat Type checkboxes (checked by default)
        if (seatHomeState) seatHomeState.checked = true;
        if (seatOtherState) seatOtherState.checked = true;
        if (seatStateLevel) seatStateLevel.checked = true;
        
        // Reset Advanced Filters (Always checked and visible by default)
        if (useCustomRange) {
            useCustomRange.checked = true;
            advancedSettingsBody.classList.remove('hidden');
            minRankRange.setAttribute('required', 'true');
            maxRankRange.setAttribute('required', 'true');
            minRankRange.value = '';
            maxRankRange.value = '';
        }
        
        // Reset Trigger UIs (restore placeholders, hide tags)
        updateCityTriggerUI();
        updateBranchTriggerUI();
        
        // Reset List Limit Selection
        if (listLimitSelect) {
            listLimitSelect.value = '40';
        }
        if (limitSelectorWrapper) {
            limitSelectorWrapper.classList.remove('hidden');
        }
        
        // Hide results and scroll to top smoothly
        resultsSection.classList.add('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    // Toggle expand/collapse `#advancedSettingsBody` and update input requirements when useCustomRange is toggled
    if (useCustomRange) {
        useCustomRange.addEventListener('change', () => {
            if (useCustomRange.checked) {
                advancedSettingsBody.classList.remove('hidden');
                minRankRange.setAttribute('required', 'true');
                maxRankRange.setAttribute('required', 'true');
            } else {
                advancedSettingsBody.classList.add('hidden');
                minRankRange.removeAttribute('required');
                maxRankRange.removeAttribute('required');
                minRankRange.value = '';
                maxRankRange.value = '';
            }
        });
    }

    // Real-time two-way Rank and Percentile Auto-Fetching
    if (studentRankInput && studentPercentileInput) {
        studentRankInput.addEventListener('input', syncPercentileFromRank);
        studentRankInput.addEventListener('change', syncPercentileFromRank);
        studentPercentileInput.addEventListener('input', syncRankFromPercentile);
        studentPercentileInput.addEventListener('change', syncRankFromPercentile);
    }

    // List Limit Dropdown Change Listener - Dynamically re-evaluate prediction
    if (listLimitSelect) {
        listLimitSelect.addEventListener('change', () => {
            if (currentRank > 0) {
                // Show loader
                loader.classList.remove('hidden');
                resultsSection.classList.add('hidden');

                setTimeout(() => {
                    performPrediction(currentRank, currentPercentile, currentGender, currentCategory);
                    loader.classList.add('hidden');
                    resultsSection.classList.remove('hidden');
                }, 300);
            }
        });
    }
}

// Load colleges data from JSON
async function loadCollegesData() {
    try {
        console.log("Fetching colleges data...");
        const response = await fetch('colleges_data.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        collegesDb = await response.json();
        console.log(`Loaded ${collegesDb.length} colleges successfully.`);
        
        // Populate selectors
        populateDropdowns();
    } catch (error) {
        console.error("Failed to load colleges data:", error);
        alert("Error loading colleges data. Please make sure the colleges_data.json file is present in the directory.");
    }
}

// Load rank-percentile mapping data from JSON
async function loadRankPercentileMapping() {
    try {
        console.log("Fetching rank-percentile mapping...");
        const response = await fetch('rank_percentile_mapping.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        rankPercMapping = await response.json();
        cleanRankPercMapping = rankPercMapping
            .map(([rank, percentile]) => [Number(rank), Number(percentile)])
            .filter(([rank, percentile]) => Number.isFinite(rank) && rank > 0 && Number.isFinite(percentile) && percentile > 0)
            .sort((a, b) => a[0] - b[0]);
        console.log(`Loaded ${cleanRankPercMapping.length} valid rank-percentile mapping records successfully.`);
        syncAutoFilledScore();
    } catch (error) {
        console.error("Failed to load rank-percentile mapping:", error);
    }
}

function syncAutoFilledScore() {
    if (!studentRankInput || !studentPercentileInput) return;

    if (studentRankInput.value.trim()) {
        syncPercentileFromRank();
        return;
    }

    if (studentPercentileInput.value.trim()) {
        syncRankFromPercentile();
    }
}

function syncPercentileFromRank() {
    if (!studentRankInput || !studentPercentileInput) return;

    const value = studentRankInput.value.trim();
    if (!value) {
        studentPercentileInput.value = '';
        return;
    }

    const rank = parseInt(value, 10);
    if (!Number.isFinite(rank) || rank <= 0) return;

    const percentile = findClosestPercentileForRank(rank);
    if (percentile !== null) {
        studentPercentileInput.value = formatPercentile(percentile);
    }
}

function syncRankFromPercentile() {
    if (!studentRankInput || !studentPercentileInput) return;

    const value = studentPercentileInput.value.trim();
    if (!value) {
        studentRankInput.value = '';
        return;
    }

    const percentile = parseFloat(value);
    if (!Number.isFinite(percentile) || percentile < 0 || percentile > 100) return;

    const rank = findClosestRankForPercentile(percentile);
    if (rank !== null) {
        studentRankInput.value = rank;
    }
}

function formatPercentile(percentile) {
    return Number(percentile).toFixed(4).replace(/\.?0+$/, '');
}

// Find closest percentile for a given rank using O(log N) binary search
function findClosestPercentileForRank(targetRank) {
    if (!cleanRankPercMapping || cleanRankPercMapping.length === 0) return null;
    let low = 0;
    let high = cleanRankPercMapping.length - 1;
    let closestIndex = 0;
    let minDiff = Infinity;
    
    while (low <= high) {
        let mid = Math.floor((low + high) / 2);
        let currentRank = cleanRankPercMapping[mid][0];
        let diff = Math.abs(currentRank - targetRank);
        
        if (diff < minDiff) {
            minDiff = diff;
            closestIndex = mid;
        }
        
        if (currentRank === targetRank) {
            break;
        } else if (currentRank < targetRank) {
            low = mid + 1;
        } else {
            high = mid - 1;
        }
    }
    return cleanRankPercMapping[closestIndex][1];
}

// Find closest rank for a given percentile using robust scan
function findClosestRankForPercentile(targetPercentile) {
    if (!cleanRankPercMapping || cleanRankPercMapping.length === 0) return null;
    let closestRank = null;
    let minDiff = Infinity;
    
    for (let i = 0; i < cleanRankPercMapping.length; i++) {
        let diff = Math.abs(cleanRankPercMapping[i][1] - targetPercentile);
        if (diff < minDiff) {
            minDiff = diff;
            closestRank = cleanRankPercMapping[i][0];
        }
    }
    return closestRank;
}

// Extract unique cities and branch groups from the database and populate selectors
function populateDropdowns() {
    const citiesSet = new Set();
    const branchGroupsSet = new Set();

    collegesDb.forEach(college => {
        if (college.ci) citiesSet.add(college.ci.toUpperCase());
        
        if (college.b && Array.isArray(college.b)) {
            college.b.forEach(branch => {
                if (branch.g) branchGroupsSet.add(branch.g);
            });
        }
    });

    // Sort and Save Cities
    allCities = Array.from(citiesSet).sort();
    renderCityCheckboxes();

    // Sort and Save Branches
    allBranches = Array.from(branchGroupsSet).sort();
    renderBranchCheckboxes();
}

/* ==========================================
   CUSTOM MULTI-SELECT CITIES LOGIC
   ========================================== */

function toggleCityDropdown() {
    const isActive = citySelectTrigger.classList.contains('active');
    if (isActive) {
        closeCityDropdown();
    } else {
        citySelectTrigger.classList.add('active');
        cityDropdown.classList.remove('hidden');
        citySearchInput.focus();
    }
}

function closeCityDropdown() {
    citySelectTrigger.classList.remove('active');
    cityDropdown.classList.add('hidden');
    citySearchInput.value = '';
    filterCityCheckboxes(''); // Reset search filters
}

function renderCityCheckboxes() {
    cityOptionsContainer.innerHTML = '';
    
    allCities.forEach(city => {
        const optionDiv = document.createElement('div');
        optionDiv.className = 'multiselect-option';
        optionDiv.dataset.city = city;
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = city;
        checkbox.id = `chk_${city.replace(/\s+/g, '_')}`;
        
        const label = document.createElement('label');
        label.htmlFor = checkbox.id;
        label.textContent = formatString(city);
        label.style.cursor = 'pointer';
        label.style.flex = '1';
        
        optionDiv.appendChild(checkbox);
        optionDiv.appendChild(label);
        
        // Single robust event listener for the entire option block
        optionDiv.addEventListener('click', (e) => {
            e.stopPropagation();
            
            // If the checkbox itself was clicked, let the browser toggle it and just sync the UI state
            if (e.target === checkbox) {
                handleCityCheckboxChange(city, checkbox.checked);
                return;
            }
            
            // Otherwise (clicked label or background), toggle checkbox manually and prevent double-triggering
            e.preventDefault();
            checkbox.checked = !checkbox.checked;
            handleCityCheckboxChange(city, checkbox.checked);
        });

        cityOptionsContainer.appendChild(optionDiv);
    });
}

function handleCityCheckboxChange(city, isChecked) {
    if (isChecked) {
        if (!selectedCities.includes(city)) {
            selectedCities.push(city);
        }
    } else {
        selectedCities = selectedCities.filter(c => c !== city);
    }
    
    const optionDiv = cityOptionsContainer.querySelector(`.multiselect-option[data-city="${city}"]`);
    if (optionDiv) {
        if (isChecked) optionDiv.classList.add('selected');
        else optionDiv.classList.remove('selected');
    }

    updateCityTriggerUI();
}

function filterCityCheckboxes(searchQuery) {
    const query = searchQuery.toUpperCase();
    const options = cityOptionsContainer.querySelectorAll('.multiselect-option');
    
    options.forEach(opt => {
        const city = opt.dataset.city;
        if (city.includes(query) || formatString(city).toUpperCase().includes(query)) {
            opt.classList.remove('hidden');
        } else {
            opt.classList.add('hidden');
        }
    });
}

function updateCityTriggerUI() {
    cityTags.innerHTML = '';
    
    if (selectedCities.length === 0) {
        cityPlaceholder.classList.remove('hidden');
        cityTags.classList.add('hidden');
        return;
    }
    
    cityPlaceholder.classList.add('hidden');
    cityTags.classList.remove('hidden');
    
    selectedCities.forEach(city => {
        const tag = document.createElement('span');
        tag.className = 'multiselect-tag';
        tag.innerHTML = `
            <span>${formatString(city)}</span>
            <i class="fa-solid fa-xmark" data-city="${city}"></i>
        `;
        
        tag.querySelector('i').addEventListener('click', (e) => {
            e.stopPropagation();
            const checkbox = document.getElementById(`chk_${city.replace(/\s+/g, '_')}`);
            if (checkbox) checkbox.checked = false;
            handleCityCheckboxChange(city, false);
        });
        
        cityTags.appendChild(tag);
    });
}

/* ==========================================
   CUSTOM MULTI-SELECT BRANCHES LOGIC
   ========================================== */

function toggleBranchDropdown() {
    const isActive = branchSelectTrigger.classList.contains('active');
    if (isActive) {
        closeBranchDropdown();
    } else {
        branchSelectTrigger.classList.add('active');
        branchDropdown.classList.remove('hidden');
        branchSearchInput.focus();
    }
}

function closeBranchDropdown() {
    branchSelectTrigger.classList.remove('active');
    branchDropdown.classList.add('hidden');
    branchSearchInput.value = '';
    filterBranchCheckboxes(''); // Reset search filters
}

function renderBranchCheckboxes() {
    branchOptionsContainer.innerHTML = '';
    
    allBranches.forEach(branch => {
        const optionDiv = document.createElement('div');
        optionDiv.className = 'multiselect-option';
        optionDiv.dataset.branch = branch;
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = branch;
        // Strip out brackets/special characters for valid html ID attributes
        const safeId = branch.replace(/[^a-zA-Z0-9]/g, '_');
        checkbox.id = `chk_br_${safeId}`;
        
        const label = document.createElement('label');
        label.htmlFor = checkbox.id;
        label.textContent = branch;
        label.style.cursor = 'pointer';
        label.style.flex = '1';
        
        optionDiv.appendChild(checkbox);
        optionDiv.appendChild(label);
        
        // Single robust event listener for the entire option block
        optionDiv.addEventListener('click', (e) => {
            e.stopPropagation();
            
            // If the checkbox itself was clicked, let the browser toggle it and just sync the UI state
            if (e.target === checkbox) {
                handleBranchCheckboxChange(branch, checkbox.checked);
                return;
            }
            
            // Otherwise (clicked label or background), toggle checkbox manually and prevent double-triggering
            e.preventDefault();
            checkbox.checked = !checkbox.checked;
            handleBranchCheckboxChange(branch, checkbox.checked);
        });

        branchOptionsContainer.appendChild(optionDiv);
    });
}

function handleBranchCheckboxChange(branch, isChecked) {
    if (isChecked) {
        if (!selectedBranches.includes(branch)) {
            selectedBranches.push(branch);
        }
    } else {
        selectedBranches = selectedBranches.filter(b => b !== branch);
    }
    
    const optionDiv = branchOptionsContainer.querySelector(`.multiselect-option[data-branch="${branch}"]`);
    if (optionDiv) {
        if (isChecked) optionDiv.classList.add('selected');
        else optionDiv.classList.remove('selected');
    }

    updateBranchTriggerUI();
}

function filterBranchCheckboxes(searchQuery) {
    const query = searchQuery.toUpperCase();
    const options = branchOptionsContainer.querySelectorAll('.multiselect-option');
    
    options.forEach(opt => {
        const branch = opt.dataset.branch;
        if (branch.toUpperCase().includes(query)) {
            opt.classList.remove('hidden');
        } else {
            opt.classList.add('hidden');
        }
    });
}

function updateBranchTriggerUI() {
    branchTags.innerHTML = '';
    
    if (selectedBranches.length === 0) {
        branchPlaceholder.classList.remove('hidden');
        branchTags.classList.add('hidden');
        return;
    }
    
    branchPlaceholder.classList.add('hidden');
    branchTags.classList.remove('hidden');
    
    selectedBranches.forEach(branch => {
        const tag = document.createElement('span');
        tag.className = 'multiselect-tag';
        tag.innerHTML = `
            <span>${branch}</span>
            <i class="fa-solid fa-xmark" data-branch="${branch}"></i>
        `;
        
        tag.querySelector('i').addEventListener('click', (e) => {
            e.stopPropagation();
            const safeId = branch.replace(/[^a-zA-Z0-9]/g, '_');
            const checkbox = document.getElementById(`chk_br_${safeId}`);
            if (checkbox) checkbox.checked = false;
            handleBranchCheckboxChange(branch, false);
        });
        
        branchTags.appendChild(tag);
    });
}

/* ==========================================
   Simplified Categories Core Mapper
   ========================================== */

function getCategoryIndex(categoryVal) {
    const mapping = {
        "OPEN": 1,
        "SC": 2,
        "ST": 3,
        "VJ (DT-A)": 4,
        "NT-1 (NT-B)": 5,
        "NT-2 (NT-C)": 6,
        "NT-3 (NT-D)": 7,
        "OBC": 8,
        "SEBC": 9,
        "DEF": 10,
        "PWD": 11,
        "EWS": 12,
        "TFWS": 13,
        "MI": 14,
        "ORPHAN": 15
    };
    return mapping[categoryVal] || 0;
}

// Resolve the seat type of a category key based on suffix (H, O, S) or specific quota name
function getSeatTypeOfCategory(catKey) {
    const rc = catKey.toUpperCase();
    if (rc.endsWith('H')) {
        return 'HomeState';
    } else if (rc.endsWith('O')) {
        return 'OtherState';
    } else if (rc.endsWith('S')) {
        return 'StateLevel';
    } else if (['EWS', 'TFWS', 'MI', 'ORPHAN'].includes(rc)) {
        return 'HomeState'; // compulsorily grouped with Home State candidate eligibility
    }
    return 'StateLevel'; // default fallback for other defense/PWD codes
}

// Check if a category matches the student's selected seat types
function matchesSeatType(catKey, selectedSeatTypes) {
    // EWS and TFWS match for all selections of Home State, Other State, or State Level
    if (['EWS', 'TFWS'].includes(catKey.toUpperCase())) {
        return true;
    }
    
    if (!selectedSeatTypes || selectedSeatTypes.length === 0) return true;
    const seatType = getSeatTypeOfCategory(catKey);
    return selectedSeatTypes.includes(seatType);
}

function matchesExcelCategory(rowCategory, gender, categoryIndex, selectedSeatTypes) {
    // 1. First filter by Seat Type!
    if (!matchesSeatType(rowCategory, selectedSeatTypes)) {
        return false;
    }

    const rc = rowCategory.toUpperCase();
    const prefix = gender === 'MALE' ? 'G' : 'L';
    
    if (categoryIndex >= 1 && categoryIndex <= 9) {
        const suffixes = [
            "",      // 0 placeholder
            "OPEN",  // 1
            "SC",    // 2
            "ST",    // 3
            "VJ",    // 4
            "NT1",   // 5
            "NT2",   // 6
            "NT3",   // 7
            "OBC",   // 8
            "SEBC"   // 9
        ];
        const suffix = suffixes[categoryIndex];
        const searchStr = prefix + suffix;
        return rc.includes(searchStr);
    } else if (categoryIndex === 10) {
        return rc.includes("DEF");
    } else if (categoryIndex === 11) {
        return rc.includes("PWD");
    } else if (categoryIndex === 12) {
        return rc === "EWS";
    } else if (categoryIndex === 13) {
        return rc === "TFWS";
    } else if (categoryIndex === 14) {
        return rc === "MI";
    } else if (categoryIndex === 15) {
        return rc === "ORPHAN";
    }
    return false;
}

// Check if a college is women-only
function isWomenOnlyCollege(college) {
    const name = college.n.toUpperCase();
    return name.includes("WOMEN") || name.includes("GIRLS") || name.includes("LADIES") || name.includes("CUMMINS");
}

/* ==========================================
   PREDICTION PROCESSING PIPELINE
   ========================================== */

// Handle Predictor Form Submission
function handleFormSubmit(e) {
    e.preventDefault();
    
    const rank = parseInt(studentRankInput.value);
    const percentile = parseFloat(studentPercentileInput.value);
    const gender = studentGenderSelect.value;
    const category = studentCategorySelect.value;

    if (isNaN(rank) || isNaN(percentile) || !category) {
        alert("Please enter valid merit rank, percentile, and select a category.");
        return;
    }

    // 1. Seat Type Validation
    const selectedSeatTypes = [];
    if (seatHomeState && seatHomeState.checked) selectedSeatTypes.push('HomeState');
    if (seatOtherState && seatOtherState.checked) selectedSeatTypes.push('OtherState');
    if (seatStateLevel && seatStateLevel.checked) selectedSeatTypes.push('StateLevel');
    
    if (selectedSeatTypes.length === 0) {
        alert("Please select at least one Seat Type (Home State, Other State, or State Level).");
        return;
    }

    // 2. Custom Range Validation
    if (useCustomRange && useCustomRange.checked) {
        const minRank = parseInt(minRankRange.value);
        const maxRank = parseInt(maxRankRange.value);
        if (isNaN(minRank) || isNaN(maxRank) || minRank <= 0 || maxRank <= 0) {
            alert("Please enter valid custom min and max ranks.");
            return;
        }
        if (minRank > maxRank) {
            alert("Min Rank cannot be greater than Max Rank.");
            return;
        }
    }

    // Show loader
    loader.classList.remove('hidden');
    resultsSection.classList.add('hidden');

    // Smooth experience delay
    setTimeout(() => {
        performPrediction(rank, percentile, gender, category);
        loader.classList.add('hidden');
        resultsSection.classList.remove('hidden');
        
        // Scroll to results
        resultsSection.scrollIntoView({ behavior: 'smooth' });
    }, 600);
}

// College Prediction Core Engine
function performPrediction(rank, percentile, gender, category) {
    // Retrieve list limit choice
    let limitSelectVal = 40;
    if (listLimitSelect) {
        limitSelectVal = parseInt(listLimitSelect.value) || 40;
    }

    const matchedBranchesList = [];
    const topCollegesMatches = [];
    const top10RecordsMap = {}; // Maps collegeCode -> best/closest branch record
    const topCollegeBranchPool = [];
    
    // Clear display grids/tables
    topCollegesGrid.innerHTML = '';
    matchesTableBody.innerHTML = '';

    const categoryIndex = getCategoryIndex(category);

    // Retrieve active Seat Types
    const selectedSeatTypes = [];
    if (seatHomeState && seatHomeState.checked) selectedSeatTypes.push('HomeState');
    if (seatOtherState && seatOtherState.checked) selectedSeatTypes.push('OtherState');
    if (seatStateLevel && seatStateLevel.checked) selectedSeatTypes.push('StateLevel');

    collegesDb.forEach(college => {
        const isTopCollege = topCollegeCodes.includes(college.c);
        
        // 1. GENDER RULE: If male, filter out women-only colleges!
        if (gender === 'MALE' && isWomenOnlyCollege(college)) {
            return;
        }

        // 2. CITY PREFERENCE RULE: Check multi-select filters (exempt top colleges so they are always evaluated)
        if (selectedCities.length > 0) {
            if (!selectedCities.includes(college.ci.toUpperCase())) {
                if (!isTopCollege) return;
            }
        }

        const collegeMatchingBranches = [];

        if (college.b && Array.isArray(college.b)) {
            college.b.forEach(branch => {
                // 3. BRANCH PREFERENCE RULE: Check multi-select filters
                if (selectedBranches.length > 0) {
                    if (!selectedBranches.includes(branch.g)) return;
                }

                // Scan all available round cutoffs for all matched category keys
                Object.keys(branch.co || {}).forEach(catKey => {
                    if (matchesExcelCategory(catKey, gender, categoryIndex, selectedSeatTypes)) {
                        const curCutoffs = branch.co[catKey];
                        
                        // Locate the round yielding the MAXIMUM rank (easiest round)
                        let maxRank = null;
                        let minPerc = null;
                        
                        if (curCutoffs.length >= 10) {
                            maxRank = curCutoffs[8];
                            minPerc = curCutoffs[9];
                        }
                        
                        // Fallback computation
                        if (maxRank === null || maxRank === undefined || maxRank === "") {
                            let computedMaxRank = null;
                            let computedMinPerc = null;
                            for (let r = 0; r < 4; r++) {
                                const rRank = curCutoffs[r * 2];
                                const rPerc = curCutoffs[r * 2 + 1];
                                if (rRank !== null && rRank !== undefined && rRank !== "") {
                                    const rankInt = parseInt(rRank);
                                    if (!isNaN(rankInt)) {
                                        if (computedMaxRank === null || rankInt > computedMaxRank) {
                                            computedMaxRank = rankInt;
                                            computedMinPerc = rPerc;
                                        }
                                    }
                                }
                            }
                            maxRank = computedMaxRank;
                            minPerc = computedMinPerc;
                        }

                        // Excel dynamic status classification
                        let status = "HIGHER CHANCES";
                        let helper = 3;

                        if (maxRank !== null) {
                            if (maxRank <= rank * 0.8) {
                                status = "DREAM";
                                helper = 1;
                            } else if (maxRank <= rank * 1.2) {
                                status = "MEDIUM CHANCES";
                                helper = 2;
                            } else {
                                status = "HIGHER CHANCES";
                                helper = 3;
                            }
                        } else if (minPerc !== null) {
                            if (minPerc >= percentile + 2.0) {
                                status = "DREAM";
                                helper = 1;
                            } else if (minPerc >= percentile - 2.0) {
                                status = "MEDIUM CHANCES";
                                helper = 2;
                            } else {
                                status = "HIGHER CHANCES";
                                helper = 3;
                            }
                        }

                        // Dynamic college description
                        const dynamicDesc = getDynamicDescription(maxRank);

                        const branchRecord = {
                            collegeCode: college.c,
                            collegeName: college.n,
                            city: college.ci,
                            branchCode: branch.c,
                            branchName: branch.n,
                            branchGroup: branch.g,
                            cutoffCategory: catKey,
                            maxRank: maxRank,
                            minPerc: minPerc,
                            chance: status,
                            helper: helper,
                            description: dynamicDesc
                        };

                        collegeMatchingBranches.push(branchRecord);
                        matchedBranchesList.push(branchRecord);
                    }
                });
            });
        }

        // Add top colleges for benchmarking grid (we select their best matching branch)
        if (isTopCollege) {
            let bestBranchRecord = null;
            if (collegeMatchingBranches.length > 0) {
                // Sort by helper descending so HIGHER CHANCES (3) comes first, then MEDIUM CHANCES (2), then DREAM (1)
                collegeMatchingBranches.sort((a, b) => b.helper - a.helper);
                bestBranchRecord = collegeMatchingBranches[0];
            } else {
                // Find closest branch fallback if user rank is higher and doesn't meet standard requirements
                let closestBranch = null;
                let closestDistance = Infinity;
                let closestCutoffRank = null;
                let closestCategory = null;

                if (college.b && Array.isArray(college.b)) {
                    college.b.forEach(branch => {
                        if (selectedBranches.length > 0 && !selectedBranches.includes(branch.g)) return;
                        Object.keys(branch.co || {}).forEach(catKey => {
                            // If college is women-only, evaluate using 'FEMALE' rules to match female cutoffs
                            const evalGender = isWomenOnlyCollege(college) ? 'FEMALE' : gender;
                            if (matchesExcelCategory(catKey, evalGender, categoryIndex, selectedSeatTypes)) {
                                const curCutoffs = branch.co[catKey];
                                for (let r = 0; r < 4; r++) {
                                    const rRank = curCutoffs[r * 2];
                                    if (rRank !== null && rRank !== undefined && rRank !== "") {
                                        const dist = Math.abs(rank - rRank);
                                        if (dist < closestDistance) {
                                            closestDistance = dist;
                                            closestBranch = branch;
                                            closestCutoffRank = rRank;
                                            closestCategory = catKey;
                                        }
                                    }
                                }
                            }
                        });
                    });
                }

                if (closestBranch) {
                    bestBranchRecord = {
                        collegeCode: college.c,
                        collegeName: college.n,
                        city: college.ci,
                        branchCode: closestBranch.c,
                        branchName: closestBranch.n,
                        branchGroup: closestBranch.g,
                        cutoffCategory: closestCategory,
                        maxRank: closestCutoffRank,
                        minPerc: null,
                        chance: 'DREAM',
                        helper: 1,
                        description: getDynamicDescription(closestCutoffRank)
                    };
                }
            }

            if (bestBranchRecord) {
                top10RecordsMap[college.c] = bestBranchRecord;
            }

            collegeMatchingBranches
                .sort(compareTopCollegeBranches)
                .forEach(record => topCollegeBranchPool.push(record));

            topCollegesMatches.push({
                code: college.c,
                name: college.n,
                city: college.ci,
                bestBranch: bestBranchRecord ? {
                    branchCode: bestBranchRecord.branchCode,
                    branchName: bestBranchRecord.branchName,
                    branchGroup: bestBranchRecord.branchGroup,
                    evaluation: {
                        chance: bestBranchRecord.chance,
                        cutoffRankUsed: bestBranchRecord.maxRank,
                        cutoffPercUsed: bestBranchRecord.minPerc
                    }
                } : null
            });
        }
    });

    const top10Records = [];
    const top10Keys = new Set();
    topCollegeCodes.forEach(code => {
        if (top10RecordsMap[code]) {
            addTopPreferenceRecord(top10Records, top10Keys, top10RecordsMap[code]);
        }
    });

    if (top10Records.length < 10) {
        topCollegeBranchPool.forEach(record => {
            if (top10Records.length < 10) {
                addTopPreferenceRecord(top10Records, top10Keys, record);
            }
        });
    }

    top10Records.sort((a, b) => {
        const collegeOrder = topCollegeCodes.indexOf(a.collegeCode) - topCollegeCodes.indexOf(b.collegeCode);
        if (collegeOrder !== 0) return collegeOrder;
        return compareTopCollegeBranches(a, b);
    });

    function addTopPreferenceRecord(records, keys, record) {
        const key = `${record.collegeCode}-${record.branchCode}-${record.cutoffCategory}`;
        if (!keys.has(key)) {
            records.push(record);
            keys.add(key);
        }
    }

    let finalMatches = [];
    const targetLimit = Math.max(0, limitSelectVal - top10Records.length);

    // Recommendation table sorting: Medium first, then Higher. Top 10 dream colleges are added separately at the top.
    const unifiedSortFunc = (a, b) => {
        if (a.helper !== b.helper) {
            return a.helper - b.helper;
        }
        
        const aVal = a.maxRank !== null ? a.maxRank : 999999;
        const bVal = b.maxRank !== null ? b.maxRank : 999999;
        if (aVal !== bVal) {
            return aVal - bVal;
        }
        
        const aPerc = a.minPerc !== null ? a.minPerc : 0;
        const bPerc = b.minPerc !== null ? b.minPerc : 0;
        return bPerc - aPerc; // higher percentile first
    };

    const isMainRecommendation = item => item.chance !== 'DREAM' && !topCollegeCodes.includes(item.collegeCode);
    const targetCandidates = matchedBranchesList.filter(isMainRecommendation);

    if (useCustomRange && useCustomRange.checked) {
        const minRank = parseInt(minRankRange.value);
        let maxRank = parseInt(maxRankRange.value);
        let rangeMatches = [];
        const expansionStep = getRankExpansionStep(maxRank);
        let expansionSteps = 0;

        while (true) {
            rangeMatches = targetCandidates.filter(item => {
                if (item.maxRank !== null) {
                    return item.maxRank >= minRank && item.maxRank <= maxRank;
                } else if (item.minPerc !== null) {
                    return item.minPerc <= percentile;
                }
                return false;
            });

            if (hasEnoughBalancedRecommendations(rangeMatches, targetLimit) || expansionSteps >= 150) {
                break;
            }

            maxRank += expansionStep;
            expansionSteps += 1;
        }

        rangeMatches.sort(unifiedSortFunc);
        finalMatches = mergeTopAndTargetRecommendations(top10Records, buildBalancedRecommendations(rangeMatches, targetLimit), limitSelectVal);
    } else {
        // Dynamic threshold search loop for medium / higher chance options
        let threshold = 1500;
        let selectedTargets = [];
        
        while (true) {
            selectedTargets = targetCandidates.filter(item => {
                if (item.maxRank !== null) {
                    return item.maxRank >= rank && item.maxRank <= rank + threshold;
                } else if (item.minPerc !== null) {
                    const percThreshold = 2.5 + ((threshold - 1500) / 1000) * 1.5;
                    return item.minPerc >= percentile - percThreshold && item.minPerc <= percentile;
                }
                return false;
            });
            
            if (hasEnoughBalancedRecommendations(selectedTargets, targetLimit) || threshold > 150000) {
                break;
            }
            threshold += 1000;
        }

        selectedTargets.sort(unifiedSortFunc);
        finalMatches = mergeTopAndTargetRecommendations(top10Records, buildBalancedRecommendations(selectedTargets, targetLimit), limitSelectVal);
    }

    // Save current match data for Excel download
    currentTopCollegesMatches = topCollegesMatches;
    currentFinalMatches = finalMatches;
    currentRank = rank;
    currentPercentile = percentile;
    currentCategory = category;
    currentGender = gender;

    // Always show the limit selector dropdown
    if (limitSelectorWrapper) {
        limitSelectorWrapper.classList.remove('hidden');
    }

    // Render count banner
    totalMatchesCount.textContent = finalMatches.length;

    // Dynamically update the header badge to show selected cities & branches
    const optionsBadge = document.getElementById('optionsBadge');
    if (optionsBadge) {
        const citiesText = selectedCities.length > 0 ? selectedCities.map(formatString).join(', ') : 'All Cities';
        const branchesText = selectedBranches.length > 0 ? selectedBranches.join(', ') : 'All Branches';
        optionsBadge.textContent = `${citiesText} | ${branchesText}`;
    }

    // Render TOP 10 Premier Colleges Grid
    renderTopCollegesGrid(topCollegesMatches, rank, percentile, category, gender);

    // Render Best Matching Colleges (with all matching branches as individual rows)
    renderBestMatches(finalMatches, rank, percentile, category, gender);
}

function getDynamicDescription(rank) {
    if (rank === null || rank === undefined) return "Below";
    const r = parseInt(rank);
    if (isNaN(r)) return "Below";
    if (r < 5000) return "Exellent";
    if (r < 10000) return "Great";
    if (r < 20000) return "Good";
    if (r < 40000) return "Average";
    return "Below";
}

// Helper to evaluate if a college's best branch cutoff lies in the "Perfect Match Zone"
function getMatchQuality(college, studentRank, studentPercentile) {
    if (!college.bestBranch || !college.bestBranch.evaluation) {
        return { inZone: false, distance: 9999999 };
    }
    
    const eval = college.bestBranch.evaluation;
    const cutoffRank = eval.cutoffRankUsed;
    const cutoffPerc = eval.cutoffPercUsed;
    
    let inZone = false;
    let distance = 0;
    
    if (cutoffRank !== null) {
        // Perfect Match Zone: [studentRank * 0.8, studentRank * 1.4]
        if (cutoffRank >= studentRank * 0.8 && cutoffRank <= studentRank * 1.4) {
            inZone = true;
        }
        distance = Math.abs(studentRank - cutoffRank);
    } else if (cutoffPerc !== null) {
        // Perfect Match Zone: [studentPercentile - 4, studentPercentile + 2]
        if (cutoffPerc >= studentPercentile - 4 && cutoffPerc <= studentPercentile + 2) {
            inZone = true;
        }
        distance = Math.abs(studentPercentile - cutoffPerc) * 1000; // scale distance to match rank magnitude
    }
    
    return {
        inZone,
        distance
    };
}

// Evaluate eligibility based on scores vs round cutoffs
function evaluateEligibility(studentRank, studentPercentile, cutoffs) {
    let eligible = false;
    let chance = 'DREAM'; 
    let chanceScore = 0; 
    let cutoffRankUsed = null;
    let cutoffPercUsed = null;

    for (let r = 0; r < 4; r++) {
        const rankIdx = r * 2;
        const percIdx = r * 2 + 1;
        
        if (rankIdx >= cutoffs.length) continue;

        const cutoffRank = cutoffs[rankIdx];
        const cutoffPerc = cutoffs[percIdx];

        if (cutoffRank !== null && cutoffRank !== undefined) {
            cutoffRankUsed = cutoffRank;
            cutoffPercUsed = cutoffPerc;
            
            if (studentRank <= cutoffRank) {
                eligible = true;
            }
            
            if (cutoffRank >= studentRank * 1.2) {
                chance = 'HIGHER CHANCES';
                chanceScore = Math.max(chanceScore, 3);
            } else if (cutoffRank >= studentRank * 0.8) {
                chance = 'MEDIUM CHANCES';
                chanceScore = Math.max(chanceScore, 2);
            } else {
                chance = 'DREAM';
                chanceScore = Math.max(chanceScore, 1);
            }
        } else if (cutoffPerc !== null && cutoffPerc !== undefined) {
            cutoffPercUsed = cutoffPerc;
            if (studentPercentile >= cutoffPerc) {
                eligible = true;
            }
            if (cutoffPerc <= studentPercentile - 2.0) {
                chance = 'HIGHER CHANCES';
                chanceScore = Math.max(chanceScore, 3);
            } else if (cutoffPerc <= studentPercentile + 2.0) {
                chance = 'MEDIUM CHANCES';
                chanceScore = Math.max(chanceScore, 2);
            } else {
                chance = 'DREAM';
                chanceScore = Math.max(chanceScore, 1);
            }
        }
    }

    return {
        eligible,
        chance,
        chanceScore,
        cutoffRankUsed,
        cutoffPercUsed
    };
}

// Render Top 10 Premier Colleges
function renderTopCollegesGrid(topColleges, studentRank, studentPercentile, category, gender) {
    const section = document.querySelector('.top-colleges-section');
    if (section) {
        if (topColleges.length === 0) {
            section.classList.add('hidden');
        } else {
            section.classList.remove('hidden');
        }
    }

    // Sort to keep user order
    topColleges.sort((a, b) => topCollegeCodes.indexOf(a.code) - topCollegeCodes.indexOf(b.code));

    topColleges.forEach(college => {
        const card = document.createElement('div');
        card.className = 'college-card';
        
        let branchHtml = '';
        let probHtml = '';
        
        if (college.bestBranch) {
            const chance = college.bestBranch.evaluation.chance; // 'HIGHER CHANCES', 'MEDIUM CHANCES', 'DREAM'
            let probClass = 'prob-dream';
            let probText = 'DREAM';
            let probIcon = '<i class="fa-solid fa-moon"></i>';
            
            if (chance === 'HIGHER CHANCES') {
                probClass = 'prob-high';
                probText = 'HIGHER CHANCES';
                probIcon = '<i class="fa-solid fa-circle-check"></i>';
            } else if (chance === 'MEDIUM CHANCES') {
                probClass = 'prob-medium';
                probText = 'MEDIUM CHANCES';
                probIcon = '<i class="fa-solid fa-circle-exclamation"></i>';
            }

            probHtml = `<span class="prob-pill ${probClass}">${probIcon} ${probText}</span>`;

            const cutoffRank = college.bestBranch.evaluation.cutoffRankUsed || '--';
            const cutoffPerc = college.bestBranch.evaluation.cutoffPercUsed ? `${college.bestBranch.evaluation.cutoffPercUsed}%` : '--';

            branchHtml = `
                <div class="card-match-details">
                    <p class="match-branch-label">Best Available Branch</p>
                    <p class="match-branch-name" title="${college.bestBranch.branchName}">${college.bestBranch.branchName}</p>
                    <div class="match-ranks-row">
                        <div class="rank-stat">
                            <span class="label">Cutoff Rank</span>
                            <span class="val">${cutoffRank}</span>
                        </div>
                        <div class="rank-stat">
                            <span class="label">Percentile</span>
                            <span class="val">${cutoffPerc}</span>
                        </div>
                    </div>
                </div>
            `;
        } else {
            probHtml = `<span class="prob-pill prob-dream"><i class="fa-solid fa-filter"></i> No Filters Match</span>`;
            branchHtml = `
                <div class="card-match-details" style="display:flex; align-items:center; justify-content:center; height:90px; text-align:center;">
                    <p class="color-text-muted" style="font-size:0.85rem; font-style:italic;">No branches matched your preferences at this premier college</p>
                </div>
            `;
        }

        card.innerHTML = `
            <div class="card-top">
                <div class="card-header-info">
                    <span class="college-code-badge">Code: ${college.code}</span>
                    ${probHtml}
                </div>
                <h3>${college.name}</h3>
                <span class="college-city"><i class="fa-solid fa-location-dot"></i> ${formatString(college.city)}</span>
            </div>
            
            ${branchHtml}
            
            <div class="card-action">
                <button class="btn-card-details" onclick="openCollegeDetails(${college.code}, ${studentRank}, ${studentPercentile}, '${category}', '${gender}')">
                    <span>Analyze All Branches</span> <i class="fa-solid fa-arrow-right"></i>
                </button>
            </div>
        `;
        
        topCollegesGrid.appendChild(card);
    });
}

// Render Best Matching Colleges in Table View
function renderBestMatches(matches, studentRank, studentPercentile, category, gender) {
    if (matches.length === 0) {
        matchesTableBody.innerHTML = `
            <tr>
                <td colspan="7" style="text-align: center; padding: 40px; color: var(--color-text-muted); font-style: italic;">
                    <i class="fa-solid fa-circle-exclamation" style="font-size: 2rem; color: var(--color-gold); margin-bottom: 10px; display: block;"></i>
                    No matches found. Try entering a higher merit rank or percentile, or select a different category/city/branch.
                </td>
            </tr>
        `;
        return;
    }

    matches.forEach(item => {
        const tr = document.createElement('tr');
        tr.onclick = () => openCollegeDetails(item.collegeCode, studentRank, studentPercentile, category, gender);
        
        const chance = item.chance; // 'HIGHER CHANCES', 'MEDIUM CHANCES', 'DREAM'
        let chanceClass = 'chance-dream';
        let chanceText = 'DREAM';
        
        if (chance === 'HIGHER CHANCES') {
            chanceClass = 'chance-high';
            chanceText = 'HIGHER CHANCES';
        } else if (chance === 'MEDIUM CHANCES') {
            chanceClass = 'chance-medium';
            chanceText = 'MEDIUM CHANCES';
        }

        const cutoffRank = item.maxRank || '--';
        const cutoffPerc = item.minPerc ? `${item.minPerc}%` : '--';

        tr.innerHTML = `
            <td><span class="chance-badge ${chanceClass}">${chanceText}</span></td>
            <td class="td-code">${item.collegeCode.toString().padStart(4, '0')}</td>
            <td class="td-name" title="${item.collegeName}">${item.collegeName}</td>
            <td>${formatString(item.city)}</td>
            <td class="td-branch" title="${item.branchName}">
                ${item.branchName}
                <div style="font-size: 0.72rem; color: var(--color-gold); font-weight: normal; margin-top: 3px;">Quota: ${item.cutoffCategory}</div>
            </td>
            <td class="td-cutoff">${cutoffRank} <span style="font-size: 0.75rem; color: var(--color-text-muted); font-weight: 500;">(${cutoffPerc})</span></td>
            <td><button class="btn-row-action"><i class="fa-solid fa-angle-right"></i></button></td>
        `;
        
        matchesTableBody.appendChild(tr);
    });
}

// Open College Details Modal
function openCollegeDetails(collegeCode, studentRank, studentPercentile, category, gender) {
    const college = collegesDb.find(c => c.c === collegeCode);
    if (!college) return;

    modalCollegeCode.textContent = `Code: ${college.c}`;
    modalCollegeName.textContent = college.n;
    modalCollegeCity.innerHTML = `<i class="fa-solid fa-location-dot"></i> ${formatString(college.ci)}`;
    
    modalStudentRank.textContent = studentRank.toLocaleString();
    modalStudentPercentile.textContent = `${studentPercentile}%`;
    modalSelectedCategory.textContent = `${category} (${gender === 'FEMALE' ? 'Female' : 'Male/General'})`;

    modalCutoffsBody.innerHTML = '';

    const categoryIndex = getCategoryIndex(category);

    // Retrieve active Seat Types from the DOM
    const selectedSeatTypes = [];
    if (seatHomeState && seatHomeState.checked) selectedSeatTypes.push('HomeState');
    if (seatOtherState && seatOtherState.checked) selectedSeatTypes.push('OtherState');
    if (seatStateLevel && seatStateLevel.checked) selectedSeatTypes.push('StateLevel');

    if (college.b && Array.isArray(college.b)) {
        college.b.forEach(branch => {
            Object.keys(branch.co || {}).forEach(catKey => {
                if (matchesExcelCategory(catKey, gender, categoryIndex, selectedSeatTypes)) {
                    const cutoffs = branch.co[catKey];
                    const evalResult = evaluateEligibility(studentRank, studentPercentile, cutoffs);
                    
                    const tr = document.createElement('tr');
                    
                    // Build cell for each round
                    let roundsHtml = '';
                    for (let r = 0; r < 4; r++) {
                        const rankIdx = r * 2;
                        const percIdx = r * 2 + 1;
                        
                        if (cutoffs && rankIdx < cutoffs.length && cutoffs[rankIdx] !== null) {
                            const cutoffRank = cutoffs[rankIdx];
                            const cutoffPerc = cutoffs[percIdx];
                            const qualifies = studentRank <= cutoffRank || studentPercentile >= cutoffPerc;
                            const qualifiesStyle = qualifies ? 'color: var(--color-success); font-weight: bold;' : '';
                            
                            roundsHtml += `
                                <td>
                                    <div class="round-cell" style="${qualifiesStyle}">
                                        <span class="rank">${cutoffRank}</span>
                                        <span class="perc">${cutoffPerc}%</span>
                                    </div>
                                </td>
                            `;
                        } else {
                            roundsHtml += `<td><span class="cutoff-null">No cutoff</span></td>`;
                        }
                    }

                    const qualifiesBranch = evalResult.eligible;
                    const rowStyle = qualifiesBranch ? 'background: rgba(46, 196, 182, 0.04);' : '';

                    tr.style = rowStyle;
                    tr.innerHTML = `
                        <td class="td-branch" style="font-weight: 600; font-size: 0.85rem;" title="${branch.n}">
                            ${branch.n}
                            <div style="font-size: 0.7rem; color: var(--color-gold); font-weight: normal; margin-top: 3px;">Quota: ${catKey}</div>
                        </td>
                        ${roundsHtml}
                    `;
                    
                    modalCutoffsBody.appendChild(tr);
                }
            });
        });
    }

    // Show modal
    collegeModal.classList.remove('hidden');
    document.body.style.overflow = 'hidden'; 
}

// Close Modal
function closeModal() {
    collegeModal.classList.add('hidden');
    document.body.style.overflow = ''; 
}

// Open Name Prompt Modal
function openNamePromptModal() {
    if (currentFinalMatches.length === 0 && currentTopCollegesMatches.length === 0) {
        alert("No recommended colleges available to download. Please run a prediction first.");
        return;
    }
    if (namePromptModal) {
        namePromptModal.classList.remove('hidden');
        document.body.style.overflow = 'hidden';
        if (studentNameInput) {
            studentNameInput.value = '';
            studentNameInput.focus();
        }
    }
}

// Close Name Prompt Modal
function closeNamePromptModal() {
    if (namePromptModal) {
        namePromptModal.classList.add('hidden');
        document.body.style.overflow = '';
    }
}

// Handle Name Prompt Submission
function handleNamePromptSubmit() {
    const studentName = studentNameInput.value.trim();
    if (!studentName) {
        alert("Please enter the student's name to personalize the report.");
        return;
    }
    closeNamePromptModal();
    downloadExcelReport(studentName);
}

// Helpers
function formatString(str) {
    if (!str) return '';
    return str.split(' ').map(word => {
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    }).join(' ');
}

// Estimate wrapped lines for Excel auto row height (approximates Excel wrap behavior)
function estimateWrappedLineCount(text, colWidthWch) {
    if (text === null || text === undefined || text === '') return 1;
    const str = String(text);
    const maxChars = Math.max(6, Math.floor(colWidthWch * 0.95));
    let totalLines = 0;

    str.split(/\r?\n/).forEach((line) => {
        if (!line) {
            totalLines += 1;
            return;
        }
        const words = line.split(/\s+/);
        let currentLen = 0;
        let lineCount = 1;
        words.forEach((word) => {
            const wordLen = word.length;
            if (wordLen > maxChars) {
                if (currentLen > 0) lineCount++;
                lineCount += Math.ceil(wordLen / maxChars);
                currentLen = 0;
            } else if (currentLen === 0) {
                currentLen = wordLen;
            } else if (currentLen + 1 + wordLen <= maxChars) {
                currentLen += 1 + wordLen;
            } else {
                lineCount++;
                currentLen = wordLen;
            }
        });
        totalLines += lineCount;
    });

    return Math.max(1, totalLines);
}

function buildExcelAutoRowHeights(data, colWidths, rowCount) {
    const rowHeights = [];
    const linePxByRow = (r) => {
        if (r === 0 || r === 5) return 16; // section headers
        if (r === 6) return 15; // table header
        if (r >= 7) return 14; // table body
        return 13;
    };
    const minPxByRow = (r) => {
        if (r === 0 || r === 5) return 22;
        if (r === 4) return 8;
        if (r === 6) return 24;
        if (r >= 7) return 20;
        return 18;
    };
    const maxPx = 140;

    for (let r = 0; r < rowCount; r++) {
        const dataRow = data[r] || [];
        let maxLines = 1;

        for (let c = 0; c < colWidths.length; c++) {
            const val = dataRow[c];
            if (val === null || val === undefined || val === '') continue;
            const wch = colWidths[c]?.wch || 10;
            maxLines = Math.max(maxLines, estimateWrappedLineCount(val, wch));
        }

        const hpx = Math.min(
            maxPx,
            Math.max(minPxByRow(r), Math.ceil(maxLines * linePxByRow(r)) + 6)
        );
        rowHeights.push({ hpx });
    }

    return rowHeights;
}

// Generate and Download Excel Report - Formatted to match MH CET OPTION ANALYSIS FORM layout
function downloadExcelReport(studentName) {
    if (typeof XLSX === 'undefined') {
        alert("Excel export library is loading, please try again in a moment.");
        return;
    }
    if (currentFinalMatches.length === 0 && currentTopCollegesMatches.length === 0) {
        alert("No recommended colleges available to download. Please run a prediction first.");
        return;
    }

    // ---- Palette ----
    const NAVY = '12243D';      // bold dark navy headings
    const TEAL = '0D8A6C';      // accent teal/green
    const BLACK = '111111';
    const WHITE = 'FFFFFF';
    const TBL_HEAD_BG = '12243D';
    const PURPLE = '6D28D9';    // DREAM
    const GREEN = '0F766E';     // HIGHER
    const AMBER = 'B45309';     // MEDIUM

    const FONT = 'Calibri';

    const today = new Date();
    const dateStr = today.getDate() + " " + today.toLocaleString('default', { month: 'short' }) + " " + today.getFullYear();
    const citiesStr = selectedCities.length > 0 ? selectedCities.map(formatString).join(", ") : "ALL";
    const branchesStr = selectedBranches.length > 0 ? selectedBranches.join(", ") : "ALL";

    let performance = "Average";
    if (currentPercentile >= 99) performance = "Top Tier";
    else if (currentPercentile >= 95) performance = "Excellent";
    else if (currentPercentile >= 90) performance = "Very Good";
    else if (currentPercentile >= 80) performance = "Good";

    // 9 columns: A..I
    // A: SR.NO/labels  B: Inst CODE  C: INSTITUTE  D: CITY  E: Branch Code/lbl
    // F: BRANCH NAME/val  G: CATEGORY/lbl  H: MAX Rank/val  I: STATUS
    const data = [];

    // Row 1: title + student name
    data.push([null, "MH CET OPTION  ANALYSIS FORM", null, "OF", studentName, null, null, null, null]);
    // Row 2: generated on + GENDER
    data.push([null, null, "Generated on: " + dateStr, null, "GENDER", currentGender, null, null, null]);
    // Row 3: CITIES + CATEGORY
    data.push(["CITIES", citiesStr, null, null, "CATEGORY", currentCategory, null, null, null]);
    // Row 4: BRANCH GRPS + RANK
    data.push(["BRANCH GRPS", branchesStr, null, null, "RANK", currentRank, null, null, null]);
    // Row 5: PERFORMANCE
    data.push([null, null, null, null, "PERFORMANCE", performance, null, null, null]);
    // Row 6: spacer
    data.push([null, null, null, null, null, null, null, null, null]);
    // Row 7: BASED ON + TOTAL COUNT
    data.push([null, null, "BASED ON YOUR SCORE THE BEST OPTIONS", null, null, null, "TOTAL COUNT", currentFinalMatches.length, null]);
    // Row 8: table headers
    data.push(["SR. NO", "Inst CODE", "INSTITUTE", "CITY", "Branch Code", "BRANCH NAME", "CATEGORY", "MAX Rank", "STATUS"]);

    // Data rows
    currentFinalMatches.forEach((item, idx) => {
        data.push([
            idx + 1,
            item.collegeCode.toString().padStart(5, '0'),
            item.collegeName,
            formatString(item.city),
            formatBranchCode(item.branchCode),
            item.branchName,
            item.cutoffCategory,
            item.maxRank !== null ? item.maxRank : '--',
            (item.chance || '').toString().toUpperCase().includes('DREAM') ? 'DREAM'
                : (item.chance || '').toString().toUpperCase().includes('HIGH') ? 'HIGHER'
                : (item.chance || '').toString().toUpperCase().includes('MED') ? 'MEDIUM'
                : (item.chance || 'SAFE').toString().toUpperCase()
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const rowCount = data.length;
    const colCount = 9;
    const getRef = (r, c) => XLSX.utils.encode_cell({ r, c });

    // Border helpers
    const noBorder = {};
    const greenTop = {
        top: { style: 'medium', color: { rgb: TEAL } }
    };
    const tblBorder = {
        top: { style: 'thin', color: { rgb: '94A3B8' } },
        bottom: { style: 'thin', color: { rgb: '94A3B8' } },
        left: { style: 'thin', color: { rgb: '94A3B8' } },
        right: { style: 'thin', color: { rgb: '94A3B8' } }
    };

    const setStyle = (r, c, s) => {
        const ref = getRef(r, c);
        if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        ws[ref].s = s;
    };

    // Ensure every cell up to colCount exists so borders/styles render
    for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
            const ref = getRef(r, c);
            if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        }
    }

    // ----- Row 1: Title + Student name + green top border -----
    // Top green border applied across row 1 cells
    for (let c = 0; c < colCount; c++) {
        setStyle(0, c, {
            font: { name: FONT, sz: 11, color: { rgb: BLACK } },
            alignment: { horizontal: 'left', vertical: 'center' },
            border: greenTop
        });
    }
    // B1: title
    setStyle(0, 1, {
        font: { name: FONT, sz: 12, bold: true, color: { rgb: NAVY } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: greenTop
    });
    // D1: "OF"
    setStyle(0, 3, {
        font: { name: FONT, sz: 11, italic: true, color: { rgb: TEAL } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: greenTop
    });
    // E1: student name
    setStyle(0, 4, {
        font: { name: FONT, sz: 12, bold: true, color: { rgb: TEAL } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: greenTop
    });

    // ----- Row 2: Generated on (C) + GENDER label (E) value (F) -----
    setStyle(1, 2, {
        font: { name: FONT, sz: 11, bold: true, color: { rgb: NAVY } },
        alignment: { horizontal: 'right', vertical: 'center' }
    });
    setStyle(1, 4, {
        font: { name: FONT, sz: 11, bold: true, color: { rgb: NAVY } },
        alignment: { horizontal: 'right', vertical: 'center' }
    });
    setStyle(1, 5, {
        font: { name: FONT, sz: 11, color: { rgb: BLACK } },
        alignment: { horizontal: 'left', vertical: 'center' }
    });

    // ----- Rows 3-5: CITIES / BRANCH GRPS (col A label, col B value) + CATEGORY/RANK/PERFORMANCE (col E label, col F value) -----
    const leftLabelRows = [2, 3]; // CITIES, BRANCH GRPS
    leftLabelRows.forEach(r => {
        setStyle(r, 0, {
            font: { name: FONT, sz: 11, bold: true, color: { rgb: NAVY } },
            alignment: { horizontal: 'left', vertical: 'center' }
        });
        setStyle(r, 1, {
            font: { name: FONT, sz: 11, color: { rgb: BLACK } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
        });
    });
    [2, 3, 4].forEach(r => {
        setStyle(r, 4, {
            font: { name: FONT, sz: 11, bold: true, color: { rgb: NAVY } },
            alignment: { horizontal: 'right', vertical: 'center' }
        });
        setStyle(r, 5, {
            font: { name: FONT, sz: 11, color: { rgb: BLACK } },
            alignment: { horizontal: 'left', vertical: 'center' }
        });
    });

    // ----- Row 7: BASED ON ... + TOTAL COUNT -----
    setStyle(6, 2, {
        font: { name: FONT, sz: 12, bold: true, color: { rgb: TEAL } },
        alignment: { horizontal: 'left', vertical: 'center' }
    });
    setStyle(6, 6, {
        font: { name: FONT, sz: 11, bold: true, color: { rgb: NAVY } },
        alignment: { horizontal: 'right', vertical: 'center' }
    });
    setStyle(6, 7, {
        font: { name: FONT, sz: 12, bold: true, color: { rgb: BLACK } },
        alignment: { horizontal: 'center', vertical: 'center' }
    });

    // ----- Row 8: Table header (dark navy bg, white bold, centered) -----
    for (let c = 0; c < colCount; c++) {
        setStyle(7, c, {
            font: { name: FONT, sz: 11, bold: true, color: { rgb: WHITE } },
            fill: { fgColor: { rgb: TBL_HEAD_BG } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
                top: { style: 'thin', color: { rgb: TBL_HEAD_BG } },
                bottom: { style: 'thin', color: { rgb: TBL_HEAD_BG } },
                left: { style: 'thin', color: { rgb: '334155' } },
                right: { style: 'thin', color: { rgb: '334155' } }
            }
        });
    }

    // ----- Data rows -----
    for (let r = 8; r < rowCount; r++) {
        const zebra = (r - 8) % 2 === 1 ? 'F8FAFC' : null;
        for (let c = 0; c < colCount; c++) {
            const ref = getRef(r, c);
            const cell = ws[ref];
            // keep codes as text with leading zeros
            if (c === 1 || c === 4) {
                cell.t = 's';
                cell.z = '@';
            }
            let align = 'center';
            if (c === 2 || c === 5) align = 'left';
            const style = {
                font: { name: FONT, sz: 10, color: { rgb: BLACK } },
                alignment: { horizontal: align, vertical: 'center', wrapText: true },
                border: tblBorder
            };
            if (zebra) style.fill = { fgColor: { rgb: zebra } };
            cell.s = style;

            // STATUS column color coding
            if (c === 8) {
                const v = (cell.v || '').toString().toUpperCase();
                if (v === 'DREAM') {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: PURPLE } };
                    cell.s.fill = { fgColor: { rgb: 'F5F3FF' } };
                } else if (v === 'HIGHER') {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: GREEN } };
                    cell.s.fill = { fgColor: { rgb: 'F0FDF4' } };
                } else if (v === 'MEDIUM') {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: AMBER } };
                    cell.s.fill = { fgColor: { rgb: 'FEF3C7' } };
                } else {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: NAVY } };
                }
            }
        }
    }

    // ----- Merges -----
    ws['!merges'] = [
        // Title spans B1:C1
        { s: { r: 0, c: 1 }, e: { r: 0, c: 2 } },
        // Student name spans E1:I1
        { s: { r: 0, c: 4 }, e: { r: 0, c: 8 } },
        // Generated on spans C2:D2 (right aligned within)
        { s: { r: 1, c: 2 }, e: { r: 1, c: 3 } },
        // CITIES value spans B3:D3
        { s: { r: 2, c: 1 }, e: { r: 2, c: 3 } },
        // BRANCH GRPS value spans B4:D4
        { s: { r: 3, c: 1 }, e: { r: 3, c: 3 } },
        // BASED ON spans C7:F7
        { s: { r: 6, c: 2 }, e: { r: 6, c: 5 } }
    ];

    // ----- Column widths (matching image) -----
    ws['!cols'] = [
        { wch: 14 },  // A SR.NO / label
        { wch: 12 },  // B Inst CODE
        { wch: 38 },  // C INSTITUTE
        { wch: 12 },  // D CITY
        { wch: 14 },  // E Branch Code / label
        { wch: 32 },  // F BRANCH NAME / value
        { wch: 14 },  // G CATEGORY
        { wch: 11 },  // H MAX Rank
        { wch: 12 }   // I STATUS
    ];

    // ----- Row heights -----
    const rows = [];
    rows[0] = { hpx: 26 };  // title
    rows[1] = { hpx: 22 };
    rows[2] = { hpx: 22 };
    rows[3] = { hpx: 22 };
    rows[4] = { hpx: 22 };
    rows[5] = { hpx: 10 };  // spacer
    rows[6] = { hpx: 24 };  // BASED ON
    rows[7] = { hpx: 30 };  // table header
    for (let r = 8; r < rowCount; r++) rows[r] = { hpx: 32 };
    ws['!rows'] = rows;

    // Print setup
    ws['!margins'] = { left: 0.3, right: 0.3, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 };
    ws['!pageSetup'] = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    ws['!sheetProperties'] = { pageSetUpPr: { fitToPage: true } };
    ws['!sheetViews'] = [{ showGridLines: false }];
    ws['!views'] = [{ showGridLines: false }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Option Analysis Form");

    const sanitizedName = studentName.replace(/[^a-zA-Z0-9\s]/g, '').trim().replace(/\s+/g, '_') || 'Student';
    XLSX.writeFile(wb, `AME_College_Recommendations_${sanitizedName}_${currentRank}.xlsx`);
}

function formatBranchCode(branchCode) {
    const code = branchCode ? branchCode.toString().trim() : '';
    const hasLetterSuffix = /[A-Za-z]$/.test(code);
    return code.padStart(hasLetterSuffix ? 11 : 10, '0');
}

function getRankExpansionStep(maxRank) {
    return Number(maxRank) >= 10000 ? 10000 : 1000;
}

function compareTopCollegeBranches(a, b) {
    const aBranchOrder = selectedBranches.length > 0 ? selectedBranches.indexOf(a.branchGroup) : -1;
    const bBranchOrder = selectedBranches.length > 0 ? selectedBranches.indexOf(b.branchGroup) : -1;
    if (aBranchOrder !== bBranchOrder) {
        if (aBranchOrder === -1) return 1;
        if (bBranchOrder === -1) return -1;
        return aBranchOrder - bBranchOrder;
    }

    const aRank = a.maxRank !== null ? a.maxRank : 999999;
    const bRank = b.maxRank !== null ? b.maxRank : 999999;
    if (aRank !== bRank) {
        return aRank - bRank;
    }

    return a.branchName.localeCompare(b.branchName);
}

function hasEnoughBalancedRecommendations(matches, limit) {
    if (limit <= 0) return true;
    const mediumCount = matches.filter(item => item.chance === 'MEDIUM CHANCES').length;
    const higherCount = matches.filter(item => item.chance === 'HIGHER CHANCES').length;
    return matches.length >= limit && mediumCount > 0 && higherCount > 0;
}

function buildBalancedRecommendations(matches, limit) {
    if (limit <= 0) return [];
    const sortedMatches = [...matches].sort((a, b) => {
        const aVal = a.maxRank !== null ? a.maxRank : 999999;
        const bVal = b.maxRank !== null ? b.maxRank : 999999;
        return aVal - bVal;
    });
    const mediumMatches = sortedMatches.filter(item => item.chance === 'MEDIUM CHANCES');
    const higherMatches = sortedMatches.filter(item => item.chance === 'HIGHER CHANCES');
    const higherTarget = higherMatches.length > 0 ? Math.min(higherMatches.length, Math.floor(limit / 2)) : 0;
    const mediumTarget = Math.min(mediumMatches.length, limit - higherTarget);
    const selected = [
        ...mediumMatches.slice(0, mediumTarget),
        ...higherMatches.slice(0, limit - mediumTarget)
    ];

    if (selected.length < limit) {
        const selectedKeys = new Set(selected.map(item => `${item.collegeCode}-${item.branchCode}-${item.cutoffCategory}`));
        sortedMatches.forEach(item => {
            const key = `${item.collegeCode}-${item.branchCode}-${item.cutoffCategory}`;
            if (selected.length < limit && !selectedKeys.has(key)) {
                selected.push(item);
                selectedKeys.add(key);
            }
        });
    }

    return selected.slice(0, limit);
}

function mergeTopAndTargetRecommendations(topRecords, targetRecords, limit) {
    const merged = [];
    const seen = new Set();
    [...topRecords, ...targetRecords].forEach(item => {
        const key = `${item.collegeCode}-${item.branchCode}-${item.cutoffCategory}`;
        if (!seen.has(key) && merged.length < limit) {
            merged.push(item);
            seen.add(key);
        }
    });
    return merged;
}

function escapeHtml(value) {
    return value === null || value === undefined ? '' : value.toString()
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function printOptionForm() {
    if (currentFinalMatches.length === 0) {
        alert("No recommended colleges available to print. Please run a prediction first.");
        return;
    }

    const printRows = currentFinalMatches.slice(0, Math.max(30, Math.min(currentFinalMatches.length, 40)));
    const today = new Date();
    const dateStr = today.getDate() + " " + today.toLocaleString('default', { month: 'short' }) + " " + today.getFullYear();
    const citiesStr = selectedCities.length > 0 ? selectedCities.map(formatString).join(", ") : "ALL";
    const branchesStr = selectedBranches.length > 0 ? selectedBranches.join(", ") : "ALL";
    const logoUrl = `${window.location.origin}${window.location.pathname.replace(/[^/]*$/, '')}logo.png`;

    const rowsHtml = printRows.map((item, index) => `
        <tr>
            <td>${index + 1}</td>
            <td>${escapeHtml(item.collegeCode.toString().padStart(5, '0'))}</td>
            <td>${escapeHtml(item.collegeName)}</td>
            <td>${escapeHtml(formatString(item.city))}</td>
            <td>${escapeHtml(formatBranchCode(item.branchCode))}</td>
            <td>${escapeHtml(item.branchName)}</td>
            <td>${escapeHtml(item.cutoffCategory)}</td>
            <td>${escapeHtml(item.maxRank !== null ? item.maxRank : '--')}</td>
            <td>${escapeHtml(item.chance)}</td>
        </tr>
    `).join('');

    const printHtml = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>AME Option Form</title>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <style>
                @page { size: A4 landscape; margin: 6mm; }
                * { box-sizing: border-box; }
                body { font-family: 'Segoe UI', Arial, sans-serif; color: #111827; margin: 0; }
                .sheet { width: 100%; position: relative; min-height: 190mm; }
                .sheet::before {
                    content: "";
                    position: fixed;
                    inset: 0;
                    background: url("${logoUrl}") center 52% / 140mm auto no-repeat;
                    opacity: 0.06;
                    pointer-events: none;
                    z-index: -1;
                }
                .print-header {
                    background: linear-gradient(180deg, #0a1628 0%, #061018 100%);
                    color: #f8f9fa;
                    text-align: center;
                    padding: 6px 10px 7px;
                    margin-bottom: 5px;
                    border-radius: 4px;
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                }
                .print-header-logo-wrap {
                    width: 48px;
                    height: 48px;
                    margin: 0 auto 4px;
                    border-radius: 50%;
                    padding: 2px;
                    background: linear-gradient(135deg, #e2c044 0%, #bfa02c 100%);
                    box-shadow: 0 0 14px rgba(226, 192, 68, 0.4);
                }
                .print-header-logo {
                    width: 100%;
                    height: 100%;
                    border-radius: 50%;
                    object-fit: cover;
                    background: #fff;
                    display: block;
                }
                .print-brand-title {
                    font-size: 13px;
                    font-weight: 800;
                    letter-spacing: 1.2px;
                    margin: 0;
                    line-height: 1.15;
                }
                .print-brand-title .brand-silver { color: #e8eef4; }
                .print-brand-title .brand-gold { color: #e2c044; }
                .print-brand-subtitle {
                    font-size: 8.5px;
                    color: #9bb1c9;
                    margin: 2px 0 0;
                    font-weight: 600;
                    letter-spacing: 0.3px;
                }
                .print-brand-subtitle i {
                    color: #e2c044;
                    margin-right: 4px;
                    font-size: 8px;
                }
                .print-header-contact {
                    font-size: 7.5px;
                    color: #e2c044;
                    margin-top: 3px;
                    font-weight: 600;
                }
                .print-header-contact i { margin-right: 3px; }
                .meta { display: grid; grid-template-columns: repeat(4, 1fr); gap: 4px 10px; font-size: 9px; margin-bottom: 6px; }
                .meta div { border: 1px solid #cbd5e1; padding: 3px 5px; min-height: 18px; }
                .meta strong { color: #0f172a; }
                table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 8.2px; }
                th, td { border: 1px solid #94a3b8; padding: 2px 3px; vertical-align: middle; line-height: 1.12; }
                th { background: #1e293b; color: #fff; font-weight: 700; text-align: center; }
                td:nth-child(1), td:nth-child(2), td:nth-child(5), td:nth-child(7), td:nth-child(8), td:nth-child(9) { text-align: center; }
                th:nth-child(1) { width: 5%; }
                th:nth-child(2) { width: 8%; }
                th:nth-child(3) { width: 28%; }
                th:nth-child(4) { width: 8%; }
                th:nth-child(5) { width: 10%; }
                th:nth-child(6) { width: 22%; }
                th:nth-child(7) { width: 7%; }
                th:nth-child(8) { width: 6%; }
                th:nth-child(9) { width: 6%; }
                tr { height: 16px; page-break-inside: avoid; }
                .footer-note { margin-top: 4px; font-size: 8px; color: #64748b; }
                @media print {
                    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                }
            </style>
        </head>
        <body>
            <div class="sheet">
                <header class="print-header">
                    <div class="print-header-logo-wrap">
                        <img src="${logoUrl}" alt="AME Engineering Latur" class="print-header-logo">
                    </div>
                    <h1 class="print-brand-title">
                        <span class="brand-silver">AME ENGINEERING </span><span class="brand-gold">LATUR</span>
                    </h1>
                    <p class="print-brand-subtitle"><i class="fa-solid fa-graduation-cap"></i> Admissions Made Easy</p>
                    <p class="print-header-contact"><i class="fa-solid fa-phone-volume"></i> Helpline: +91 94226 11661</p>
                </header>
                <div class="meta">
                    <div><strong>Date:</strong> ${escapeHtml(dateStr)}</div>
                    <div><strong>Rank:</strong> ${escapeHtml(currentRank)}</div>
                    <div><strong>Percentile:</strong> ${escapeHtml(currentPercentile)}%</div>
                    <div><strong>Category:</strong> ${escapeHtml(currentCategory)}</div>
                    <div><strong>Gender:</strong> ${escapeHtml(currentGender)}</div>
                    <div><strong>Cities:</strong> ${escapeHtml(citiesStr)}</div>
                    <div><strong>Branches:</strong> ${escapeHtml(branchesStr)}</div>
                    <div><strong>Total Options:</strong> ${escapeHtml(printRows.length)}</div>
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>SR</th>
                            <th>Inst Code</th>
                            <th>Institute</th>
                            <th>City</th>
                            <th>Branch Code</th>
                            <th>Branch Name</th>
                            <th>Category</th>
                            <th>Max Rank</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
                <div class="footer-note">Print setting: A4 paper, landscape orientation. The layout is compacted to fit at least 30 options on one page.</div>
            </div>
            <script>
                window.onload = function () {
                    window.focus();
                    window.print();
                };
            </script>
        </body>
        </html>
    `;

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        alert("Please allow pop-ups to print the option form.");
        return;
    }
    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
}
