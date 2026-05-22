// AME Engineering Latur - Application Logic

let collegesDb = [];
let rankPercMapping = [];
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
const loader = document.getElementById('loader');
const resultsSection = document.getElementById('resultsSection');
const totalMatchesCount = document.getElementById('totalMatchesCount');
const topCollegesGrid = document.getElementById('topCollegesGrid');
const matchesTableBody = document.getElementById('matchesTableBody');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');

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
        
        // Uncheck all checkboxes in dropdowns
        const checkboxes = predictorForm.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(chk => {
            chk.checked = false;
        });
        
        // Remove 'selected' class from option divs
        const optionDivs = predictorForm.querySelectorAll('.multiselect-option');
        optionDivs.forEach(div => {
            div.classList.remove('selected');
        });
        
        // Reset Trigger UIs (restore placeholders, hide tags)
        updateCityTriggerUI();
        updateBranchTriggerUI();
        
        // Hide results and scroll to top smoothly
        resultsSection.classList.add('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    // Real-time two-way Rank and Percentile Auto-Fetching
    if (studentRankInput && studentPercentileInput) {
        studentRankInput.addEventListener('input', () => {
            // Only auto-fill if the user is typing/focusing on studentRankInput
            if (document.activeElement !== studentRankInput) return;
            
            const val = studentRankInput.value.trim();
            if (!val) {
                studentPercentileInput.value = '';
                return;
            }
            
            const rank = parseInt(val, 10);
            if (isNaN(rank) || rank < 0) {
                return;
            }
            
            const perc = findClosestPercentileForRank(rank);
            if (perc !== null) {
                studentPercentileInput.value = perc;
            }
        });

        studentPercentileInput.addEventListener('input', () => {
            // Only auto-fill if the user is typing/focusing on studentPercentileInput
            if (document.activeElement !== studentPercentileInput) return;
            
            const val = studentPercentileInput.value.trim();
            if (!val) {
                studentRankInput.value = '';
                return;
            }
            
            const percentile = parseFloat(val);
            if (isNaN(percentile) || percentile < 0 || percentile > 100) {
                return;
            }
            
            const rank = findClosestRankForPercentile(percentile);
            if (rank !== null) {
                studentRankInput.value = rank;
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
        console.log(`Loaded ${rankPercMapping.length} rank-percentile mapping records successfully.`);
    } catch (error) {
        console.error("Failed to load rank-percentile mapping:", error);
    }
}

// Find closest percentile for a given rank using O(log N) binary search
function findClosestPercentileForRank(targetRank) {
    if (!rankPercMapping || rankPercMapping.length === 0) return null;
    let low = 0;
    let high = rankPercMapping.length - 1;
    let closestIndex = 0;
    let minDiff = Infinity;
    
    while (low <= high) {
        let mid = Math.floor((low + high) / 2);
        let currentRank = rankPercMapping[mid][0];
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
    return rankPercMapping[closestIndex][1];
}

// Find closest rank for a given percentile using robust scan
function findClosestRankForPercentile(targetPercentile) {
    if (!rankPercMapping || rankPercMapping.length === 0) return null;
    let closestRank = null;
    let minDiff = Infinity;
    
    for (let i = 0; i < rankPercMapping.length; i++) {
        let diff = Math.abs(rankPercMapping[i][1] - targetPercentile);
        if (diff < minDiff) {
            minDiff = diff;
            closestRank = rankPercMapping[i][0];
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

function matchesExcelCategory(rowCategory, gender, categoryIndex) {
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
    const matchedBranchesList = [];
    const topCollegesMatches = [];
    
    // Clear display grids/tables
    topCollegesGrid.innerHTML = '';
    matchesTableBody.innerHTML = '';

    const categoryIndex = getCategoryIndex(category);

    collegesDb.forEach(college => {
        const isTopCollege = topCollegeCodes.includes(college.c);
        
        // 1. GENDER RULE: If male, filter out women-only colleges!
        if (gender === 'MALE' && isWomenOnlyCollege(college)) {
            return;
        }

        // 2. CITY PREFERENCE RULE: Check multi-select filters
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
                    if (matchesExcelCategory(catKey, gender, categoryIndex)) {
                        const curCutoffs = branch.co[catKey];
                        
                        // Locate the round yielding the MAXIMUM rank (easiest round) - directly read from index 8/9 if present
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
                // Find closest branch fallback
                let closestBranch = null;
                let closestDistance = Infinity;
                let closestCutoffRank = null;
                let closestCategory = null;

                if (college.b && Array.isArray(college.b)) {
                    college.b.forEach(branch => {
                        Object.keys(branch.co || {}).forEach(catKey => {
                            if (matchesExcelCategory(catKey, gender, categoryIndex)) {
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
                        helper: 1
                    };
                }
            }

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

    // Partition matched candidates into dreams and targets/backups
    const dreamCandidates = [];
    const targetCandidates = [];

    matchedBranchesList.forEach(item => {
        if (item.chance === 'DREAM') {
            dreamCandidates.push(item);
        } else {
            targetCandidates.push(item);
        }
    });

    // Dynamic threshold search loop for target / safe options
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
        
        if (selectedTargets.length >= 30 || threshold > 150000) {
            break;
        }
        threshold += 1000;
    }

    // Double Sorting Rule: Status Helper ascending, then Max Rank ascending
    const sortFunc = (a, b) => {
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
        return bPerc - aPerc; // higher percentile first as secondary fallback
    };

    // Sort dreams and take top 5 closest dreams
    dreamCandidates.sort(sortFunc);
    const limitedDreams = dreamCandidates.slice(0, 5);

    // Sort selected targets
    selectedTargets.sort(sortFunc);

    // Combine dreams and targets, capped at exactly 45 total options
    const finalMatches = [...limitedDreams, ...selectedTargets].slice(0, 45);

    // Save current match data for Excel download
    currentTopCollegesMatches = topCollegesMatches;
    currentFinalMatches = finalMatches;
    currentRank = rank;
    currentPercentile = percentile;
    currentCategory = category;
    currentGender = gender;

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
            <td class="td-branch" title="${item.branchName}">${item.branchName}</td>
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

    if (college.b && Array.isArray(college.b)) {
        college.b.forEach(branch => {
            // Find the best category mapping cutoff for this branch
            let targetCategory = null;
            let cutoffs = null;
            let bestEval = null;

            Object.keys(branch.co || {}).forEach(catKey => {
                if (matchesExcelCategory(catKey, gender, categoryIndex)) {
                    const curCutoffs = branch.co[catKey];
                    const curEval = evaluateEligibility(studentRank, studentPercentile, curCutoffs);
                    
                    if (!bestEval || curEval.chanceScore > bestEval.chanceScore) {
                        bestEval = curEval;
                        targetCategory = catKey;
                        cutoffs = curCutoffs;
                    }
                }
            });

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

            // Check if student qualifies in ANY round for this branch
            let qualifiesBranch = false;
            if (cutoffs) {
                qualifiesBranch = bestEval.eligible;
            }

            const rowStyle = qualifiesBranch ? 'background: rgba(46, 196, 182, 0.04);' : '';

            tr.style = rowStyle;
            tr.innerHTML = `
                <td class="td-branch" style="font-weight: 600; font-size: 0.85rem;" title="${branch.n}">
                    ${branch.n}
                    ${targetCategory ? `<div style="font-size: 0.7rem; color: var(--color-gold); font-weight: normal; margin-top: 3px;">Quota: ${targetCategory}</div>` : ''}
                </td>
                ${roundsHtml}
            `;
            
            modalCutoffsBody.appendChild(tr);
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

// Generate and Download Excel Report using SheetJS
function downloadExcelReport(studentName) {
    if (typeof XLSX === 'undefined') {
        alert("Excel export library is loading, please try again in a moment.");
        return;
    }

    if (currentFinalMatches.length === 0 && currentTopCollegesMatches.length === 0) {
        alert("No recommended colleges available to download. Please run a prediction first.");
        return;
    }

    const wb = XLSX.utils.book_new();
    const data = [];

    // Title Block resembling 'Sorted view'
    data.push([null, null, "MH CET OPTION  ANALYSIS FORM", "OF", studentName.toUpperCase()]);
    
    const today = new Date();
    const dateStr = today.getDate() + " " + today.toLocaleString('default', { month: 'short' }) + " " + today.getFullYear();
    data.push([null, null, "Generated on: " + dateStr, null, "GENDER", currentGender.toUpperCase()]);
    
    const citiesStr = selectedCities.length > 0 ? selectedCities.map(formatString).join(", ") : "ALL";
    data.push(["CITIES", null, citiesStr, null, "CATEGORY", currentCategory.toUpperCase()]);
    
    const branchesStr = selectedBranches.length > 0 ? selectedBranches.join(", ") : "ALL";
    data.push(["BRANCH GRPS", null, branchesStr, null, "RANK", currentRank]);
    
    // Performance label: Top Tier, Excellent, Very Good, Good, Average
    let performance = "Average";
    if (currentPercentile >= 99) performance = "Top Tier";
    else if (currentPercentile >= 95) performance = "Excellent";
    else if (currentPercentile >= 90) performance = "Very Good";
    else if (currentPercentile >= 80) performance = "Good";
    data.push([null, null, null, null, "PERFORMANCE", performance]);
    
    data.push([]); // blank row 6
    
    data.push([null, null, "BASED ON YOUR SCORE THE BEST OPTIONS", null, "TOTAL CONUNT", currentFinalMatches.length]);
    
    // Table Headers (Row 8 in Sorted view)
    data.push([
        "SR. NO",
        "Inst CODE",
        "INSTITUTE",
        "CITY",
        "Branch Code",
        "BRANCH NAME",
        "CATEGORY",
        "MAX Rank",
        "STATUS"
    ]);

    // Data rows
    let srNo = 1;
    currentFinalMatches.forEach(item => {
        const cutoffRank = item.maxRank !== null ? item.maxRank : '--';
        data.push([
            srNo++,
            item.collegeCode.toString().padStart(4, '0'),
            item.collegeName,
            formatString(item.city),
            item.branchCode,
            item.branchName,
            item.cutoffCategory,
            cutoffRank,
            item.chance
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);

    // Apply Column Widths
    const colWidths = [
        { wch: 10 },  // SR. NO
        { wch: 12 },  // Inst CODE
        { wch: 60 },  // INSTITUTE
        { wch: 15 },  // CITY
        { wch: 15 },  // Branch Code
        { wch: 45 },  // BRANCH NAME
        { wch: 15 },  // CATEGORY
        { wch: 12 },  // MAX Rank
        { wch: 12 }   // STATUS
    ];
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, "Option Analysis Form");

    // Trigger download
    const sanitizedName = studentName.replace(/[^a-zA-Z0-9\s]/g, '').trim().replace(/\s+/g, '_');
    XLSX.writeFile(wb, `AME_College_Recommendations_${sanitizedName}_${currentRank}.xlsx`);
}
