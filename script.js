// Official SDU Data Analytics in Accounting Question Bank
const questionBank = [
    {
        id: 1,
        category: "excel-pwc",
        type: "excel",
        title: "Excel Part 1 (PwC): NYC Film Permit XLOOKUP / VLOOKUP Audit",
        accountingContext: "Auditors need to cross-reference event permit IDs from 'NYC_Film_Permits.txt' with master billing sheets to confirm approved production categories and verify revenue classification.",
        technicalPrompt: "Write an Excel XLOOKUP formula to lookup the Permit ID in cell A2 from the 'NYC_Film_Permits' sheet (Column A) and return the 'Category' column (Column D).",
        solution: "=XLOOKUP(A2, NYC_Film_Permits!A:A, NYC_Film_Permits!D:D, \"Not Found\")",
        explanation: {
            accountingRationale: "Reconciling operational permit data against financial records ensures revenue recognition compliance and identifies unbilled permit fees.",
            stepByStep: [
                "1. Search Key (A2): Unique Permit ID on the current worksheet.",
                "2. Lookup Range (NYC_Film_Permits!A:A): Contains source IDs in the external sheet.",
                "3. Return Range (NYC_Film_Permits!D:D): Specifies the 'Category' column to be retrieved.",
                "4. If Not Found Value (\"Not Found\"): Prevents #N/A errors and flags unverified records for audit."
            ],
            examTip: "XLOOKUP is preferred over VLOOKUP because it does not break when columns are reordered and defaults to exact matching."
        },
        keywords: ["XLOOKUP", "VLOOKUP"]
    },
    {
        id: 2,
        category: "excel-etl",
        type: "excel",
        title: "Excel Part 2 (ETL Cases 1-4): Parsing Text Identifiers",
        accountingContext: "Raw transaction files imported into the accounting system contain combined Vendor Name and Location ID in a single field (e.g., 'ACME_NY101'). You must extract the Location ID for localized expense allocation.",
        technicalPrompt: "Write an Excel formula using TEXTAFTER or RIGHT/MID functions to extract the substring after the underscore '_' character in cell A2.",
        solution: "=TEXTAFTER(A2, \"_\")\nOr: =RIGHT(A2, LEN(A2) - FIND(\"_\", A2))",
        explanation: {
            accountingRationale: "Extracting cost-center codes from unstructured text fields allows automated cost accounting and managerial segment reporting.",
            stepByStep: [
                "1. Modern Method (=TEXTAFTER(A2, \"_\")): Locates '_' and returns all characters to the right.",
                "2. Legacy Method (=RIGHT(...)): Uses FIND(\"_\", A2) to locate character position, subtracts from total length, and extracts substring."
            ],
            examTip: "If TEXTAFTER is unavailable in older Excel versions, use the LEN + FIND combination."
        },
        keywords: ["TEXTAFTER", "RIGHT", "MID", "FIND"]
    },
    {
        id: 3,
        category: "excel-integrateco",
        type: "excel",
        title: "Excel Part 3 (IntegrateCo): Multi-Criteria Payroll SUMIFS Audit",
        accountingContext: "During the IntegrateCo merger audit, management wants to calculate total 2016 gross salary paid to employees in the 'Executive' job code (JobCode 401) working at the 'Chicago' facility.",
        technicalPrompt: "Write a SUMIFS formula to sum 'GrossPay' (Column E) where 'JobCode' (Column C) equals 401 and 'Location' (Column D) equals 'Chicago'.",
        solution: "=SUMIFS(Payroll2016!E:E, Payroll2016!C:C, 401, Payroll2016!D:D, \"Chicago\")",
        explanation: {
            accountingRationale: "Post-merger payroll audits verify whether executive compensation aligns with merger contracts and highlights regional pay variances.",
            stepByStep: [
                "1. Sum_Range (Payroll2016!E:E): Column containing numerical values to sum (Gross Pay).",
                "2. Criteria_Range1 & Criteria1 (Column C, 401): Restricts sum to Executive job codes.",
                "3. Criteria_Range2 & Criteria2 (Column D, \"Chicago\"): Restricts sum to the Chicago branch."
            ],
            examTip: "In SUMIFS, the sum range comes FIRST (unlike SUMIF where the sum range comes last)."
        },
        keywords: ["SUMIFS"]
    },
    {
        id: 4,
        category: "excel-regression",
        type: "excel",
        title: "Excel Part 4 (Regression): Restaurant Sales Determinants",
        accountingContext: "An analytical review model evaluates restaurant sales based on store square footage and advertising expenditures. Interpret the regression output: R-Square = 0.84, Advertising P-value = 0.002, Square Footage P-value = 0.42.",
        technicalPrompt: "Which independent variable is statistically significant at alpha = 0.05, and what proportion of total revenue variance is explained by the model?",
        solution: "1. Advertising is statistically significant (P-value 0.002 < 0.05).\n2. Square Footage is NOT significant (P-value 0.42 > 0.05).\n3. R-Squared = 0.84 (84% of revenue variance explained).",
        explanation: {
            accountingRationale: "Regression analysis helps managerial accountants determine which operational expenses statistically drive top-line revenue.",
            stepByStep: [
                "1. Statistical Significance Rule: Variable is significant if P-value < alpha (0.05). Since 0.002 < 0.05, Advertising is significant.",
                "2. Insignificance Rule: Square Footage P-value (0.42) > 0.05, meaning square footage is not a statistically significant driver in this model.",
                "3. Goodness of Fit (R-Squared): R^2 = 0.84 indicates 84% of total sales variation is explained by the independent variables."
            ],
            examTip: "Always compare P-values to 0.05. If P-value < 0.05, reject the null hypothesis."
        },
        keywords: ["SIGNIFICANT", "84%", "P-VALUE"]
    },
    {
        id: 5,
        category: "sql-cs50",
        type: "sql",
        title: "SQL Part 1 (CS50 Fiftyville): Multi-Table Forensic Join",
        accountingContext: "Internal auditors investigating asset misappropriation need to list all suspects who made a bank transaction on 'Leggett Street' on July 28, 2023, cross-referenced with phone call logs lasting under 60 seconds.",
        technicalPrompt: "Write an SQL query joining 'people', 'bank_accounts', 'atm_transactions', and 'phone_calls' to find names of suspects meeting both conditions.",
        solution: "SELECT p.name \nFROM people p\nJOIN bank_accounts b ON p.id = b.person_id\nJOIN atm_transactions atm ON b.account_number = atm.account_number\nWHERE atm.year = 2023 AND atm.month = 7 AND atm.day = 28 \n  AND atm.atm_location = 'Leggett Street'\n  AND p.phone_number IN (\n      SELECT caller FROM phone_calls \n      WHERE year = 2023 AND month = 7 AND day = 28 AND duration < 60\n  );",
        explanation: {
            accountingRationale: "Forensic accountants combine bank account logs with external metadata to establish evidence chains in fraud investigations.",
            stepByStep: [
                "1. JOIN Chain: Connects 'people' (p) -> 'bank_accounts' (b) -> 'atm_transactions' (atm) via relational keys.",
                "2. Outer WHERE Clause: Filters ATM withdrawals on Leggett Street on July 28, 2023.",
                "3. Subquery (IN): Filters phone numbers to callers making short calls (< 60s) on the same date."
            ],
            examTip: "Verify foreign keys match across tables: p.id = b.person_id and b.account_number = atm.account_number."
        },
        keywords: ["SELECT", "JOIN", "WHERE"]
    },
    {
        id: 6,
        category: "sql-pwc",
        type: "sql",
        title: "SQL Part 2 (PwC Chinook): Aggregate Customer Revenue Analysis",
        accountingContext: "Management requires a list of top commercial clients whose total aggregate purchases exceed $45 across all historical billing periods to target for loyalty incentives.",
        technicalPrompt: "Write a SQL query using GROUP BY and HAVING on the 'Invoice' table in Chinook DB to aggregate Total spend by CustomerId where SUM(Total) > 45.",
        solution: "SELECT CustomerId, SUM(Total) AS TotalSpent\nFROM Invoice\nGROUP BY CustomerId\nHAVING SUM(Total) > 45\nORDER BY TotalSpent DESC;",
        explanation: {
            accountingRationale: "Categorizing accounts by cumulative spend helps managerial accountants evaluate customer lifetime value (CLV) and revenue concentration.",
            stepByStep: [
                "1. GROUP BY CustomerId: Aggregates individual invoices by customer.",
                "2. SUM(Total): Calculates total aggregate spend per customer.",
                "3. HAVING SUM(Total) > 45: Filters aggregated sums (WHERE cannot filter aggregate functions like SUM)."
            ],
            examTip: "Use WHERE to filter individual rows BEFORE grouping; use HAVING to filter aggregated results AFTER grouping."
        },
        keywords: ["GROUP BY", "HAVING", "SUM"]
    },
    {
        id: 7,
        category: "sql-pwc",
        type: "sql",
        title: "SQL Part 2 (P-Card Audit): Split Purchasing Fraud Detection",
        accountingContext: "Purchasing card policies require authorization for transactions over $2,500. Fraudulent cardholders attempt to bypass approval limits by splitting purchases into smaller amounts on the same day with the same vendor.",
        technicalPrompt: "Write an SQL query to group P-Card transactions by Cardholder_ID, Vendor_Name, and Transaction_Date where COUNT(*) > 1 and SUM(Amount) > 2500.",
        solution: "SELECT Cardholder_ID, Vendor_Name, Transaction_Date, COUNT(*) AS Tx_Count, SUM(Amount) AS Total_Amount\nFROM PCard_Transactions\nGROUP BY Cardholder_ID, Vendor_Name, Transaction_Date\nHAVING COUNT(*) > 1 AND SUM(Amount) > 2500;",
        explanation: {
            accountingRationale: "Detecting split transactions is an essential internal control procedure to catch procurement fraud and spending limit evasions.",
            stepByStep: [
                "1. Multi-Column GROUP BY: Groups transactions by cardholder, vendor, and date simultaneously.",
                "2. COUNT(*) > 1: Identifies multiple purchases made on the same date with the same vendor.",
                "3. SUM(Amount) > 2500: Identifies instances where total daily spend exceeds the authorization threshold."
            ],
            examTip: "Always include all non-aggregated SELECT columns inside the GROUP BY clause."
        },
        keywords: ["GROUP BY", "HAVING", "COUNT"]
    }
];

// State Variables
let userProgress = {};
let userAnswers = {};
let solutionVisibility = {};
let contextModes = {};
let currentMode = "practice";
let examTimerInterval = null;
let timeRemaining = 45 * 60; // 45 minutes in seconds

const container = document.getElementById('question-container');
const categoryFilter = document.getElementById('category-filter');
const modeSelect = document.getElementById('mode-select');
const completedCountEl = document.getElementById('completed-count');
const totalCountEl = document.getElementById('total-count');
const timerBar = document.getElementById('exam-timer-bar');
const timerDisplay = document.getElementById('timer-display');

// Track text input changes in real time
window.updateUserAnswer = function(id, value) {
    userAnswers[id] = value;
};

// Render Questions
function renderQuestions() {
    const selectedCategory = categoryFilter.value;
    container.innerHTML = '';
    
    const filtered = questionBank.filter(q => selectedCategory === 'all' || q.category === selectedCategory);
    totalCountEl.innerText = filtered.length;
    let completed = 0;

    filtered.forEach((q, idx) => {
        if (userProgress[q.id]) completed++;

        const currentAnswer = userAnswers[q.id] || '';
        const isSolutionVisible = solutionVisibility[q.id] ? 'block' : 'none';
        const currentModeState = contextModes[q.id] || 'context';
        const isContext = currentModeState === 'context';
        const promptText = isContext ? q.accountingContext : q.technicalPrompt;
        const toggleBtnLabel = isContext ? "Switch to Technical Prompt" : "Switch to Accounting Context";

        const card = document.createElement('div');
        card.className = 'question-card';
        card.id = `card-${q.id}`;
        
        card.innerHTML = `
            <div class="card-header">
                <span class="tag ${q.type}">${q.type.toUpperCase()}</span>
                <button class="toggle-mode-btn" onclick="toggleContext(${q.id})">${toggleBtnLabel}</button>
            </div>
            <div class="question-title">${idx + 1}. ${q.title}</div>
            <p class="question-text" id="text-${q.id}" data-mode="${currentModeState}" 
               data-context="${q.accountingContext}" 
               data-tech="${q.technicalPrompt}">
               ${promptText}
            </p>
            
            <div class="interactive-area">
                <label style="font-size: 0.85rem; color: #64748b; font-weight:600;">Your Working Formula / Query:</label>
                <textarea id="input-${q.id}" 
                          placeholder="Type your Excel formula or SQL query here..."
                          oninput="updateUserAnswer(${q.id}, this.value)">${currentAnswer}</textarea>
                <div class="action-btns">
                    <button class="btn-primary" onclick="verifyAnswer(${q.id})">Validate Answer</button>
                    <button class="btn-secondary" onclick="toggleExplanation(${q.id})">Toggle Solution & Explanation</button>
                </div>
            </div>

            <div class="explanation-box" id="explanation-${q.id}" style="display: ${isSolutionVisible};">
                <div class="explanation-section">
                    <h4>Official Excel / SQL Solution</h4>
                    <div class="solution-code">${q.solution}</div>
                </div>
                <div class="explanation-section">
                    <h4>Accounting / Audit Rationale</h4>
                    <p style="margin:0; font-size:0.95rem; color:#334155;">${q.explanation.accountingRationale}</p>
                </div>
                <div class="explanation-section">
                    <h4>Step-by-Step Logic Breakdown</h4>
                    <ul style="margin:0; padding-left:1.2rem; font-size:0.93rem; color:#334155;">
                        ${q.explanation.stepByStep.map(s => `<li>${s}</li>`).join('')}
                    </ul>
                </div>
                <div class="exam-tip-box">
                    <strong>💡 Exam Tip:</strong> ${q.explanation.examTip}
                </div>
            </div>
        `;
        container.appendChild(card);
    });

    completedCountEl.innerText = completed;
}

// Toggle between Accounting Context and Technical Prompt
window.toggleContext = function(id) {
    const current = contextModes[id] || 'context';
    contextModes[id] = (current === 'context') ? 'tech' : 'context';
    renderQuestions();
};

// Validate Answer Input
window.verifyAnswer = function(id) {
    const inputEl = document.getElementById(`input-${id}`);
    const inputVal = inputEl ? inputEl.value : (userAnswers[id] || '');
    userAnswers[id] = inputVal;

    if (!inputVal.trim()) {
        alert("Please write your formula or query before validating.");
        return;
    }

    const q = questionBank.find(item => item.id === id);
    const matched = q.keywords.some(kw => inputVal.toUpperCase().includes(kw.toUpperCase()));
    
    // Reveal official solution box upon validation
    solutionVisibility[id] = true;

    if (matched) {
        userProgress[id] = true;
        alert("✅ Answer Validated! Core Excel/SQL syntax recognized.");
    } else {
        alert("⚠️ Check your syntax. Key expected terms (e.g. " + q.keywords.join(", ") + ") were missing.");
    }

    renderQuestions();
};

// Toggle Solution Box Display
window.toggleExplanation = function(id) {
    solutionVisibility[id] = !solutionVisibility[id];
    renderQuestions();
};

// Countdown Timer Logic
function startExamTimer() {
    clearInterval(examTimerInterval);
    timeRemaining = 45 * 60; // 45 minutes

    examTimerInterval = setInterval(() => {
        timeRemaining--;
        
        const minutes = Math.floor(timeRemaining / 60);
        const seconds = timeRemaining % 60;
        
        const formattedMin = minutes < 10 ? `0${minutes}` : minutes;
        const formattedSec = seconds < 10 ? `0${seconds}` : seconds;
        
        if (timerDisplay) {
            timerDisplay.innerText = `${formattedMin}:${formattedSec}`;
        }

        if (timeRemaining <= 0) {
            clearInterval(examTimerInterval);
            alert("⏰ Time is up! Your Mock Exam is being automatically submitted.");
            submitFullExam();
        }
    }, 1000);
}

function stopExamTimer() {
    clearInterval(examTimerInterval);
}

// Submit Full Exam
window.submitFullExam = function() {
    stopExamTimer();
    let solved = Object.keys(userProgress).length;
    let total = questionBank.length;
    let percentage = Math.round((solved / total) * 100);

    alert(`📋 Exam Submitted!\n\nYou solved ${solved} out of ${total} questions (${percentage}%).`);
};

// Mode Selection Handler
modeSelect.addEventListener('change', (e) => {
    currentMode = e.target.value;
    if (currentMode === 'exam') {
        timerBar.style.display = 'flex';
        categoryFilter.value = 'all';
        document.getElementById('module-filter-group').style.display = 'none';
        startExamTimer();
        alert("⏱️ Timed Mock Exam Mode Started! You have 45:00 minutes.");
    } else {
        stopExamTimer();
        timerBar.style.display = 'none';
        document.getElementById('module-filter-group').style.display = 'flex';
    }
    renderQuestions();
});

categoryFilter.addEventListener('change', renderQuestions);

// Initial Page Load
renderQuestions();
