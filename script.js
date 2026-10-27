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
        }
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
        }
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
        }
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
        }
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
        }
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
        }
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
        }
    }
];

// Persistent State Storage
let userProgress = JSON.parse(localStorage.getItem('exam_progress') || '{}');
let savedAnswers = JSON.parse(localStorage.getItem('exam_answers') || '{}');
let currentMode = "practice";
let examTimerInterval = null;
let timeRemaining = 45 * 60;

const container = document.getElementById('question-container');
const categoryFilter = document.getElementById('category-filter');
const modeSelect = document.getElementById('mode-select');
const completedCountEl = document.getElementById('completed-count');
const totalCountEl = document.getElementById('total-count');
const timerBar = document.getElementById('exam-timer-bar');
const timerDisplay = document.getElementById('timer-display');

// Save typed text into localStorage
window.saveAnswer = function(id, text) {
    savedAnswers[id] = text;
    localStorage.setItem('exam_answers', JSON.stringify(savedAnswers));
};

// Render Questions
function renderQuestions() {
    const selectedCategory = categoryFilter.value;
    container.innerHTML = '';
    
    const filtered = questionBank.filter(q => selectedCategory === 'all' || q.category === selectedCategory);
    totalCountEl.innerText = filtered.length;
    updateProgressCounter();

    filtered.forEach((q, idx) => {
        const card = document.createElement('div');
        card.className = 'question-card';
        card.id = `card-${q.id}`;
        
        card.innerHTML = `
            <div class="card-header">
                <span class="tag ${q.type}">${q.type.toUpperCase()}</span>
                <button type="button" class="toggle-mode-btn" id="btn-toggle-ctx-${q.id}" onclick="toggleContext(${q.id})">Switch to Technical Prompt</button>
            </div>
            <div class="question-title">${idx + 1}. ${q.title}</div>
            <p class="question-text" id="text-${q.id}" data-mode="context" 
               data-context="${q.accountingContext}" 
               data-tech="${q.technicalPrompt}">
               ${q.accountingContext}
            </p>
            
            <div class="interactive-area">
                <label style="font-size: 0.85rem; color: #64748b; font-weight:600;">Your Working Formula / Query:</label>
                <textarea id="input-${q.id}" 
                          placeholder="Type your Excel formula or SQL query here..." 
                          oninput="saveAnswer(${q.id}, this.value)"></textarea>
                <div class="action-btns">
                    <button type="button" class="btn-primary" onclick="verifyAnswer(${q.id})">Validate Answer</button>
                    <button type="button" class="btn-secondary" onclick="toggleExplanation(${q.id})">Toggle Solution & Explanation</button>
                </div>
            </div>

            <div class="explanation-box" id="explanation-${q.id}" style="display: none;">
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

    // Restore text into textareas
    filtered.forEach(q => {
        const textarea = document.getElementById(`input-${q.id}`);
        if (textarea && savedAnswers[q.id] !== undefined) {
            textarea.value = savedAnswers[q.id];
        }
    });
}

function updateProgressCounter() {
    let completed = Object.keys(userProgress).length;
    completedCountEl.innerText = completed;
}

// Toggle Context
window.toggleContext = function(id) {
    const textEl = document.getElementById(`text-${id}`);
    const btnEl = document.getElementById(`btn-toggle-ctx-${id}`);
    if (!textEl || !btnEl) return;

    if (textEl.getAttribute('data-mode') === 'context') {
        textEl.innerText = textEl.getAttribute('data-tech');
        textEl.setAttribute('data-mode', 'tech');
        btnEl.innerText = "Switch to Accounting Context";
    } else {
        textEl.innerText = textEl.getAttribute('data-context');
        textEl.setAttribute('data-mode', 'context');
        btnEl.innerText = "Switch to Technical Prompt";
    }
};

// ==========================================
// GRANULAR PEDAGOGICAL EVALUATOR FUNCTIONS
// ==========================================

function evaluateQuestion(id, upperInput, rawVal) {
    let errors = [];

    switch (id) {
        case 1: // Excel Part 1: XLOOKUP / VLOOKUP
            if (!upperInput.startsWith("=")) {
                return { valid: false, msg: "⚠️ Formula Format Error: Excel formulas must start with '=' (e.g. =XLOOKUP(A2, ...))." };
            }
            const isXlookup = upperInput.includes("XLOOKUP");
            const isVlookup = upperInput.includes("VLOOKUP");
            if (!isXlookup && !isVlookup) {
                return { valid: false, msg: "⚠️ Function Error: Missing Excel lookup function. Expected =XLOOKUP(...) or =VLOOKUP(...)." };
            }
            if (isXlookup) {
                if (!upperInput.includes("A2")) errors.push("Missing parameter 1: Search key cell 'A2'");
                if (!upperInput.includes("A:A")) errors.push("Missing parameter 2: Lookup range 'NYC_Film_Permits!A:A'");
                if (!upperInput.includes("D:D")) errors.push("Missing parameter 3: Return range 'NYC_Film_Permits!D:D'");
            } else {
                if (!upperInput.includes("A2")) errors.push("Missing parameter 1: Search key cell 'A2'");
                if (!upperInput.includes("A:D") && !upperInput.includes("A:E")) errors.push("Missing parameter 2: Table array 'NYC_Film_Permits!A:D'");
                if (!upperInput.includes("4")) errors.push("Missing parameter 3: Column index number '4'");
            }
            if (errors.length > 0) return { valid: false, msg: `⚠️ Formula Argument Error:\n• ${errors.join('\n• ')}` };
            return { valid: true, msg: "✅ Answer Validated! Formula search key, lookup range, and return range match the official solution." };

        case 2: // Excel Part 2: TEXTAFTER / RIGHT
            if (!upperInput.startsWith("=")) {
                return { valid: false, msg: "⚠️ Formula Format Error: Excel formulas must start with '=' (e.g. =TEXTAFTER(A2, \"_\"))." };
            }
            const isTextafter = upperInput.includes("TEXTAFTER");
            const isRight = upperInput.includes("RIGHT");
            if (!isTextafter && !isRight) {
                return { valid: false, msg: "⚠️ Function Error: Missing text parsing function. Expected =TEXTAFTER(...) or =RIGHT(...)." };
            }
            if (isTextafter) {
                if (!upperInput.includes("A2")) errors.push("Missing parameter 1: Target text cell 'A2'");
                if (!upperInput.includes("_")) errors.push("Missing parameter 2: Delimiter string \"_\"");
            } else {
                if (!upperInput.includes("A2")) errors.push("Missing parameter 1: Target text cell 'A2'");
                if (!upperInput.includes("LEN")) errors.push("Missing parameter: LEN() function for length calculation");
                if (!upperInput.includes("FIND") && !upperInput.includes("SEARCH")) errors.push("Missing parameter: FIND(\"_\", A2) to locate underscore position");
                if (!upperInput.includes("_")) errors.push("Missing delimiter string \"_\"");
            }
            if (errors.length > 0) return { valid: false, msg: `⚠️ Text Parsing Formula Error:\n• ${errors.join('\n• ')}` };
            return { valid: true, msg: "✅ Answer Validated! Text parsing arguments match the official solution." };

        case 3: // Excel Part 3: SUMIFS
            if (!upperInput.startsWith("=")) {
                return { valid: false, msg: "⚠️ Formula Format Error: Excel formulas must start with '=' (e.g. =SUMIFS(...))." };
            }
            if (!upperInput.includes("SUMIFS")) {
                return { valid: false, msg: "⚠️ Function Error: Expected =SUMIFS(...) function for multi-criteria summation." };
            }
            if (!upperInput.includes("E:E")) errors.push("Missing parameter 1: Sum range 'Payroll2016!E:E' (Gross Pay column)");
            if (!upperInput.includes("C:C") || !upperInput.includes("401")) errors.push("Missing criteria pair 1: JobCode column 'C:C' and value 401");
            if (!upperInput.includes("D:D") || !upperInput.includes("CHICAGO")) errors.push("Missing criteria pair 2: Location column 'D:D' and value \"Chicago\"");

            if (errors.length > 0) return { valid: false, msg: `⚠️ SUMIFS Parameter Error:\n• ${errors.join('\n• ')}` };
            return { valid: true, msg: "✅ Answer Validated! All SUMIFS ranges and criteria match the official audit solution." };

        case 4: // Excel Part 4: Regression Interpretation
            const hasAdv = upperInput.includes("ADVERTISING");
            const hasAdvSig = upperInput.includes("SIGNIFICANT") && !upperInput.includes("NOT SIGNIFICANT");
            if (!hasAdv || !hasAdvSig) {
                errors.push("Missing finding: Advertising expenditure must be identified as statistically significant (P-value 0.002 < 0.05).");
            }
            const hasSqFt = upperInput.includes("SQUARE") || upperInput.includes("FOOTAGE");
            const hasSqFtInsig = upperInput.includes("NOT SIGNIFICANT") || upperInput.includes("INSIGNIFICANT") || upperInput.includes("0.42");
            if (!hasSqFt || !hasSqFtInsig) {
                errors.push("Missing finding: Square Footage must be identified as NOT statistically significant (P-value 0.42 > 0.05).");
            }
            const hasR2 = upperInput.includes("84") || upperInput.includes("0.84");
            if (!hasR2) {
                errors.push("Missing measure: R-squared value of 0.84 (84% revenue variance explained) must be specified.");
            }
            if (errors.length > 0) return { valid: false, msg: `⚠️ Regression Analysis Error:\n• ${errors.join('\n• ')}` };
            return { valid: true, msg: "✅ Answer Validated! All statistical significance findings and R-squared values match." };

        case 5: // SQL Part 1: CS50 Fiftyville Join
            if (!upperInput.includes("SELECT") || (!upperInput.includes("P.NAME") && !upperInput.includes("NAME"))) {
                errors.push("Missing main SELECT clause: Must start with 'SELECT p.name' (or 'SELECT name')");
            }
            if (!upperInput.includes("FROM PEOPLE") && !upperInput.includes("PEOPLE P") && !upperInput.includes("FROM  PEOPLE")) {
                errors.push("Missing main FROM clause: Must query the 'people' table");
            }
            if (!upperInput.includes("JOIN BANK_ACCOUNTS") && !upperInput.includes("BANK_ACCOUNTS")) {
                errors.push("Missing JOIN clause: Must join 'bank_accounts' on person_id");
            }
            if (!upperInput.includes("JOIN ATM_TRANSACTIONS") && !upperInput.includes("ATM_TRANSACTIONS")) {
                errors.push("Missing JOIN clause: Must join 'atm_transactions' on account_number");
            }
            if (!upperInput.includes("LEGGETT STREET")) {
                errors.push("Missing location filter: Must filter atm_location = 'Leggett Street'");
            }
            if (!upperInput.includes("PHONE_CALLS") || !upperInput.includes("60")) {
                errors.push("Missing phone calls subquery: Must filter phone callers with call duration < 60 seconds");
            }
            if (errors.length > 0) return { valid: false, msg: `⚠️ SQL Query Clause Error:\n• ${errors.join('\n• ')}` };
            return { valid: true, msg: "✅ Answer Validated! Relational JOIN chain and subquery filters match the forensic solution." };

        case 6: // SQL Part 2: PwC Chinook Aggregate Customer Revenue
            if (!upperInput.includes("SELECT") || !upperInput.includes("CUSTOMERID") || !upperInput.includes("SUM(")) {
                errors.push("Missing SELECT columns: Must SELECT CustomerId and SUM(Total)");
            }
            if (!upperInput.includes("FROM INVOICE") && !upperInput.includes("FROM  INVOICE")) {
                errors.push("Missing FROM clause: Must specify 'FROM Invoice'");
            }
            if (!upperInput.includes("GROUP BY") || !upperInput.includes("CUSTOMERID")) {
                errors.push("Missing GROUP BY clause: Must include 'GROUP BY CustomerId'");
            }
            if (!upperInput.includes("HAVING") || !upperInput.includes("45")) {
                errors.push("Missing HAVING filter: Must include 'HAVING SUM(Total) > 45'");
            }
            if (!upperInput.includes("ORDER BY")) {
                errors.push("Missing ORDER BY clause: Must include 'ORDER BY TotalSpent DESC' (or ORDER BY SUM(Total) DESC)");
            }
            if (errors.length > 0) return { valid: false, msg: `⚠️ SQL Query Clause Error:\n• ${errors.join('\n• ')}` };
            return { valid: true, msg: "✅ Answer Validated! Grouping, aggregate HAVING filter, and sorting clauses match." };

        case 7: // SQL Part 2: P-Card Split Purchasing Audit
            if (!upperInput.includes("SELECT") || !upperInput.includes("CARDHOLDER_ID") || !upperInput.includes("VENDOR_NAME")) {
                errors.push("Missing SELECT columns: Must SELECT Cardholder_ID, Vendor_Name, Transaction_Date, COUNT(*), SUM(Amount)");
            }
            if (!upperInput.includes("FROM PCARD_TRANSACTIONS") && !upperInput.includes("FROM  PCARD_TRANSACTIONS")) {
                errors.push("Missing FROM clause: Must specify 'FROM PCard_Transactions'");
            }
            if (!upperInput.includes("GROUP BY") || !upperInput.includes("CARDHOLDER_ID") || !upperInput.includes("VENDOR_NAME")) {
                errors.push("Missing or incomplete GROUP BY: Must group by Cardholder_ID, Vendor_Name, and Transaction_Date");
            }
            if (!upperInput.includes("HAVING") || !upperInput.includes("2500")) {
                errors.push("Missing HAVING thresholds: Must filter HAVING COUNT(*) > 1 AND SUM(Amount) > 2500");
            }
            if (errors.length > 0) return { valid: false, msg: `⚠️ P-Card Audit Query Error:\n• ${errors.join('\n• ')}` };
            return { valid: true, msg: "✅ Answer Validated! Split-transaction fraud detection query matches official control solution." };

        default:
            return { valid: false, msg: "⚠️ Unknown question ID." };
    }
}

// Validation Handler
window.verifyAnswer = function(id) {
    const inputEl = document.getElementById(`input-${id}`);
    const expBox = document.getElementById(`explanation-${id}`);
    
    if (!inputEl) return;
    const rawVal = inputEl.value.trim();
    saveAnswer(id, rawVal);

    if (!rawVal) {
        alert("Please write your formula or query before validating.");
        return;
    }

    const upperInput = rawVal.toUpperCase();
    const evalResult = evaluateQuestion(id, upperInput, rawVal);

    if (expBox) expBox.style.display = 'block';

    if (evalResult.valid) {
        userProgress[id] = true;
        localStorage.setItem('exam_progress', JSON.stringify(userProgress));
        updateProgressCounter();
        alert(evalResult.msg);
    } else {
        delete userProgress[id];
        localStorage.setItem('exam_progress', JSON.stringify(userProgress));
        updateProgressCounter();
        alert(evalResult.msg);
    }
};

// Toggle Solution Box Display
window.toggleExplanation = function(id) {
    const expBox = document.getElementById(`explanation-${id}`);
    if (expBox) {
        expBox.style.display = (expBox.style.display === 'block') ? 'none' : 'block';
    }
};

// Countdown Timer Logic
function startExamTimer() {
    clearInterval(examTimerInterval);
    timeRemaining = 45 * 60;

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

// Interactive Exam Submission & Results Banner
window.submitFullExam = function() {
    stopExamTimer();
    
    let solvedCount = Object.keys(userProgress).length;
    let totalCount = questionBank.length;
    let percentage = Math.round((solvedCount / totalCount) * 100);
    let statusText = percentage >= 70 ? "PASS ✅" : "NEEDS PRACTICE ⚠️";
    let statusColor = percentage >= 70 ? "#16a34a" : "#d97706";

    questionBank.forEach(q => {
        const expBox = document.getElementById(`explanation-${q.id}`);
        if (expBox) expBox.style.display = 'block';
    });

    const oldBanner = document.getElementById('exam-results-banner');
    if (oldBanner) oldBanner.remove();

    const resultsBanner = document.createElement('div');
    resultsBanner.id = 'exam-results-banner';
    resultsBanner.style.cssText = `
        background: #ffffff;
        border: 2px solid ${statusColor};
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
    `;

    resultsBanner.innerHTML = `
        <h2 style="margin: 0 0 0.5rem 0; color: #0f172a; font-size: 1.5rem;">📋 Exam Performance Scorecard</h2>
        <div style="font-size: 1.25rem; font-weight: 700; color: ${statusColor}; margin-bottom: 0.5rem;">
            Final Score: ${solvedCount} / ${totalCount} Questions Solved (${percentage}%) — ${statusText}
        </div>
        <p style="margin: 0 0 1.25rem 0; color: #475569; font-size: 0.95rem;">
            All official Excel formulas, SQL queries, and audit explanations have been expanded below for your review.
        </p>
        <button type="button" class="btn-primary" onclick="restartExam()" style="background: #2563eb; color: #fff; padding: 0.6rem 1.2rem; border-radius: 6px; border: none; font-weight: 600; cursor: pointer;">
            🔄 Retake Mock Exam
        </button>
    `;

    container.insertAdjacentElement('beforebegin', resultsBanner);
    window.scrollTo({ top: 0, behavior: 'smooth' });
};

// Restart Exam Handler
window.restartExam = function() {
    userProgress = {};
    savedAnswers = {};
    localStorage.removeItem('exam_progress');
    localStorage.removeItem('exam_answers');
    
    const banner = document.getElementById('exam-results-banner');
    if (banner) banner.remove();

    if (currentMode === 'exam') {
        startExamTimer();
    }
    renderQuestions();
};

// Mode Selection Handler
modeSelect.addEventListener('change', (e) => {
    currentMode = e.target.value;
    const banner = document.getElementById('exam-results-banner');
    if (banner) banner.remove();

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
