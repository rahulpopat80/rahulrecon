// Global error catcher to display any JS errors to the user
window.addEventListener('error', function(event) {
    console.error("Captured global error event:", event);
    
    // Ignore cross-origin "Script error." which lacks details and is usually harmless or caused by blocked scripts/extensions
    if (event.message === "Script error." && (!event.filename || event.filename === "")) {
        console.warn("Ignored generic cross-origin 'Script error.' to prevent blocking UI alerts.");
        return;
    }
    
    let errorMsg = event.message;
    if (event.error && event.error.stack) {
        errorMsg += '\nStack: ' + event.error.stack;
    } else {
        errorMsg += '\nAt: ' + (event.filename || 'unknown') + ':' + (event.lineno || 0);
    }
    alert("સિસ્ટમમાં એરર આવી છે:\n" + errorMsg);
});


// Bank Reconciliation Application Logic

// Application State
const state = {
    currentDate: '2026-07-20',
    similarityThreshold: 0.60,
    dateTolerance: 7,
    files: {
        HDFC: { loaded: false, data: [], balance: 0, rawName: '' },
        345051: { loaded: false, data: [], balance: 0, rawName: '' },
        3493: { loaded: false, data: [], balance: 0, rawName: '' },
        3496: { loaded: false, data: [], balance: 0, rawName: '' },
        history: { loaded: false, data: [], balance: 0, rawName: '' }
    },
    mergedData: [],
    selectedIds: new Set(), // Track selected IDs for bulk actions
    selectedNpciDates: new Set(), // Track selected NPCI dates
    currentTab: 'pending', // 'pending' or 'reconciled'
    searchQuery: '',
    filterType: 'all',
    currentPage: 1,
    pageSize: 999999, // Render all records on one screen
    matchGroupCounter: 0
};

// DOM Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const currentDateInput = document.getElementById('current-date');
const similarityThresholdInput = document.getElementById('similarity-threshold');
const dateToleranceInput = document.getElementById('date-tolerance');
const thresholdVal = document.getElementById('threshold-val');
const toleranceVal = document.getElementById('tolerance-val');

const btnAutoReconcile = document.getElementById('btn-auto-reconcile');
const btnExportPending = document.getElementById('btn-export-pending');
const btnExportReconciled = document.getElementById('btn-export-reconciled');
const btnStartNewDay = document.getElementById('btn-start-new-day');
const btnUndoNewDay = document.getElementById('btn-undo-new-day');
const btnReset = document.getElementById('btn-reset');
const btnBulkAction = document.getElementById('btn-bulk-action');
const bulkActionsContainer = document.getElementById('bulk-actions-container');
const bulkActionText = document.getElementById('bulk-action-text');

const valHdfcBal = document.getElementById('val-hdfc-bal');
const val345051Bal = document.getElementById('val-345051-bal');
const val3493Bal = document.getElementById('val-3493-bal');
const val3496Bal = document.getElementById('val-3496-bal');
const valTotalBal = document.getElementById('val-total-bal');

const subHdfcBal = document.getElementById('sub-hdfc-bal');
const sub345051Bal = document.getElementById('sub-345051-bal');
const sub3493Bal = document.getElementById('sub-3493-bal');
const sub3496Bal = document.getElementById('sub-3496-bal');

const tabPending = document.getElementById('tab-pending');
const tabReconciled = document.getElementById('tab-reconciled');
const tabNPCIPivot = document.getElementById('tab-npci-pivot');
const badgePending = document.getElementById('badge-pending');
const badgeReconciled = document.getElementById('badge-reconciled');

const searchBox = document.getElementById('search-box');
const filterTypeSelect = document.getElementById('filter-type');
const tableBody = document.getElementById('table-body');
const entriesCount = document.getElementById('entries-count');
const paginationControls = document.getElementById('pagination-controls');
const toastContainer = document.getElementById('toast-container');

// RBI Report DOM Elements
const tabRbi = document.getElementById('tab-rbi');
const rbiReportContainer = document.getElementById('rbi-report-container');
const rbiDropZone = document.getElementById('rbi-drop-zone');
const rbiFileInput = document.getElementById('rbi-file-input');
const rbiTableBody = document.getElementById('rbi-table-body');
const rbiResultArea = document.getElementById('rbi-result-area');
const rbiStatementDateLabel = document.getElementById('rbi-statement-date-label');
const btnExportRbi = document.getElementById('btn-export-rbi');

// Event Listeners
window.addEventListener('DOMContentLoaded', () => {
    initApp();
});

function initApp() {
    // Setup file drag and drop
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('active');
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('active');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('active');
        if (e.dataTransfer.files.length > 0) {
            handleFiles(e.dataTransfer.files);
        }
    });
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFiles(e.target.files);
        }
    });

    // Control Inputs
    currentDateInput.addEventListener('change', (e) => {
        state.currentDate = e.target.value;
        showToast('તારીખ અપડેટ કરવામાં આવી: ' + state.currentDate, 'info');
        if (state.mergedData.length > 0) {
            recalculateDays();
            saveStateToLocalStorage();
            renderTable();
        }
    });

    similarityThresholdInput.addEventListener('input', (e) => {
        const val = e.target.value;
        thresholdVal.textContent = val + '%';
        state.similarityThreshold = parseFloat(val) / 100;
    });

    dateToleranceInput.addEventListener('input', (e) => {
        const val = e.target.value;
        toleranceVal.textContent = val + ' દિવસ';
        state.dateTolerance = parseInt(val);
    });

    // Buttons
    btnAutoReconcile.addEventListener('click', runAutoReconciliation);
    btnExportPending.addEventListener('click', () => exportToExcel('pending'));
    btnExportReconciled.addEventListener('click', () => exportToExcel('reconciled'));
    btnStartNewDay.addEventListener('click', startNewDay);
    btnUndoNewDay.addEventListener('click', undoNewDay);
    btnReset.addEventListener('click', resetApp);
    btnBulkAction.addEventListener('click', runBulkAction);
    
    const btnGlobalUndo = document.getElementById('btn-global-undo');
    if (btnGlobalUndo) {
        btnGlobalUndo.addEventListener('click', triggerUndo);
    }
    
    const btnNpciPending = document.getElementById('btn-npci-pending');
    if (btnNpciPending) {
        btnNpciPending.addEventListener('click', runNpciPendingAction);
    }

    // Tab Switching
    tabPending.addEventListener('click', () => switchTab('pending'));
    tabReconciled.addEventListener('click', () => switchTab('reconciled'));
    if (tabNPCIPivot) tabNPCIPivot.addEventListener('click', () => switchTab('npci-pivot'));
    if (tabRbi) tabRbi.addEventListener('click', () => switchTab('rbi'));

    // RBI File Upload Listeners
    if (rbiDropZone && rbiFileInput) {
        rbiDropZone.addEventListener('click', () => rbiFileInput.click());
        rbiDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            rbiDropZone.classList.add('active');
        });
        rbiDropZone.addEventListener('dragleave', () => {
            rbiDropZone.classList.remove('active');
        });
        rbiDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            rbiDropZone.classList.remove('active');
            if (e.dataTransfer.files.length > 0) {
                processRbiExcelFile(e.dataTransfer.files[0]);
            }
        });
        rbiFileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                processRbiExcelFile(e.target.files[0]);
            }
        });
    }
    
    if (btnExportRbi) {
        btnExportRbi.addEventListener('click', exportRbiReportToExcel);
    }

    // Search and Filter
    searchBox.addEventListener('input', (e) => {
        state.searchQuery = e.target.value;
        state.currentPage = 1;
        renderTable();
    });

    filterTypeSelect.addEventListener('change', (e) => {
        state.filterType = e.target.value;
        state.currentPage = 1;
        renderTable();
    });
    
    // Set initial values in inputs
    currentDateInput.value = state.currentDate;
    dateToleranceInput.value = state.dateTolerance;
    toleranceVal.textContent = state.dateTolerance + ' દિવસ';
    similarityThresholdInput.value = Math.round(state.similarityThreshold * 100);
    thresholdVal.textContent = Math.round(state.similarityThreshold * 100) + '%';
    
    // Load persisted state if exists
    loadStateFromLocalStorage();
    
    // Initialize Firebase Cloud Service
    initFirebase();
    
    // Check if new day backup exists to show/hide Undo button
    if (localStorage.getItem('recon_new_day_backup')) {
        btnUndoNewDay.style.display = 'block';
    } else {
        btnUndoNewDay.style.display = 'none';
    }
    
    // Initialize icons
    lucide.createIcons();
}

// Show custom toast notification
function showToast(message, type = 'success') {
    let container = document.getElementById('toast-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        container.className = 'toast-container';
        document.body.appendChild(container);
    }
    
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let iconName = 'check-circle';
    if (type === 'error') iconName = 'alert-triangle';
    if (type === 'info') iconName = 'info';
    
    toast.innerHTML = `
        <i data-lucide="${iconName}"></i>
        <span>${message}</span>
    `;
    container.appendChild(toast);
    if (typeof lucide !== 'undefined' && lucide.createIcons) {
        try { lucide.createIcons({ attrs: { class: 'toast-icon' } }); } catch(e) {}
    }
    
    setTimeout(() => {
        toast.style.animation = 'fadeOut 0.3s forwards';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// Get list of all active reference numbers from HDFC files
function getActiveRefNos() {
    const refs = new Set();
    state.mergedData.forEach(item => {
        if (item.refNo && item.refNo.trim().length >= 4) {
            refs.add(item.refNo.trim());
        }
    });
    return Array.from(refs);
}

// Highlight matching terms, account numbers, IFSC codes, transaction IDs, etc. in Description
function highlightDescription(desc) {
    if (!desc) return '';
    let html = desc;
    
    // First: Match and highlight HDFC Reference Numbers
    const refNos = getActiveRefNos();
    refNos.forEach(ref => {
        const escapedRef = ref.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        const regex = new RegExp('(' + escapedRef + ')', 'gi');
        html = html.replace(regex, (match) => {
            return `<span class="desc-ref-highlight">${match}</span>`;
        });
    });
    
    // 1. Detect IFSC Codes: 4 letters + 0 + 6 alphanumeric
    const ifscRegex = /\b[A-Za-z]{4}0[A-Za-z0-9]{6}\b/g;
    html = html.replace(ifscRegex, (match) => {
        return `<span style="color: #2dd4bf; background-color: rgba(45, 212, 191, 0.12); padding: 1px 4px; border-radius: 4px; font-family: monospace; font-size: 11px; font-weight: 600; border: 1px solid rgba(45, 212, 191, 0.2);">${match}</span>`;
    });
    
    // 2. Detect Account Numbers: e.g. A/C No. XXXXXX1234 or plain numbers of length 9-18
    const accRegex = /(?:A\/C|ACC|Account)(?:\s*(?:No|Number))?\.?\s*[A-Za-z0-9*]+/gi;
    const plainAccRegex = /\b\d{9,18}\b/g;
    const maskedAccRegex = /\b\d{2,6}[*xX-]+\d{2,6}\b/g;
    
    html = html.replace(accRegex, (match) => {
        return `<span style="color: #fb923c; background-color: rgba(251, 146, 60, 0.12); padding: 1px 4px; border-radius: 4px; font-family: monospace; font-size: 11px; font-weight: 600; border: 1px solid rgba(251, 146, 60, 0.2);">${match}</span>`;
    });
    
    html = html.replace(plainAccRegex, (match) => {
        // Skip year-like or date-like numbers
        if (match.length >= 9) {
            return `<span style="color: #fb923c; background-color: rgba(251, 146, 60, 0.12); padding: 1px 4px; border-radius: 4px; font-family: monospace; font-size: 11px; font-weight: 600; border: 1px solid rgba(251, 146, 60, 0.2);">${match}</span>`;
        }
        return match;
    });
    
    html = html.replace(maskedAccRegex, (match) => {
        return `<span style="color: #fb923c; background-color: rgba(251, 146, 60, 0.12); padding: 1px 4px; border-radius: 4px; font-family: monospace; font-size: 11px; font-weight: 600; border: 1px solid rgba(251, 146, 60, 0.2);">${match}</span>`;
    });

    // 3. Detect Reference / Transaction Numbers: Ref No, Txn ID, Transfer reference numbers
    const refRegex = /\b(?:Ref|Txn|Transfer|UPI|NEFT|RTGS|IMPS)(?:\s*(?:No|ID|Ref))?\.?\s*[A-Za-z0-9-]{8,22}\b/gi;
    html = html.replace(refRegex, (match) => {
        return `<span style="color: #60a5fa; background-color: rgba(96, 165, 250, 0.12); padding: 1px 4px; border-radius: 4px; font-family: monospace; font-size: 11px; border: 1px solid rgba(96, 165, 250, 0.2);">${match}</span>`;
    });

    // 4. Common keywords: SALARY, RENT, INTEREST, CHARGES, CASH, CHEQUE, CHQ
    const keywordsRegex = /\b(salary|rent|interest|charges|cash|cheque|chq|refund|commission|loan|emi|fee|bonus|gst|tax|payment|bill)\b/gi;
    html = html.replace(keywordsRegex, (match) => {
        return `<span style="color: #c084fc; background-color: rgba(192, 132, 252, 0.12); padding: 1px 5px; border-radius: 4px; font-weight: 600; font-size: 11px; text-transform: uppercase; border: 1px solid rgba(192, 132, 252, 0.2);">${match}</span>`;
    });
    
    return html;
}

// Automatically determine category flag based on transaction description
function getAutoFlag(desc) {
    if (!desc) return 'DAILY';
    
    // NPCI matches (using letter-boundary check to avoid false positives like KACHOT, RANACHHO)
    const npciRegex = /(?:^|[^A-Z])(NPCI|ACH|APBS)(?:$|[^A-Z])/i;
    if (npciRegex.test(desc)) {
        return 'NPCI';
    }
    
    const upperDesc = desc.toUpperCase();
    
    // A matches
    if (upperDesc.includes('CTS CHEQUE') || upperDesc.includes('CTS')) {
        return 'A';
    }
    
    // AA matches
    if (upperDesc.includes('POS SETT') || upperDesc.includes('ATM SETT')) {
        return 'AA';
    }
    
    // IMPS 1C to IMPS 15C
    const impsMatch = upperDesc.match(/IMPS\s*(\d+)C/i);
    if (impsMatch) {
        const num = parseInt(impsMatch[1]);
        if (num >= 1 && num <= 15) return 'AA';
    }
    
    // UPI 1C to UPI 15C
    const upiMatch = upperDesc.match(/UPI\s*(\d+)C/i);
    if (upiMatch) {
        const num = parseInt(upiMatch[1]);
        if (num >= 1 && num <= 15) return 'AA';
    }
    
    // AB matches
    if (upperDesc.includes('LPG BILLING') || 
        upperDesc.includes('TDS ON NACH') || 
        upperDesc.includes('JCOM OW IW CLG CHARGES') || 
        upperDesc.includes('NPCI ALL PRODUCT CHARGES') || 
        upperDesc.includes('NPCI AVCBDT')) {
        return 'AB';
    }
    
    // ECMS matches
    if (upperDesc.includes('FT CR ECMS')) {
        return 'ECMS';
    }
    
    return 'DAILY';
}

// Extract date pattern from transaction description (e.g. DD/MM/YYYY, DD-MM-YYYY, DD-MMM-YYYY, DD/MM/YY, etc.)
function extractDateFromDescription(description, fallbackDate) {
    if (!description) return fallbackDate || '';
    const str = String(description);
    
    // Pattern 1: DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY (e.g. 05/07/2026, 05-07-2026, 05.07.2026)
    const matchFullDate = str.match(/\b([0-3]?\d)[\/\.\-]([0-1]?\d)[\/\.\-](20\d{2})\b/);
    if (matchFullDate) {
        const day = matchFullDate[1].padStart(2, '0');
        const month = matchFullDate[2].padStart(2, '0');
        const year = matchFullDate[3];
        return `${day}/${month}/${year}`;
    }
    
    // Pattern 2: DD/MM/YY or DD-MM-YY or DD.MM.YY (e.g. 05/07/26, 05-07-26)
    const matchShortYear = str.match(/\b([0-3]?\d)[\/\.\-]([0-1]?\d)[\/\.\-](\d{2})\b/);
    if (matchShortYear) {
        const day = matchShortYear[1].padStart(2, '0');
        const month = matchShortYear[2].padStart(2, '0');
        const year = '20' + matchShortYear[3];
        return `${day}/${month}/${year}`;
    }
    
    // Pattern 3: DD-MMM-YYYY or DD/MMM/YYYY or DD-MMM-YY (e.g. 05-JUL-2026, 05-JUL-26)
    const monthsMap = {
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
        'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'
    };
    const matchTextMonth = str.match(/\b([0-3]?\d)[\/\.\-]([A-Za-z]{3})[\/\.\-](20\d{2}|\d{2})\b/i);
    if (matchTextMonth) {
        const day = matchTextMonth[1].padStart(2, '0');
        const monthStr = matchTextMonth[2].toUpperCase();
        let year = matchTextMonth[3];
        if (year.length === 2) year = '20' + year;
        if (monthsMap[monthStr]) {
            return `${day}/${monthsMap[monthStr]}/${year}`;
        }
    }
    
    // Pattern 4: 8-digit date string DDMMYYYY (e.g. 05072026)
    const matchDigits = str.match(/\b([0-3]\d)(0[1-9]|1[0-2])(20\d{2})\b/);
    if (matchDigits) {
        return `${matchDigits[1]}/${matchDigits[2]}/${matchDigits[3]}`;
    }

    return fallbackDate || '';
}

// Update FLAG value in state on inline edit
function updateFlagValue(itemId, newValue) {
    const item = state.mergedData.find(t => t.id === itemId);
    if (item) {
        const cleanedValue = String(newValue || '').trim();
        item.flag = cleanedValue;
        console.log(`Updated flag for ${itemId} to: ${cleanedValue}`);
        saveStateToLocalStorage();
        refreshLedgerCounts();
        updateBRSLiveWidget();
        renderTable();
    }
}

// Update TYPE value in state on inline edit
function updateTransactionType(itemId, newType, element) {
    const item = state.mergedData.find(t => t.id === itemId);
    if (!item) return;
    
    const cleanType = String(newType || '').trim().toUpperCase();
    if (!cleanType) {
        element.textContent = item.type;
        return;
    }
    
    if (item.type === cleanType) return;
    
    item.type = cleanType;
    
    // Update styling class of parent td element to change type badge color
    const td = element.closest('td');
    if (td) {
        td.className = 'type-cell ' + (cleanType === 'HDFC' ? 'HDFC' : 'x' + cleanType);
    }
    
    saveStateToLocalStorage();
    refreshLedgerCounts();
    updateBRSLiveWidget();
}

// Import new transactions for a file type while preserving existing ledger entries
function importNewTransactions(newRows, type) {
    // Clear existing transactions of this type first to prevent duplication
    if (type === 'history') {
        state.mergedData = state.mergedData.filter(item => 
            item.id && !item.id.startsWith('history_') && item.type !== 'history' && item.parentType !== 'history' && !item.isCarriedOver
        );
    } else {
        state.mergedData = state.mergedData.filter(item => item.type !== type || item.isCarriedOver || (item.id && item.id.startsWith('history_')));
    }
    
    // Clear selection for deleted IDs to be clean
    state.selectedIds.clear();
    updateBulkActionButtons();
    
    const existingPool = {};
    
    let newImportedCount = 0;
    
    newRows.forEach((row, index) => {
        let dateStr = '';
        let desc = '';
        let credit = 0;
        let debit = 0;
        let refNo = '';
        let actualDate = '';
        let cdFlag = '';
        let hasCDFlagColumn = false;
        
        if (type === 'HDFC') {
            let rawDate = String(getRowValue(row, ['TransactionDate', 'Date', 'TxnDate', 'ValueDate']) || '').trim();
            if (rawDate.includes(' ')) {
                rawDate = rawDate.split(' ')[0];
            }
            dateStr = rawDate;
            desc = String(getRowValue(row, ['Description', 'Narration', 'Particulars', 'Particular', 'Details']) || '').trim();
            cdFlag = String(getRowValue(row, ['CDFalg', 'CDFlag', 'DrCr', 'C.D.Falg', 'C.D.Flag']) || '').trim().toUpperCase();
            hasCDFlagColumn = getRowValue(row, ['CDFalg', 'CDFlag', 'DrCr', 'C.D.Falg', 'C.D.Flag']) !== '';
            refNo = extractRefNoFromRow(row);
            
            if (!dateStr || !desc) return;
            
            if (hasCDFlagColumn) {
                const amtStr = String(getRowValue(row, ['Amount', 'TransactionAmount', 'Value', 'રકમ']) || '').replace(/,/g, '').trim();
                const amount = parseFloat(amtStr) || 0;
                if (cdFlag === 'CR' || cdFlag === 'CREDIT' || cdFlag === 'C') {
                    credit = amount;
                } else {
                    debit = amount;
                }
            } else {
                const creditStr = String(getRowValue(row, ['CreditAmount', 'Credit', 'Deposit', 'CreditAmt']) || '').replace(/,/g, '').trim();
                const debitStr = String(getRowValue(row, ['Debitamount', 'Debit', 'Withdrawal', 'DebitAmt']) || '').replace(/,/g, '').trim();
                credit = parseFloat(creditStr) || 0;
                debit = parseFloat(debitStr) || 0;
            }
            if (credit === 0 && debit === 0 && desc && desc.toLowerCase().includes('balance')) return;
        } else if (type === 'history') {
            dateStr = String(getRowValue(row, ['DATE', 'Date']) || '').trim();
            desc = String(getRowValue(row, ['DISCRIPTION', 'Description', 'Narration']) || '').trim();
            const typeVal = String(getRowValue(row, ['TYPE', 'Type']) || '').trim().toUpperCase();
            const actualDateVal = String(getRowValue(row, ['ACTUAL DATE', 'ActualDate']) || '').trim();
            const creditStr = String(getRowValue(row, ['CREDIT TRN', 'Credit']) || '').replace(/,/g, '').trim();
            const debitStr = String(getRowValue(row, ['DEBIT TRN', 'Debit']) || '').replace(/,/g, '').trim();
            
            if (!dateStr || !desc) return;
            credit = parseFloat(creditStr) || 0;
            debit = parseFloat(debitStr) || 0;
            
            const flagStr = String(getRowValue(row, ['FLAG', 'Flag']) || '').trim();
            
            const key = [
                dateStr,
                desc.trim(),
                credit.toFixed(2),
                debit.toFixed(2),
                ''
            ].join('|');
            
            if (existingPool[key] > 0) {
                existingPool[key]--;
                return;
            }
            
            state.mergedData.push({
                id: 'history_' + index,
                date: formatDateToSlash(dateStr),
                description: desc,
                type: typeVal || 'PEND',
                parentType: typeVal || 'HDFC',
                actualDate: formatDateToSlash(actualDateVal || dateStr),
                creditTrn: credit,
                debitTrn: debit,
                refNo: '',
                flag: flagStr || getAutoFlag(desc),
                day: '',
                count: 0,
                reconciled: false,
                matchGroupId: null
            });
            newImportedCount++;
            return;
        } else {
            dateStr = String(getRowValue(row, ['POSTDATE', 'Date']) || '').trim();
            desc = String(getRowValue(row, ['USERNARRATION', 'Description']) || '').trim();
            const creditStr = String(getRowValue(row, ['CreditAmount', 'Credit']) || '').replace(/,/g, '').trim();
            const debitStr = String(getRowValue(row, ['Debitamount', 'Debit']) || '').replace(/,/g, '').trim();
            
            if (!dateStr && !desc) return;
            credit = parseFloat(creditStr) || 0;
            debit = parseFloat(debitStr) || 0;
            if (credit === 0 && debit === 0 && desc && desc.toLowerCase().includes('balance')) return;
        }
        
        const dateObj = parseDateString(dateStr);
        actualDate = formatDateOnly(dateObj);
        
        const candidateKey = [
            dateStr,
            desc.trim(),
            credit.toFixed(2),
            debit.toFixed(2),
            refNo.trim()
        ].join('|');
        
        if (existingPool[candidateKey] > 0) {
            existingPool[candidateKey]--;
        } else {
            state.mergedData.push({
                id: type.toLowerCase() + '_' + index,
                date: formatDateToSlash(dateStr),
                description: desc,
                type: type,
                actualDate: formatDateToSlash(actualDate),
                creditTrn: credit,
                debitTrn: debit,
                refNo: refNo,
                flag: getAutoFlag(desc),
                day: 0,
                count: 0,
                reconciled: false,
                matchGroupId: null
            });
            newImportedCount++;
        }
    });
    
    return newImportedCount;
}

// Calculate amount frequencies for COUNT column
function recalculateAmountFrequencies() {
    const creditFreq = {};
    const debitFreq = {};
    
    state.mergedData.forEach(item => {
        if (item.creditTrn > 0) {
            const amt = item.creditTrn.toFixed(2);
            creditFreq[amt] = (creditFreq[amt] || 0) + 1;
        }
        if (item.debitTrn > 0) {
            const amt = item.debitTrn.toFixed(2);
            debitFreq[amt] = (debitFreq[amt] || 0) + 1;
        }
    });
    
    state.mergedData.forEach(item => {
        if (item.creditTrn > 0) {
            const amt = item.creditTrn.toFixed(2);
            item.count = creditFreq[amt] || 1;
        } else if (item.debitTrn > 0) {
            const amt = item.debitTrn.toFixed(2);
            item.count = debitFreq[amt] || 1;
        } else {
            item.count = 0;
        }
    });
}

// Refresh Serial numbers and sort ledger
function refreshLedgerCounts() {
    // 1. Recalculate days difference first so that they are always fresh
    recalculateDays();
    
    // 2. Recalculate frequencies first so that the sort comparator has access to fresh counts
    recalculateAmountFrequencies();
    
    // 2.5 Auto-recover group IDs for reconciled items of the same amount that don't have a group ID
    const reconciledItems = state.mergedData.filter(item => item.reconciled);
    const unmatchedReconciled = reconciledItems.filter(item => !item.matchGroupId);
    if (unmatchedReconciled.length > 0) {
        const byAmt = {};
        unmatchedReconciled.forEach(item => {
            const amt = Math.max(item.creditTrn, item.debitTrn).toFixed(2);
            if (!byAmt[amt]) byAmt[amt] = [];
            byAmt[amt].push(item);
        });
        
        Object.keys(byAmt).forEach(amtStr => {
            const group = byAmt[amtStr];
            const hdfcItems = group.filter(item => getFileCategory(item) === 'HDFC');
            const glItems = group.filter(item => getFileCategory(item) !== 'HDFC');
            
            const minLen = Math.min(hdfcItems.length, glItems.length);
            for (let i = 0; i < minLen; i++) {
                const h = hdfcItems[i];
                const g = glItems[i];
                
                const newGroupId = 'recovered_group_' + state.matchGroupCounter;
                state.matchGroupCounter++;
                
                h.matchGroupId = newGroupId;
                g.matchGroupId = newGroupId;
            }
        });
    }
    
    // 3. Pre-compute group maximum amounts for sorting matched pairs/groups
    const groupMaxAmt = {};
    state.mergedData.forEach(item => {
        const amt = Math.max(item.creditTrn, item.debitTrn);
        if (item.matchGroupId) {
            groupMaxAmt[item.matchGroupId] = Math.max(groupMaxAmt[item.matchGroupId] || 0, amt);
        }
    });
    
    // 4. Sort the ledger data
    state.mergedData.sort((a, b) => {
        const getGroupId = (item) => {
            if (item.reconciled && item.matchGroupId && item.matchGroupId !== 'null' && item.matchGroupId !== 'undefined') {
                return item.matchGroupId;
            }
            return null;
        };
        
        const gA = getGroupId(a);
        const gB = getGroupId(b);
        
        const getSortAmount = (item, gId) => {
            if (gId) {
                return groupMaxAmt[gId] || Math.max(item.creditTrn, item.debitTrn);
            }
            return Math.max(item.creditTrn, item.debitTrn);
        };
        
        const amtA = getSortAmount(a, gA);
        const amtB = getSortAmount(b, gB);
        
        // 1. Sort by amount descending
        if (Math.abs(amtA - amtB) > 0.001) {
            return amtB - amtA;
        }
        
        // 2. If same amount
        if (gA && gB) {
            if (gA === gB) {
                // Same group: HDFC first, then GL
                const isHdfcA = getFileCategory(a) === 'HDFC';
                const isHdfcB = getFileCategory(b) === 'HDFC';
                if (isHdfcA !== isHdfcB) {
                    return isHdfcA ? -1 : 1;
                }
                return b.creditTrn - a.creditTrn;
            } else {
                // Different groups: sort by group ID to keep them together
                return String(gA).localeCompare(String(gB));
            }
        }
        
        if (gA) return -1;
        if (gB) return 1;
        
        const idA = a.id || '';
        const idB = b.id || '';
        return String(idA).localeCompare(String(idB));
    });
}

// Compress state.mergedData to save space
function compressMergedData(data) {
    if (!data || !Array.isArray(data)) return [];
    return data.map(item => ({
        id: item.id || '',
        d: item.date || '',
        ds: item.description || '',
        t: item.type || '',
        c: item.creditTrn || 0,
        db: item.debitTrn || 0,
        f: item.flag || 'DAILY',
        ad: item.actualDate || null,
        r: item.reconciled ? 1 : 0,
        mg: item.matchGroupId || null,
        p: item.parentType || null,
        co: item.isCarriedOver ? 1 : 0
    }));
}

// Decompress state.mergedData
function decompressMergedData(compressed) {
    if (!compressed || !Array.isArray(compressed)) return [];
    return compressed.map(c => ({
        id: c.id,
        date: c.d,
        description: c.ds,
        type: c.t,
        creditTrn: c.c || 0,
        debitTrn: c.db || 0,
        flag: c.f || 'DAILY',
        actualDate: c.ad,
        reconciled: c.r === 1,
        matchGroupId: c.mg || null,
        parentType: c.p || undefined,
        isCarriedOver: c.co === 1
    }));
}

// Push current state to Undo stack before performing changes
function pushToUndoStack() {
    if (!state.undoStack) state.undoStack = [];
    if (state.undoStack.length >= 20) {
        state.undoStack.shift();
    }
    state.undoStack.push(JSON.stringify({
        mergedData: compressMergedData(state.mergedData),
        matchGroupCounter: state.matchGroupCounter,
        selectedIds: Array.from(state.selectedIds)
    }));
    updateUndoButtonVisibility();
}

// Trigger Undo to restore previous state
function triggerUndo() {
    if (!state.undoStack || state.undoStack.length === 0) {
        showToast("પાછી મેળવવા માટે કોઈ એક્શન નથી (Undo stack empty).", "warning");
        return;
    }
    try {
        const prev = JSON.parse(state.undoStack.pop());
        state.mergedData = decompressMergedData(prev.mergedData);
        state.matchGroupCounter = prev.matchGroupCounter || 0;
        state.selectedIds = new Set(prev.selectedIds || []);
        
        saveStateToLocalStorage();
        refreshLedgerCounts();
        updateBRSLiveWidget();
        renderTable();
        updateUndoButtonVisibility();
        showToast("છેલ્લી ક્રિયા પાછી મેળવી લેવામાં આવી છે (Undo success)!", "success");
    } catch (e) {
        console.error("[Undo] Restoring state failed: ", e);
        showToast("Undo કરવામાં કોઈ ભૂલ આવી.", "error");
    }
}

// Update the global Undo button display state
function updateUndoButtonVisibility() {
    const btn = document.getElementById('btn-global-undo');
    if (btn) {
        btn.style.display = (state.undoStack && state.undoStack.length > 0) ? 'inline-flex' : 'none';
    }
}

// Save mergedData and metadata to localStorage (optimized to prevent QuotaExceededError)
function saveStateToLocalStorage() {
    try {
        localStorage.setItem('recon_merged_data', JSON.stringify(compressMergedData(state.mergedData)));
        localStorage.setItem('recon_match_counter', state.matchGroupCounter.toString());
        if (state.currentDate) {
            localStorage.setItem('recon_current_date', state.currentDate);
        }
        
        // Strip large raw data arrays to keep localStorage footprint under 200KB
        const filesMetaToSave = {};
        for (let key in state.files) {
            filesMetaToSave[key] = {
                loaded: state.files[key].loaded,
                balance: state.files[key].balance,
                previousBalance: state.files[key].previousBalance || 0,
                rawName: state.files[key].rawName,
                rowCount: state.files[key].data ? state.files[key].data.length : 0
            };
        }
        localStorage.setItem('recon_files_meta', JSON.stringify(filesMetaToSave));
        
        // Auto sync state to Firebase Cloud Storage
        if (typeof saveStateToFirebase === 'function') {
            saveStateToFirebase(false);
        }
    } catch (e) {
        console.error("[Recon] LocalStorage save failed: ", e);
        showToast("લોકલ સ્ટોરેજ સેવ કરવામાં ભૂલ આવી (શક્ય છે કે સ્ટોરેજ ફૂલ છે).", 'warning');
    }
}

// ------------------------------------------------------------------
// FIREBASE CLOUD BACKEND SERVICE INTEGRATION
// ------------------------------------------------------------------
const firebaseConfig = {
  apiKey: "AIzaSyBEqZS_NZERxw2JIEcpLymo734vU5SPP9A",
  authDomain: "rahulrecon-a5083.firebaseapp.com",
  projectId: "rahulrecon-a5083",
  storageBucket: "rahulrecon-a5083.firebasestorage.app",
  messagingSenderId: "432963671020",
  appId: "1:432963671020:web:bf3e76aaf692a80c2226c9",
  measurementId: "G-GTTSS1CM96"
};

let db = null;
let analytics = null;
let firebaseInitialized = false;
let isSelfSaving = false;

function initFirebase() {
    try {
        if (typeof firebase !== 'undefined') {
            firebase.initializeApp(firebaseConfig);
            db = firebase.firestore();
            if (firebase.analytics) {
                try { analytics = firebase.analytics(); } catch(e) {}
            }
            firebaseInitialized = true;
            console.log("[Firebase] Successfully initialized Firebase Cloud Service.");
            updateFirebaseUIStatus(true, '☁️ Cloud Synced');
            
            // ALWAYS load latest state from Cloud on startup for multi-PC sync (Home PC <-> Bank PC)
            loadStateFromFirebase(true);
            
            // Listen for live Cloud updates across multiple PCs
            listenToCloudChanges();
        } else {
            console.warn("[Firebase] Firebase SDK scripts not loaded.");
            updateFirebaseUIStatus(false, 'Offline Mode');
        }
    } catch (err) {
        console.error("[Firebase] Initialization error:", err);
        updateFirebaseUIStatus(false, 'Firebase Offline');
    }
}

function listenToCloudChanges() {
    if (!firebaseInitialized || !db) return;
    
    db.collection("recon_snapshots").doc("current_state")
        .onSnapshot((doc) => {
            if (doc.exists && !isSelfSaving) {
                const data = doc.data();
                if (data.mergedData && Array.isArray(data.mergedData)) {
                    const firstItem = data.mergedData[0];
                    state.mergedData = (firstItem && 'd' in firstItem) ? decompressMergedData(data.mergedData) : data.mergedData;
                    if (data.currentDate) {
                        state.currentDate = data.currentDate;
                        if (currentDateInput) currentDateInput.value = data.currentDate;
                    }
                    state.matchGroupCounter = data.matchGroupCounter || 0;
                    if (data.files) state.files = data.files;
                    
                    saveStateToLocalStorage();
                    refreshLedgerCounts();
                    updateBRSLiveWidget();
                    renderTable();
                    
                    updateFirebaseUIStatus(true, '☁️ Cloud Synced');
                }
            }
        }, (error) => {
            console.warn("[Firebase Realtime Sync] Listener note:", error);
        });
}

let firebaseSaveTimeout = null;

function saveStateToFirebase(showNotification = false) {
    if (!firebaseInitialized || !db) {
        if (showNotification) showToast("ફાયરબેઝ કનેક્ટેડ નથી (ઓફલાઈન મોડ).", "warning");
        return;
    }
    
    // Mark as self-saving immediately to prevent incoming onSnapshot events
    // from overwriting the state while we wait for the debounce timeout.
    isSelfSaving = true;
    
    if (firebaseSaveTimeout) {
        clearTimeout(firebaseSaveTimeout);
    }
    
    firebaseSaveTimeout = setTimeout(() => {
        
        const filesMetaToSave = {};
        for (let key in state.files) {
            filesMetaToSave[key] = {
                loaded: state.files[key].loaded || false,
                balance: state.files[key].balance || 0,
                previousBalance: state.files[key].previousBalance || 0,
                rawName: state.files[key].rawName || '',
                rowCount: state.files[key].rowCount || (state.files[key].data ? state.files[key].data.length : 0)
            };
        }
        
        const payload = {
            currentDate: state.currentDate || '',
            similarityThreshold: state.similarityThreshold || 0.60,
            dateTolerance: state.dateTolerance || 7,
            matchGroupCounter: state.matchGroupCounter || 0,
            files: filesMetaToSave,
            mergedData: compressMergedData(state.mergedData),
            lastUpdatedFormatted: new Date().toLocaleString('en-IN')
        };
        
        updateFirebaseUIStatus(true, 'Saving to Cloud...');
        
        db.collection("recon_snapshots").doc("current_state").set(payload)
            .then(() => {
                console.log("[Firebase] State saved successfully to Firestore.");
                updateFirebaseUIStatus(true, '☁️ Cloud Synced');
                setTimeout(() => { isSelfSaving = false; }, 1500);
                if (showNotification) {
                    showToast("તમામ ડેટા ક્લાઉડ સર્વર પર સેવ થઈ ગયો છે.", "success");
                }
                
                // Also save daily history snapshot
                if (state.currentDate) {
                    db.collection("recon_daily_history").doc(state.currentDate).set(payload).catch(e => console.warn(e));
                }
            })
            .catch((error) => {
                isSelfSaving = false;
                console.warn("[Firebase] Error saving state:", error);
                updateFirebaseUIStatus(false, 'Cloud Perm Error');
                if (showNotification) {
                    showToast("ફાયરબેઝ પરમિશન એરર: Firebase Console માં Firestore Database 'Rules' ચાલુ કરો.", "error");
                }
            });
    }, 1000);
}

function loadStateFromFirebase(silent = false) {
    if (!firebaseInitialized || !db) {
        if (!silent) showToast("ફાયરબેઝ કનેક્ટેડ નથી!", "error");
        updateFirebaseUIStatus(true, 'System Ready (Local)');
        return;
    }
    
    if (!silent) updateFirebaseUIStatus(true, 'Fetching Cloud Data...');
    
    db.collection("recon_snapshots").doc("current_state").get()
        .then((doc) => {
            if (doc.exists) {
                const data = doc.data();
                
                // Prevent Firebase from downgrading newer local data (fixes refresh reverting issue)
                if (data.matchGroupCounter !== undefined && data.matchGroupCounter < state.matchGroupCounter) {
                    console.log("[Firebase] Local state is newer than Cloud state. Skipping overwrite.");
                    return;
                }
                
                if (data.mergedData && Array.isArray(data.mergedData)) {
                    const firstItem = data.mergedData[0];
                    state.mergedData = (firstItem && 'd' in firstItem) ? decompressMergedData(data.mergedData) : data.mergedData;
                    if (data.currentDate) {
                        state.currentDate = data.currentDate;
                        if (currentDateInput) currentDateInput.value = data.currentDate;
                    }
                    state.matchGroupCounter = data.matchGroupCounter || 0;
                    if (data.files) state.files = data.files;
                    
                    saveStateToLocalStorage();
                    refreshLedgerCounts();
                    updateBRSLiveWidget();
                    renderTable();
                    
                    updateFirebaseUIStatus(true, '☁️ Cloud Synced');
                    if (!silent) {
                        showToast(`ક્લાઉડમાંથી ${state.mergedData.length} વ્યવહારો અને સ્ટેટસ સફળતાપૂર્વક લોડ થયા.`, "success");
                    }
                } else {
                    updateFirebaseUIStatus(true, '☁️ Cloud Ready');
                    if (!silent) showToast("ક્લાઉડમાં કોઈ જૂનો ડેટા મળ્યો નથી.", "info");
                }
            } else {
                updateFirebaseUIStatus(true, '☁️ Cloud Ready');
                if (!silent) showToast("ક્લાઉડમાં કોઈ ડેટા સેવ કરેલ નથી.", "info");
            }
        })
        .catch((error) => {
            console.warn("[Firebase] Error fetching state:", error);
            if (silent) {
                // Silent startup fallback to local mode
                updateFirebaseUIStatus(true, 'System Ready');
            } else {
                updateFirebaseUIStatus(false, 'Cloud Perm Error');
                showToast("ક્લાઉડ પરમિશન જરૂરી છે: Firebase Console માં Firestore Rules ચાલુ કરો.", "error");
            }
        });
}

function downloadFirebaseBackup() {
    if (!firebaseInitialized || !db) {
        showToast("ફાયરબેઝ કનેક્ટેડ નથી. લોકલ ડેટા ડાઉનલોડ થાય છે...", "info");
        const jsonStr = JSON.stringify(state, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Local_Recon_Backup_${state.currentDate}.json`;
        a.click();
        return;
    }
    
    showToast("ફાયરબેઝમાંથી બેકઅપ ફાઈલ ડાઉનલોડ થઈ રહી છે...", "info");
    
    db.collection("recon_snapshots").doc("current_state").get()
        .then((doc) => {
            let backupObj = {};
            if (doc.exists) {
                backupObj = doc.data();
            } else {
                backupObj = {
                    currentDate: state.currentDate,
                    files: state.files,
                    mergedData: state.mergedData,
                    matchGroupCounter: state.matchGroupCounter
                };
            }
            
            const jsonStr = JSON.stringify(backupObj, null, 2);
            const blob = new Blob([jsonStr], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            
            const dateStr = state.currentDate || new Date().toISOString().slice(0, 10);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Firebase_Recon_Backup_${dateStr}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
            showToast("ફાયરબેઝ ડેટા બેકઅપ JSON ફાઈલ સફળતાપૂર્વક ડાઉનલોડ થઈ ગઈ.", "success");
        })
        .catch((err) => {
            console.error("[Firebase] Download error:", err);
            showToast("ફાયરબેઝ ડાઉનલોડમાં એરર આવી: " + err.message, "error");
        });
}

function updateFirebaseUIStatus(online, text) {
    const badge = document.getElementById('firebase-status-badge');
    if (!badge) return;
    
    if (online) {
        badge.style.background = 'rgba(16, 185, 129, 0.15)';
        badge.style.color = '#10b981';
        badge.style.border = '1px solid rgba(16, 185, 129, 0.3)';
        badge.innerHTML = `<span class="dot"></span> 🔥 ${text}`;
    } else {
        badge.style.background = 'rgba(239, 68, 68, 0.15)';
        badge.style.color = '#ef4444';
        badge.style.border = '1px solid rgba(239, 68, 68, 0.3)';
        badge.innerHTML = `<span class="dot" style="background:#ef4444;"></span> ⚠️ ${text}`;
    }
}

// Update live BRS widget details in sidebar
// Update live BRS widget details in sidebar
function updateBRSLiveWidget() {
    const widget = document.getElementById('brs-live-widget');
    const container = document.getElementById('brs-live-table-container');
    if (!widget || !container) return;
    
    if (state.mergedData.length === 0) {
        widget.style.display = 'none';
        return;
    }
    
    widget.style.display = 'block';
    
    console.log("[Recon Debug] updateBRSLiveWidget executing. state.selectedIds:", Array.from(state.selectedIds));
    
    const hdfc_bal = (state.files && state.files.HDFC && state.files.HDFC.loaded) ? state.files.HDFC.balance : 0;
    const gl_345051_bal = (state.files && state.files['345051'] && state.files['345051'].loaded) ? state.files['345051'].balance : 0;
    const gl_3493_bal = (state.files && state.files['3493'] && state.files['3493'].loaded) ? Math.abs(state.files['3493'].balance) : 0;
    const gl_3496_bal = (state.files && state.files['3496'] && state.files['3496'].loaded) ? Math.abs(state.files['3496'].balance) : 0;
    
    let gl_pending_credits = 0;
    let gl_pending_debits = 0;
    let hdfc_pending_credits = 0;
    let hdfc_pending_debits = 0;
    
    state.mergedData.forEach(item => {
        // Exclude selected pending entries from live calculation (virtual match)
        if (!item.reconciled && !item.matchGroupId && !state.selectedIds.has(item.id)) {
            const tType = getFileCategory(item);
            if (tType === 'HDFC') {
                hdfc_pending_credits += item.creditTrn;
                hdfc_pending_debits += item.debitTrn;
            } else {
                gl_pending_credits += item.creditTrn;
                gl_pending_debits += item.debitTrn;
            }
        }
    });
    
    // BRS calculations
    const val_row5 = gl_345051_bal;
    const val_row6 = val_row5 - gl_pending_credits + gl_pending_debits;
    const val_row7 = val_row6 - hdfc_pending_credits + hdfc_pending_debits;
    const diff_any = hdfc_bal + val_row7;
    const other_gls_sum = gl_3496_bal - gl_3493_bal;
    const total_unreconciled_diff = diff_any - other_gls_sum;
    
    function formatBrsLiveNum(val) {
        return new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(val);
    }
    
    // Generate live BRS table HTML matching the user's template
    let tableHtml = `
        <table style="width:100%; border-collapse:collapse; font-size:10px; line-height:1.4; color:var(--text-primary); margin-top:8px;">
            <tbody>
                <!-- Row 1: Opening GL Balance -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
                    <td style="padding:4px 0; color:var(--text-secondary);">GL 345051 Balance</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; font-weight:600; color:${val_row5 < 0 ? 'var(--danger)' : 'var(--text-primary)'};">${formatBrsLiveNum(val_row5)}</td>
                </tr>
                <!-- Row 2: GL Pending Credits -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); background:${gl_pending_credits > 0 ? 'rgba(239, 68, 68, 0.05)' : 'transparent'};">
                    <td style="padding:4px 0; padding-left:8px; color:var(--text-muted);">(-) GL Credit Pending</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; font-weight:600; color:var(--danger);">${gl_pending_credits > 0 ? formatBrsLiveNum(gl_pending_credits) : '0.00'}</td>
                </tr>
                <!-- Row 3: GL Pending Debits -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); background:${gl_pending_debits > 0 ? 'rgba(16, 185, 129, 0.05)' : 'transparent'};">
                    <td style="padding:4px 0; padding-left:8px; color:var(--text-muted);">(+) GL Debit Pending</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; font-weight:600; color:var(--success);">${gl_pending_debits > 0 ? formatBrsLiveNum(gl_pending_debits) : '0.00'}</td>
                </tr>
                <!-- Row 4: GL Subtotal -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); font-weight:600;">
                    <td style="padding:4px 0; color:var(--text-secondary);">Subtotal (GL Side)</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; color:${val_row6 < 0 ? 'var(--danger)' : 'var(--text-primary)'};">${formatBrsLiveNum(val_row6)}</td>
                </tr>
                <!-- Row 5: HDFC Pending Credits -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); background:${hdfc_pending_credits > 0 ? 'rgba(16, 185, 129, 0.05)' : 'transparent'};">
                    <td style="padding:4px 0; padding-left:8px; color:var(--text-muted);">(+) HDFC Credit Pending</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; font-weight:600; color:var(--success);">${hdfc_pending_credits > 0 ? formatBrsLiveNum(hdfc_pending_credits) : '0.00'}</td>
                </tr>
                <!-- Row 6: HDFC Pending Debits -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); background:${hdfc_pending_debits > 0 ? 'rgba(239, 68, 68, 0.05)' : 'transparent'};">
                    <td style="padding:4px 0; padding-left:8px; color:var(--text-muted);">(-) HDFC Debit Pending</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; font-weight:600; color:var(--danger);">${hdfc_pending_debits > 0 ? formatBrsLiveNum(hdfc_pending_debits) : '0.00'}</td>
                </tr>
                <!-- Row 7: GL Adjusted Balance -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); font-weight:600;">
                    <td style="padding:4px 0; color:var(--text-secondary);">GL Adjusted Balance</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; color:${val_row7 < 0 ? 'var(--danger)' : 'var(--text-primary)'};">${formatBrsLiveNum(val_row7)}</td>
                </tr>
                <!-- Row 8: HDFC Bank Balance -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
                    <td style="padding:4px 0; color:var(--text-secondary);">HDFC Bank Balance</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; font-weight:600; color:${hdfc_bal < 0 ? 'var(--danger)' : 'var(--text-primary)'};">${formatBrsLiveNum(hdfc_bal)}</td>
                </tr>
                <!-- Row 9: Intermediate Diff -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); font-weight:600;">
                    <td style="padding:4px 0; color:var(--text-secondary);">Intermediate Diff</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; color:${diff_any < 0 ? 'var(--danger)' : 'var(--text-primary)'};">${formatBrsLiveNum(diff_any)}</td>
                </tr>
                <!-- Row 10: GL 3493 -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
                    <td style="padding:4px 0; padding-left:8px; color:var(--text-muted);">GL 3493 Balance</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; color:${gl_3493_bal < 0 ? 'var(--danger)' : 'var(--text-primary)'};">${formatBrsLiveNum(gl_3493_bal)}</td>
                </tr>
                <!-- Row 11: GL 3496 Balance -->
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
                    <td style="padding:4px 0; padding-left:8px; color:var(--text-muted);">GL 3496 Balance</td>
                    <td style="text-align:right; padding:4px 0; font-family:monospace; color:${gl_3496_bal < 0 ? 'var(--danger)' : 'var(--text-primary)'};">${formatBrsLiveNum(gl_3496_bal)}</td>
                </tr>
                <!-- Row 12: Final Net Diff -->
                <tr style="font-weight:700; border-top:1px solid rgba(255,255,255,0.15);">
                    <td style="padding:6px 0; color:var(--text-primary); font-size:11px;">Final Net Difference</td>
                    <td id="widget-net-diff" onclick="showDiffAnalysisModal()" style="text-align:right; padding:6px 4px; font-family:monospace; font-size:12px; font-weight:800; border-radius:4px; transition: all 0.2s ease; cursor:pointer;" title="તફાવતનું વિગતવાર વિશ્લેષણ જોવા માટે ક્લિક કરો">
                        ${formatBrsLiveNum(total_unreconciled_diff)}
                    </td>
                </tr>
            </tbody>
        </table>
    `;
    
    // Calculate sums of currently SELECTED (checked) items
    let selected_hdfc = 0;
    let selected_gl = 0;
    
    state.mergedData.forEach(item => {
        if (state.selectedIds.has(item.id)) {
            const tType = getFileCategory(item);
            if (tType === 'HDFC') {
                selected_hdfc += (item.creditTrn - item.debitTrn);
            } else {
                selected_gl += (item.debitTrn - item.creditTrn);
            }
        }
    });
    
    const selection_diff = Math.abs(selected_hdfc - selected_gl);
    let selectionHtml = '';
    if (state.selectedIds.size > 0) {
        selectionHtml = `
            <div style="margin-top:12px; padding:10px; background:rgba(139, 92, 246, 0.08); border:1px solid rgba(139, 92, 246, 0.3); border-radius:6px;">
                <span style="font-weight:700; color:#8b5cf6; font-size:10px; display:block; margin-bottom:6px; text-transform:uppercase;">
                    ✓ પસંદ કરેલ મેળવણી (Selected Match)
                </span>
                <div style="display:flex; justify-content:space-between; font-size:10px; margin-bottom:4px;">
                    <span style="color:var(--text-secondary);">પસંદ કરેલ HDFC:</span>
                    <span style="font-family:monospace; font-weight:600;">₹ ${formatBrsLiveNum(selected_hdfc)}</span>
                </div>
                <div style="display:flex; justify-content:space-between; font-size:10px; margin-bottom:4px;">
                    <span style="color:var(--text-secondary);">પસંદ કરેલ GL:</span>
                    <span style="font-family:monospace; font-weight:600;">₹ ${formatBrsLiveNum(selected_gl)}</span>
                </div>
                <div style="border-top:1px dashed rgba(139, 92, 246, 0.2); margin:4px 0;"></div>
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-weight:700; font-size:10px; color:var(--text-primary);">પસંદગી તફાવત:</span>
                    <span style="font-family:monospace; font-weight:800; font-size:11px; padding:2px 6px; border-radius:4px;
                        color:${selection_diff < 0.01 ? 'var(--success)' : 'var(--danger)'};
                        background:${selection_diff < 0.01 ? 'rgba(16, 185, 129, 0.1)' : 'rgba(239, 68, 68, 0.1)'};">
                        ₹ ${formatBrsLiveNum(selection_diff)}
                    </span>
                </div>
            </div>
        `;
    }
    
    container.innerHTML = tableHtml + selectionHtml;
    
    if (valTotalBal) {
        valTotalBal.textContent = formatCurrency(total_unreconciled_diff);
        if (total_unreconciled_diff < 0) {
            valTotalBal.classList.add('negative-balance');
        } else {
            valTotalBal.classList.remove('negative-balance');
        }
    }
    
    const successAlert = document.getElementById('brs-success-alert');
    if (Math.abs(total_unreconciled_diff) < 0.01) {
        if (successAlert) successAlert.style.display = 'flex';
    } else {
        if (successAlert) successAlert.style.display = 'none';
    }
}

// Open modal showing breakdown of the net unreconciled difference
function showDiffAnalysisModal() {
    const formatBrsLiveNum = (val) => {
        if (val === null || val === undefined || isNaN(val)) return '0.00';
        return new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(val);
    };
    
    const pendingHdfc = state.mergedData.filter(item => getFileCategory(item) === 'HDFC' && !item.reconciled && !item.matchGroupId);
    const pendingGl = state.mergedData.filter(item => getFileCategory(item) !== 'HDFC' && !item.reconciled && !item.matchGroupId);
    
    // Sort descending by amount
    const getAmt = (item) => Math.max(item.creditTrn, item.debitTrn);
    pendingHdfc.sort((a, b) => getAmt(b) - getAmt(a));
    pendingGl.sort((a, b) => getAmt(b) - getAmt(a));
    
    let gl_pending_credits = 0;
    let gl_pending_debits = 0;
    let hdfc_pending_credits = 0;
    let hdfc_pending_debits = 0;
    
    state.mergedData.forEach(item => {
        if (!item.reconciled && !item.matchGroupId) {
            const tType = getFileCategory(item);
            if (tType === 'HDFC') {
                hdfc_pending_credits += item.creditTrn;
                hdfc_pending_debits += item.debitTrn;
            } else {
                gl_pending_credits += item.creditTrn;
                gl_pending_debits += item.debitTrn;
            }
        }
    });
    
    const hdfc_bal = (state.files && state.files.HDFC && state.files.HDFC.loaded) ? state.files.HDFC.balance : 0;
    const gl_345051_bal = (state.files && state.files['345051'] && state.files['345051'].loaded) ? state.files['345051'].balance : 0;
    const gl_3493_bal = (state.files && state.files['3493'] && state.files['3493'].loaded) ? Math.abs(state.files['3493'].balance) : 0;
    const gl_3496_bal = (state.files && state.files['3496'] && state.files['3496'].loaded) ? Math.abs(state.files['3496'].balance) : 0;
    
    const val_row5 = gl_345051_bal;
    const val_row6 = val_row5 - gl_pending_credits + gl_pending_debits;
    const val_row7 = val_row6 - hdfc_pending_credits + hdfc_pending_debits;
    const diff_any = hdfc_bal + val_row7;
    const other_gls_sum = gl_3496_bal - gl_3493_bal;
    const total_unreconciled_diff = diff_any - other_gls_sum;
    
    const modalId = 'diff-analysis-modal';
    let modal = document.getElementById(modalId);
    if (!modal) {
        modal = document.createElement('div');
        modal.id = modalId;
        modal.style.position = 'fixed';
        modal.style.top = '0';
        modal.style.left = '0';
        modal.style.width = '100%';
        modal.style.height = '100%';
        modal.style.backgroundColor = 'rgba(0, 0, 0, 0.75)';
        modal.style.display = 'flex';
        modal.style.justifyContent = 'center';
        modal.style.alignItems = 'center';
        modal.style.zIndex = '9999';
        document.body.appendChild(modal);
    }
    
    // 1. Find combinations of 1 or 2 pending transactions whose sum is exactly equal to the target diff
    const targetDiff = Math.abs(total_unreconciled_diff);
    let suggestionHtml = '';
    let foundSolutions = [];
    
    // Check single items
    for (let i = 0; i < pendingHdfc.length; i++) {
        const a = pendingHdfc[i];
        if (Math.abs(getAmt(a) - targetDiff) < 0.05) {
            foundSolutions.push([a]);
        }
    }
    for (let i = 0; i < pendingGl.length; i++) {
        const a = pendingGl[i];
        if (Math.abs(getAmt(a) - targetDiff) < 0.05) {
            foundSolutions.push([a]);
        }
    }
    
    // Check pairs of HDFC
    if (foundSolutions.length === 0) {
        for (let i = 0; i < pendingHdfc.length; i++) {
            for (let j = i + 1; j < pendingHdfc.length; j++) {
                const a = pendingHdfc[i];
                const b = pendingHdfc[j];
                const sum = getAmt(a) + getAmt(b);
                if (Math.abs(sum - targetDiff) < 0.05) {
                    foundSolutions.push([a, b]);
                }
            }
        }
    }
    // Check pairs of GL
    if (foundSolutions.length === 0) {
        for (let i = 0; i < pendingGl.length; i++) {
            for (let j = i + 1; j < pendingGl.length; j++) {
                const a = pendingGl[i];
                const b = pendingGl[j];
                const sum = getAmt(a) + getAmt(b);
                if (Math.abs(sum - targetDiff) < 0.05) {
                    foundSolutions.push([a, b]);
                }
            }
        }
    }
    
    if (foundSolutions.length > 0) {
        const solution = foundSolutions[0];
        const itemsHtml = solution.map(item => {
            const cat = getFileCategory(item);
            const amt = getAmt(item);
            const isCredit = item.creditTrn > 0;
            return `
                <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.05); padding:8px 12px; border-radius:6px; margin-bottom:6px; border-left:4px solid ${cat === 'HDFC' ? '#60a5fa' : '#34d399'};">
                    <div>
                        <span style="font-size:10px; color:var(--text-muted); display:block;">${item.date} | ${cat} (${isCredit ? 'ક્રેડિટ (+)' : 'ડેબિટ (-)'})</span>
                        <span style="font-size:11px; font-weight:600; color:var(--text-primary); display:block; max-width:400px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${item.description}</span>
                    </div>
                    <span style="font-family:monospace; font-weight:700; color:${isCredit ? 'var(--success)' : 'var(--danger)'};">₹ ${formatBrsLiveNum(amt)}</span>
                </div>
            `;
        }).join('');
        
        suggestionHtml = `
            <div style="background:rgba(139, 92, 246, 0.1); border:1px solid rgba(139, 92, 246, 0.3); border-radius:8px; padding:12px; margin-bottom:16px;">
                <span style="font-weight:700; color:#c084fc; font-size:11px; display:block; margin-bottom:8px; text-transform:uppercase;">
                    💡 સંભવિત તફાવતનું મુખ્ય કારણ (Matched Pending Transactions):
                </span>
                ${itemsHtml}
                <div style="font-size:10px; color:var(--text-secondary); margin-top:6px;">
                    * આ એન્ટ્રીઓ હજુ વણમેળવાયેલી (Pending) હોવાથી તમારો તફાવત બરાબર <strong>₹ ${formatBrsLiveNum(targetDiff)}</strong> નો આવી રહ્યો છે.
                </div>
            </div>
        `;
    }
    
    // 2. Find opposite-side same-amount pending pairs
    const sameAmtPairs = [];
    const matchedGlIds = new Set();
    pendingHdfc.forEach(h => {
        const hAmt = getAmt(h);
        const match = pendingGl.find(g => {
            if (matchedGlIds.has(g.id)) return false;
            const isOppositeSide = (h.creditTrn > 0 && g.debitTrn > 0) || (h.debitTrn > 0 && g.creditTrn > 0);
            return Math.abs(hAmt - getAmt(g)) < 0.05 && isOppositeSide;
        });
        if (match) {
            sameAmtPairs.push({ hdfc: h, gl: match, amt: hAmt });
            matchedGlIds.add(match.id);
        }
    });
    
    let pairsHtml = '';
    if (sameAmtPairs.length > 0) {
        const pairsListHtml = sameAmtPairs.slice(0, 5).map(pair => `
            <div style="background:rgba(245, 158, 11, 0.05); border:1px solid rgba(245, 158, 11, 0.2); border-radius:8px; padding:10px; margin-bottom:8px;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                    <span style="font-weight:700; color:#f59e0b; font-size:10px; text-transform:uppercase;">⚖️ સંભવિત મેળ ખાતી જોડી (Same Amount Pair)</span>
                    <span style="font-family:monospace; font-weight:700; color:var(--text-primary);">₹ ${formatBrsLiveNum(pair.amt)}</span>
                </div>
                <div style="display:flex; gap:12px; font-size:11px;">
                    <div style="flex:1; border-right:1px dashed rgba(255,255,255,0.1); padding-right:8px;">
                        <span style="color:#60a5fa; font-weight:600; display:block; font-size:9.5px;">🏦 HDFC (${pair.hdfc.date})</span>
                        <span style="color:var(--text-secondary); display:block; max-width:260px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${pair.hdfc.description}">${pair.hdfc.description}</span>
                    </div>
                    <div style="flex:1; padding-left:4px;">
                        <span style="color:#34d399; font-weight:600; display:block; font-size:9.5px;">📖 GL (${pair.gl.date})</span>
                        <span style="color:var(--text-secondary); display:block; max-width:260px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${pair.gl.description}">${pair.gl.description}</span>
                    </div>
                </div>
            </div>
        `).join('');
        
        pairsHtml = `
            <div style="border:1px solid rgba(245, 158, 11, 0.3); border-radius:8px; padding:12px; margin-bottom:16px; background:rgba(245, 158, 11, 0.02);">
                <h4 style="margin:0 0 10px 0; font-size:13px; color:#f59e0b; border-bottom:1px solid rgba(245, 158, 11, 0.2); padding-bottom:4px;">
                    🔗 સંભવિત મેચિંગ જોડીઓ (Unmatched Same Amounts - Manual check required)
                </h4>
                ${pairsListHtml}
                <div style="font-size:10px; color:var(--text-secondary); margin-top:6px;">
                    * આ એન્ટ્રીઓની રકમ સેમ હોવા છતાં અલગ-અલગ વિગતોને લીધે મેચ નથી થઈ. તમે તેને સિલેક્ટ કરી મેન્યુઅલ મેચ કરી શકો છો.
                </div>
            </div>
        `;
    }
    
    const topHdfcHtml = pendingHdfc.map(item => `
        <tr style="border-bottom:1px solid rgba(255,255,255,0.05); font-size:11px;">
            <td style="padding:6px; color:var(--text-secondary);">${item.date}</td>
            <td style="padding:6px; max-width:280px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${item.description}">${item.description}</td>
            <td style="padding:6px; text-align:right; color:var(--success); font-family:monospace;">${item.creditTrn > 0 ? formatBrsLiveNum(item.creditTrn) : '-'}</td>
            <td style="padding:6px; text-align:right; color:var(--danger); font-family:monospace;">${item.debitTrn > 0 ? formatBrsLiveNum(item.debitTrn) : '-'}</td>
        </tr>
    `).join('') || `<tr><td colspan="4" style="text-align:center; padding:12px; color:var(--text-muted);">કોઈ બાકી વ્યવહાર નથી</td></tr>`;
    
    const topGlHtml = pendingGl.map(item => `
        <tr style="border-bottom:1px solid rgba(255,255,255,0.05); font-size:11px;">
            <td style="padding:6px; color:var(--text-secondary);">${item.date}</td>
            <td style="padding:6px; max-width:280px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${item.description}">${item.description}</td>
            <td style="padding:6px; text-align:right; color:var(--success); font-family:monospace;">${item.creditTrn > 0 ? formatBrsLiveNum(item.creditTrn) : '-'}</td>
            <td style="padding:6px; text-align:right; color:var(--danger); font-family:monospace;">${item.debitTrn > 0 ? formatBrsLiveNum(item.debitTrn) : '-'}</td>
        </tr>
    `).join('') || `<tr><td colspan="4" style="text-align:center; padding:12px; color:var(--text-muted);">કોઈ બાકી વ્યવહાર નથી</td></tr>`;
    
    modal.innerHTML = `
        <div style="background:#111827; border:1px solid rgba(255,255,255,0.1); border-radius:12px; width:95%; max-width:850px; max-height:85vh; overflow-y:auto; padding:24px; color:var(--text-primary); box-shadow:0 20px 25px -5px rgba(0,0,0,0.5);">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(255,255,255,0.1); padding-bottom:12px; margin-bottom:16px;">
                <h3 style="margin:0; font-size:16px; font-weight:700; color:#a78bfa; display:flex; align-items:center; gap:8px;">
                    📊 BRS તફાવત વિશ્લેષણ (Difference Analysis)
                </h3>
                <button onclick="document.getElementById('${modalId}').remove()" style="background:none; border:none; color:var(--text-secondary); font-size:24px; cursor:pointer; line-height:1;">&times;</button>
            </div>
            
            <div style="background:rgba(239, 68, 68, 0.05); border:1px solid rgba(239, 68, 68, 0.2); border-radius:8px; padding:12px; margin-bottom:16px;">
                <span style="font-size:11px; color:var(--text-muted);">આખરી નેટ તફાવત (Final Net Difference)</span>
                <div style="font-size:24px; font-weight:800; color:var(--danger); font-family:monospace; margin-top:4px;">
                    ₹ ${formatBrsLiveNum(total_unreconciled_diff)}
                </div>
            </div>
            
            ${suggestionHtml}
            
            ${pairsHtml}
            
            <div style="display:grid; grid-template-columns:1fr; gap:16px; margin-bottom:16px;">
                <!-- HDFC Pending -->
                <div style="border:1px solid rgba(255,255,255,0.08); border-radius:8px; padding:12px; background:rgba(255,255,255,0.01);">
                    <h4 style="margin:0 0 8px 0; font-size:13px; color:#60a5fa; border-bottom:1px solid rgba(96,165,250,0.2); padding-bottom:4px; display:flex; justify-content:space-between;">
                        <span>🏦 HDFC બેંકના બાકી વ્યવહારો (HDFC Pending - કુલ ${pendingHdfc.length} વ્યવહારો)</span>
                        <span style="font-size:11px; font-family:monospace; color:var(--text-secondary);">કુલ ક્રેડિટ: ${formatBrsLiveNum(hdfc_pending_credits)} | કુલ ડેબિટ: ${formatBrsLiveNum(hdfc_pending_debits)}</span>
                    </h4>
                    <div style="max-height: 250px; overflow-y: auto; border: 1px solid rgba(255,255,255,0.05); border-radius: 6px;">
                        <table style="width:100%; border-collapse:collapse;">
                            <thead>
                                <tr style="font-size:10px; color:var(--text-muted); text-align:left; border-bottom:1px solid rgba(255,255,255,0.1); background:rgba(0,0,0,0.2); position:sticky; top:0; z-index:1;">
                                    <th style="padding:6px; width:15%;">તારીખ</th>
                                    <th style="padding:6px; width:55%;">વિગત</th>
                                    <th style="padding:6px; text-align:right; width:15%;">ક્રેડિટ (+)</th>
                                    <th style="padding:6px; text-align:right; width:15%;">ડેબિટ (-)</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${topHdfcHtml}
                            </tbody>
                        </table>
                    </div>
                </div>
                
                <!-- GL Pending -->
                <div style="border:1px solid rgba(255,255,255,0.08); border-radius:8px; padding:12px; background:rgba(255,255,255,0.01);">
                    <h4 style="margin:0 0 8px 0; font-size:13px; color:#34d399; border-bottom:1px solid rgba(52,211,153,0.2); padding-bottom:4px; display:flex; justify-content:space-between;">
                        <span>📖 ચોપડે બાકી વ્યવહારો (GL Pending - કુલ ${pendingGl.length} વ્યવહારો)</span>
                        <span style="font-size:11px; font-family:monospace; color:var(--text-secondary);">કુલ ક્રેડિટ: ${formatBrsLiveNum(gl_pending_credits)} | કુલ ડેબિટ: ${formatBrsLiveNum(gl_pending_debits)}</span>
                    </h4>
                    <div style="max-height: 250px; overflow-y: auto; border: 1px solid rgba(255,255,255,0.05); border-radius: 6px;">
                        <table style="width:100%; border-collapse:collapse;">
                            <thead>
                                <tr style="font-size:10px; color:var(--text-muted); text-align:left; border-bottom:1px solid rgba(255,255,255,0.1); background:rgba(0,0,0,0.2); position:sticky; top:0; z-index:1;">
                                    <th style="padding:6px; width:15%;">તારીખ</th>
                                    <th style="padding:6px; width:55%;">વિગત</th>
                                    <th style="padding:6px; text-align:right; width:15%;">ક્રેડિટ (+)</th>
                                    <th style="padding:6px; text-align:right; width:15%;">ડેબિટ (-)</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${topGlHtml}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div style="display:flex; justify-content:flex-end;">
                <button onclick="document.getElementById('${modalId}').remove()" style="background:#4c1d95; border:none; color:white; padding:8px 16px; border-radius:6px; font-weight:600; cursor:pointer; font-size:12px; transition:background 0.2s;" onmouseover="this.style.background='#5b21b6'" onmouseout="this.style.background='#4c1d95'">
                    બંધ કરો (Close)
                </button>
            </div>
        </div>
    `;
}

// Helper to calculate the current live BRS difference
function calculateCurrentDifference() {
    const hdfc_bal = (state.files && state.files.HDFC && state.files.HDFC.loaded) ? state.files.HDFC.balance : 0;
    const gl_345051_bal = (state.files && state.files['345051'] && state.files['345051'].loaded) ? state.files['345051'].balance : 0;
    const gl_3493_bal = (state.files && state.files['3493'] && state.files['3493'].loaded) ? Math.abs(state.files['3493'].balance) : 0;
    const gl_3496_bal = (state.files && state.files['3496'] && state.files['3496'].loaded) ? Math.abs(state.files['3496'].balance) : 0;
    
    let gl_pending_credits = 0;
    let gl_pending_debits = 0;
    let hdfc_pending_credits = 0;
    let hdfc_pending_debits = 0;
    
    state.mergedData.forEach(item => {
        if (!item.reconciled && !item.matchGroupId) {
            const tType = getFileCategory(item);
            if (tType === 'HDFC') {
                hdfc_pending_credits += item.creditTrn;
                hdfc_pending_debits += item.debitTrn;
            } else {
                gl_pending_credits += item.creditTrn;
                gl_pending_debits += item.debitTrn;
            }
        }
    });
    
    const val_row5 = gl_345051_bal;
    const val_row6 = val_row5 - gl_pending_credits + gl_pending_debits;
    const val_row7 = val_row6 - hdfc_pending_credits + hdfc_pending_debits;
    const diff_any = hdfc_bal + val_row7;
    const other_gls_sum = gl_3496_bal - gl_3493_bal;
    return diff_any - other_gls_sum;
}

// Restore state from localStorage
function loadStateFromLocalStorage() {
    const savedDate = localStorage.getItem('recon_current_date');
    if (savedDate) {
        state.currentDate = savedDate;
        if (currentDateInput) currentDateInput.value = savedDate;
    }

    const savedMergedData = localStorage.getItem('recon_merged_data');
    const savedMatchCounter = localStorage.getItem('recon_match_counter');
    const savedFilesMeta = localStorage.getItem('recon_files_meta');
    
    if (savedMergedData) {
        const parsed = JSON.parse(savedMergedData);
        if (parsed && parsed.length > 0 && 'd' in parsed[0]) {
            state.mergedData = decompressMergedData(parsed);
        } else {
            state.mergedData = parsed || [];
        }
        state.matchGroupCounter = parseInt(savedMatchCounter) || 0;
    }
    
    if (savedFilesMeta) {
        state.files = JSON.parse(savedFilesMeta);
        Object.keys(state.files).forEach(fileType => {
            if (state.files[fileType].loaded) {
                const card = document.getElementById(`status-${fileType.toLowerCase()}`);
                if (card) {
                    card.classList.add('loaded');
                    card.querySelector('.file-indicator').className = 'file-indicator success';
                    const rowCount = state.files[fileType].rowCount !== undefined 
                        ? state.files[fileType].rowCount 
                        : (state.files[fileType].data ? state.files[fileType].data.length : 0);
                    card.querySelector('.file-meta').textContent = `સક્રિય | ${rowCount} રોઝ | ${state.files[fileType].rawName}`;
                    
                    // Show clear button
                    const trashBtn = card.querySelector('.clear-file-btn');
                    if (trashBtn) trashBtn.style.display = 'flex';
                }
                
                if (fileType !== 'history') {
                    const balVal = document.getElementById(`val-${fileType.toLowerCase()}-bal`);
                    if (balVal) {
                        const balance = state.files[fileType].balance;
                        balVal.textContent = formatCurrency(balance);
                        if (balance < 0) {
                            balVal.classList.add('negative-balance');
                        } else {
                            balVal.classList.remove('negative-balance');
                        }
                    }
                }
            }
        });
        
        let totalBal = 0;
        Object.values(state.files).forEach(f => {
            if (f.loaded) totalBal += f.balance;
        });
        valTotalBal.textContent = formatCurrency(totalBal);
        if (totalBal < 0) {
            valTotalBal.classList.add('negative-balance');
        } else {
            valTotalBal.classList.remove('negative-balance');
        }
    }
    
    if (state.mergedData.length > 0) {
        btnAutoReconcile.disabled = false;
        btnExportPending.disabled = false;
        btnExportReconciled.disabled = false;
        updateBRSLiveWidget();
        refreshLedgerCounts();
        renderTable();
    }
}

// Format Numbers to Indian Currency Style (Lakh/Crore)
function formatCurrency(val) {
    if (isNaN(val) || val === null) return '₹ 0.00';
    const isNegative = val < 0;
    const absVal = Math.abs(val);
    
    // Format to Indian style
    let x = absVal.toFixed(2);
    let lastThree = x.substring(x.length - 6);
    let otherNumbers = x.substring(0, x.length - 6);
    if (otherNumbers !== '') {
        lastThree = ',' + lastThree;
    }
    let res = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ",") + lastThree;
    return (isNegative ? '-₹ ' : '₹ ') + res;
}

// Parse various date strings into Date Objects (ignoring time)
function parseDateString(dateStr) {
    if (!dateStr) return null;
    dateStr = dateStr.trim();
    
    // Strip time if present
    if (dateStr.includes(' ')) {
        dateStr = dateStr.split(' ')[0];
    }
    if (dateStr.includes('T')) {
        dateStr = dateStr.split('T')[0];
    }
    
    // Clean separators to '-'
    dateStr = dateStr.replace(/[\/\.]/g, '-');
    
    // 1. Check DD-MM-YYYY or DD-MM-YY
    let match = dateStr.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
    if (match) {
        let d = parseInt(match[1]);
        let m = parseInt(match[2]) - 1;
        let y = parseInt(match[3]);
        if (y < 100) {
            y = y < 50 ? 2000 + y : 1900 + y;
        }
        return new Date(y, m, d);
    }
    
    // 2. Check YYYY-MM-DD
    match = dateStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (match) {
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
    }
    
    // 3. Check DD-MMM-YY or DD-MMM-YYYY (e.g. 22-Jun-26 or 22-Jun-2026)
    match = dateStr.match(/^(\d{1,2})-([A-Za-z]{3,9})-(\d{2,4})$/);
    if (match) {
        const months = {
            jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
            january: 0, february: 1, march: 2, april: 3, june: 5, july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
        };
        let d = parseInt(match[1]);
        let mStr = match[2].toLowerCase();
        let m = months[mStr] !== undefined ? months[mStr] : 0;
        let y = parseInt(match[3]);
        if (y < 100) {
            y = y < 50 ? 2000 + y : 1900 + y;
        }
        return new Date(y, m, d);
    }
    
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? null : d;
}

// Format Date Object to DD-MM-YYYY
function formatDateOnly(dateObj) {
    if (!dateObj) return '';
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    return `${day}-${month}-${year}`;
}

// Format Date Object or Date String to DD/MM/YYYY
function formatDateToSlash(val) {
    if (!val) return '';
    let dateObj = val;
    if (typeof val === 'string') {
        dateObj = parseDateString(val);
    }
    if (!dateObj || isNaN(dateObj.getTime())) return typeof val === 'string' ? val : '';
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    return `${day}/${month}/${year}`;
}

// Calculate days between two dates
function calculateDaysDifference(date1, date2) {
    if (!date1 || !date2) return '';
    const timeDiff = date1.getTime() - date2.getTime();
    return Math.floor(timeDiff / (1000 * 3600 * 24));
}

// Recalculate DAY field for all entries based on the current system date (Today)
function recalculateDays() {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Reset time components of today
    
    state.mergedData.forEach(item => {
        const itemDate = parseDateString(item.date);
        if (itemDate) {
            itemDate.setHours(0, 0, 0, 0); // Reset time components of transaction date
            const timeDiff = today.getTime() - itemDate.getTime();
            const diffDays = Math.floor(timeDiff / (1000 * 3600 * 24));
            item.day = diffDays >= 0 ? diffDays : 0;
        } else {
            item.day = 0;
        }
    });
}

// Handle Multiple Files Upload
function handleFiles(fileList) {
    let filesLoadedInBatch = 0;
    
    for (let i = 0; i < fileList.length; i++) {
        const file = fileList[i];
        const name = file.name.toUpperCase();
        
        // Check if it's an Excel sheet
        if (name.endsWith('.XLSX') || name.endsWith('.XLS')) {
            processExcelFile(file);
            continue;
        }
        
        let fileType = null;
        if (name.includes('HDFC')) {
            fileType = 'HDFC';
        } else if (name.includes('345051')) {
            fileType = '345051';
        } else if (name.includes('3493')) {
            fileType = '3493';
        } else if (name.includes('3496')) {
            fileType = '3496';
        } else if (name.includes('HISTORY') || name.includes('PENDING_HISTORY') || name.includes('PENDING')) {
            fileType = 'history';
        }
        
        if (fileType) {
            filesLoadedInBatch++;
            processCSV(file, fileType);
        } else {
            showToast(`અમાન્ય ફાઇલ નામ: ${file.name}. કૃપા કરીને HDFC, 345051, 3493, 3496 અથવા PENDING_HISTORY ધરાવતી ફાઇલ અપલોડ કરો.`, 'error');
        }
    }
}

// Read and Parse CSV File
// Read and Parse CSV File
function processCSV(file, fileType) {
    try {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                let text = e.target.result;
                
                // Detect UTF-16 LE encoding (check for null bytes)
                if (text.includes('\u0000')) {
                    console.log(`[Recon] Null bytes detected in ${file.name}, re-reading as UTF-16LE...`);
                    const utf16Reader = new FileReader();
                    utf16Reader.onload = function(evt) {
                        try {
                            processCSVContent(evt.target.result, file, fileType);
                        } catch (err) {
                            alert("UTF-16 content processing error:\n" + err.message + "\nStack: " + err.stack);
                        }
                    };
                    utf16Reader.readAsText(file, 'utf-16le');
                    return;
                }
                
                processCSVContent(text, file, fileType);
            } catch (err) {
                alert("CSV read callback error:\n" + err.message + "\nStack: " + err.stack);
            }
        };
        reader.readAsText(file);
    } catch (err) {
        alert("FileReader startup error:\n" + err.message + "\nStack: " + err.stack);
    }
}

// Process CSV content string
function processCSVContent(text, file, fileType) {
    const lines = text.split(/\r\n|\r|\n/);
    let headerLineIndex = -1;
    let balance = 0;
    
    if (fileType === 'history') {
        headerLineIndex = 0;
    } else if (fileType === 'HDFC') {
        // Robust check: Scan lines to find header index for HDFC
        for (let idx = 0; idx < Math.min(lines.length, 50); idx++) {
            const cleanLine = lines[idx].toLowerCase().trim();
            
            const hasDate = cleanLine.includes('date') || cleanLine.includes('txn') || cleanLine.includes('post');
            const hasDesc = cleanLine.includes('narration') || cleanLine.includes('particular') || cleanLine.includes('desc') || cleanLine.includes('detail') || cleanLine.includes('remark') || cleanLine.includes('ref') || cleanLine.includes('chq');
            const hasMoney = cleanLine.includes('amt') || cleanLine.includes('withdrawal') || cleanLine.includes('deposit') || cleanLine.includes('amount') || cleanLine.includes('credit') || cleanLine.includes('debit') || cleanLine.includes('balance') || cleanLine.includes('value');
            
            if (hasDate && hasDesc && hasMoney) {
                headerLineIndex = idx;
                break;
            }
        }
        // Fallback to row 17 (index 16) if not detected, but ensure we don't exceed lines length
        if (headerLineIndex === -1) {
            headerLineIndex = lines.length > 16 ? 16 : 0;
        }
        
        // Closing balance scanner: Scan first 15 lines of HDFC for "Closing Balance"
        let foundBalance = false;
        for (let idx = 0; idx < Math.min(lines.length, 15); idx++) {
            const balanceRow = lines[idx];
            if (balanceRow) {
                const match = balanceRow.match(/Closing\s+Balance\s*:\s*([\d,.-]+)/i);
                if (match) {
                    balance = parseFloat(match[1].replace(/,/g, '')) || 0;
                    foundBalance = true;
                    break;
                }
            }
        }
        // Fallback to row 7 (index 6) if not found by scanner
        if (!foundBalance && lines.length > 6) {
            const balanceRow = lines[6];
            if (balanceRow) {
                const match = balanceRow.match(/Closing\s+Balance\s*:\s*([\d,.-]+)/i);
                if (match) {
                    balance = parseFloat(match[1].replace(/,/g, '')) || 0;
                }
            }
        }
    } else {
        // Scan for other files headers
        for (let idx = 0; idx < Math.min(lines.length, 25); idx++) {
            const cleanLine = lines[idx].toLowerCase().trim();
            if (cleanLine.includes('postdate') || cleanLine.includes('post date') || cleanLine.includes('usernarration') || cleanLine.includes('narration') || cleanLine.includes('textbox81')) {
                headerLineIndex = idx;
                break;
            }
        }
        if (headerLineIndex === -1) {
            headerLineIndex = lines.length > 6 ? 6 : 0;
        }
    }
    
    // Parse table contents starting from headers
    if (lines.length > headerLineIndex) {
        const csvDataText = lines.slice(headerLineIndex).join('\n');
        
        Papa.parse(csvDataText, {
            header: true,
            skipEmptyLines: true,
            dynamicTyping: false,
            complete: function(results) {
                try {
                    let rows = results.data;
                    
                    // Log details to console for debugging column mismatches
                    console.log(`Parsed ${fileType} CSV. Headers:`, results.meta.fields);
                    console.log(`First row:`, rows[0]);
                    
                    // Extract balance for other files from Textbox81 in the last row
                    if (fileType !== 'HDFC' && fileType !== 'history' && rows.length > 0) {
                        // Find last row with non-empty Textbox81
                        let lastRowWithBalance = null;
                        for (let j = rows.length - 1; j >= 0; j--) {
                            const val = getRowValue(rows[j], ['Textbox81']);
                            if (val && val.trim() !== '') {
                                lastRowWithBalance = rows[j];
                                break;
                            }
                        }
                        
                        if (lastRowWithBalance) {
                            let balStr = getRowValue(lastRowWithBalance, ['Textbox81']).trim();
                            let isDebit = balStr.toUpperCase().endsWith('DR');
                            balStr = balStr.replace(/CR|DR/gi, '').replace(/,/g, '').trim();
                            balance = parseFloat(balStr) || 0;
                            if (isDebit) {
                                balance = -balance;
                            }
                        }
                    }
                    
                    // Fallback to previous day's balance if 0 or empty file on holiday
                    if (balance === 0 && fileType !== 'history') {
                        if (state.files[fileType] && state.files[fileType].previousBalance) {
                            balance = state.files[fileType].previousBalance;
                        }
                    } else if (fileType !== 'history') {
                        if (state.files[fileType]) {
                            state.files[fileType].previousBalance = balance;
                        }
                    }
                    
                    // Store in state
                    state.files[fileType].loaded = true;
                    state.files[fileType].data = rows;
                    state.files[fileType].balance = balance;
                    state.files[fileType].rawName = file.name;
                    
                    // Incrementally import transactions to ledger
                    const importedCount = importNewTransactions(rows, fileType);
                    
                    // Save to local storage
                    saveStateToLocalStorage();
                    
                    // Sort ledger by amount descending & refresh indices
                    refreshLedgerCounts();
                    
                    updateFileUI(fileType, importedCount);
                    checkAllFilesLoaded();
                    renderTable();
                } catch (e) {
                    console.error("Error in Papa.parse complete callback:", e);
                    alert("ફાઇલ પ્રોસેસ કરવામાં એરર આવી છે:\n" + e.message + "\nStack: " + e.stack);
                }
            },
            error: function(err) {
                showToast(`${file.name} વાંચવામાં ભૂલ આવી: ${err.message}`, 'error');
            }
        });
    } else {
        // Handle empty/holiday file with 0 rows: carry forward previous day's balance
        let carryForwardBal = 0;
        if (state.files[fileType] && state.files[fileType].previousBalance) {
            carryForwardBal = state.files[fileType].previousBalance;
        } else if (state.files[fileType] && state.files[fileType].balance) {
            carryForwardBal = state.files[fileType].balance;
        }
        
        state.files[fileType].loaded = true;
        state.files[fileType].data = [];
        state.files[fileType].balance = carryForwardBal;
        state.files[fileType].rawName = file.name;
        
        const importedCount = importNewTransactions([], fileType);
        saveStateToLocalStorage();
        refreshLedgerCounts();
        updateFileUI(fileType, 0);
        checkAllFilesLoaded();
        renderTable();
    }
}

// Case insensitive and fuzzy column helper
function getRowValue(row, possibleHeaders) {
    if (!row) return '';
    const keys = Object.keys(row);
    
    // First pass: exact clean match
    for (let header of possibleHeaders) {
        const cleanHeader = header.toLowerCase().replace(/[^a-z0-9]/g, '');
        const match = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === cleanHeader);
        if (match) return row[match];
    }
    
    // Second pass: partial matching (fuzzy fallback)
    for (let header of possibleHeaders) {
        const cleanHeader = header.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (cleanHeader.length <= 3) continue; // Skip short words to avoid collision
        const match = keys.find(k => {
            const cleanKey = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            return cleanKey.includes(cleanHeader) || cleanHeader.includes(cleanKey);
        });
        if (match) return row[match];
    }
    return '';
}

// Robust helper to extract Reference Number from any bank row
function extractRefNoFromRow(row) {
    if (!row) return '';
    const keys = Object.keys(row);
    
    // Look for keys that contain 'chq', 'ref', 'cheque', or 'reference' case-insensitively
    // excluding keys that are obviously descriptions, dates, or amounts
    const bestKey = keys.find(k => {
        const cleanK = k.toLowerCase();
        if (cleanK.includes('desc') || cleanK.includes('narr') || cleanK.includes('particular') || cleanK.includes('detail') || cleanK.includes('date') || cleanK.includes('amt') || cleanK.includes('credit') || cleanK.includes('debit') || cleanK.includes('bal')) {
            return false;
        }
        return cleanK.includes('ref') || cleanK.includes('chq') || cleanK.includes('cheque');
    });
    
    if (bestKey) {
        return String(row[bestKey] || '').trim();
    }
    
    // Fallback to the traditional getRowValue
    return String(getRowValue(row, ['ChqRefNo', 'ChequeRefNo', 'ReferenceNo', 'RefNo', 'ChqNo', 'ChequeNo', 'ChqRef', 'ChequeRef', 'Ref', 'Cheque', 'Reference', 'Chq./Ref.No.']) || '').trim();
}

// Update Upload Cards UI
function updateFileUI(fileType, importedCount = 0) {
    const card = document.getElementById(`status-${fileType.toLowerCase()}`);
    if (!card) return;
    
    const indicator = card.querySelector('.file-indicator');
    const meta = card.querySelector('.file-meta');
    
    card.classList.add('loaded');
    indicator.className = 'file-indicator success';
    
    if (fileType === 'history') {
        meta.textContent = `સક્રિય | +${importedCount} નવા ઇતિહાસ | ${state.files[fileType].rawName}`;
    } else {
        const totalRows = state.files[fileType].data ? state.files[fileType].data.length : 0;
        if (totalRows === 0) {
            meta.textContent = `સક્રિય | ૦ રોઝ (રજાનો દિવસ) | ${state.files[fileType].rawName}`;
        } else {
            meta.textContent = `સક્રિય | ${totalRows} રોઝ (+${importedCount} નવા) | ${state.files[fileType].rawName}`;
        }
    }
    
    // Show clear button
    const trashBtn = card.querySelector('.clear-file-btn');
    if (trashBtn) trashBtn.style.display = 'flex';
    
    // Update balance card for main files only
    if (fileType !== 'history') {
        const balVal = document.getElementById(`val-${fileType.toLowerCase()}-bal`);
        if (balVal) {
            const balance = state.files[fileType].balance;
            balVal.textContent = formatCurrency(balance);
            if (balance < 0) {
                balVal.classList.add('negative-balance');
            } else {
                balVal.classList.remove('negative-balance');
            }
        }
    }
    
    if (state.files[fileType].data && state.files[fileType].data.length === 0 && fileType !== 'history') {
        showToast(`${fileType}.CSV માં કોઈ વ્યવહારો નથી (રજાનો દિવસ). બેલેન્સ શૂન્ય (₹ 0.00) સેવ થયું.`, 'info');
    } else {
        showToast(`${fileType === 'history' ? 'પેન્ડીંગ ઇતિહાસ' : fileType + '.CSV'} માંથી ${importedCount} નવા વ્યવહારો લેજરમાં ઉમેરાયા.`, 'success');
    }
}

// Check if any transactions are present to enable buttons
function checkAllFilesLoaded() {
    if (state.mergedData.length > 0) {
        btnAutoReconcile.disabled = false;
        btnExportPending.disabled = false;
        btnExportReconciled.disabled = false;
    }
}

// Merge and Process Uploaded Files
function mergeFiles() {
    const todayObj = parseDateString(state.currentDate);
    const merged = [];
    
    // 1. Process HDFC.CSV
    if (state.files.HDFC.loaded) {
        state.files.HDFC.data.forEach(row => {
            const dateStr = String(getRowValue(row, ['TransactionDate', 'Date', 'TxnDate', 'ValueDate']) || '').trim();
            const desc = String(getRowValue(row, ['Description', 'Narration', 'Particulars', 'Particular', 'Details']) || '').trim();
            const cdFlag = String(getRowValue(row, ['CDFalg', 'CDFlag', 'DrCr', 'C.D.Falg', 'C.D.Flag']) || '').trim().toUpperCase();
            
            // Extract Reference No
            const refNo = extractRefNoFromRow(row);
            
            // Check if C.D.Falg column is present (Layout A indicator)
            const hasCDFlagColumn = getRowValue(row, ['CDFalg', 'CDFlag', 'DrCr', 'C.D.Falg', 'C.D.Flag']) !== '';
            
            if (!dateStr || !desc) return; // Skip invalid/footer rows
            
            let credit = 0;
            let debit = 0;
            
            if (hasCDFlagColumn) {
                // Layout A: Single Amount column and Dr/Cr flag
                const amtStr = String(getRowValue(row, ['Amount', 'TransactionAmount', 'Value', 'રકમ']) || '').replace(/,/g, '').trim();
                const amount = parseFloat(amtStr) || 0;
                
                if (cdFlag.startsWith('C')) {
                    credit = amount;
                } else if (cdFlag.startsWith('D')) {
                    debit = amount;
                }
            } else {
                // Layout B: Separate Debit and Credit columns
                const depositVal = String(getRowValue(row, ['DepositAmt', 'DepositAmount', 'Credit', 'Deposit', 'DepositAmtINR', 'CreditAmt']) || '').replace(/,/g, '').trim();
                const withdrawalVal = String(getRowValue(row, ['WithdrawalAmt', 'WithdrawalAmount', 'Debit', 'Withdrawal', 'WithdrawalAmtINR', 'DebitAmt']) || '').replace(/,/g, '').trim();
                
                credit = parseFloat(depositVal) || 0;
                debit = parseFloat(withdrawalVal) || 0;
            }
            
            const dateObj = parseDateString(dateStr);
            let actualDate = formatDateOnly(dateObj);
            const autoFlag = getAutoFlag(desc);
            if (autoFlag === 'NPCI') {
                actualDate = extractDateFromDescription(desc, actualDate);
            }
            const daysDiff = dateObj ? calculateDaysDifference(todayObj, dateObj) : '';
            
            merged.push({
                id: 'hdfc_' + Math.random().toString(36).substr(2, 9),
                date: dateStr,
                description: desc,
                type: 'HDFC',
                actualDate: actualDate,
                creditTrn: credit,
                debitTrn: debit,
                refNo: refNo,
                flag: autoFlag,
                day: daysDiff,
                count: 0,
                reconciled: false,
                matchGroupId: null
            });
        });
    }
    
    // 2. Process 345051.CSV, 3493.CSV, 3496.CSV
    const otherTypes = ['345051', '3493', '3496'];
    otherTypes.forEach(type => {
        if (state.files[type].loaded) {
            state.files[type].data.forEach(row => {
                const dateStr = String(getRowValue(row, ['POSTDATE', 'Date']) || '').trim();
                const desc = String(getRowValue(row, ['USERNARRATION', 'Description']) || '').trim();
                const creditStr = String(getRowValue(row, ['CreditAmount', 'Credit']) || '').replace(/,/g, '').trim();
                const debitStr = String(getRowValue(row, ['Debitamount', 'Debit']) || '').replace(/,/g, '').trim();
                
                // If it's a summary row or empty row, skip
                if (!dateStr && !desc) return;
                // Skip rows that are purely balance rows (no amounts but have balance)
                const credit = parseFloat(creditStr) || 0;
                const debit = parseFloat(debitStr) || 0;
                if (credit === 0 && debit === 0 && desc && desc.toLowerCase().includes('balance')) return;
                
                const dateObj = parseDateString(dateStr);
                let actualDate = formatDateOnly(dateObj);
                const autoFlag = getAutoFlag(desc);
                if (autoFlag === 'NPCI') {
                    actualDate = extractDateFromDescription(desc, actualDate);
                }
                const daysDiff = dateObj ? calculateDaysDifference(todayObj, dateObj) : '';
                
                merged.push({
                    id: type + '_' + Math.random().toString(36).substr(2, 9),
                    date: dateStr,
                    description: desc,
                    type: type,
                    actualDate: actualDate,
                    creditTrn: credit,
                    debitTrn: debit,
                    refNo: '',
                    flag: autoFlag,
                    day: daysDiff,
                    count: 0,
                    reconciled: false,
                    matchGroupId: null
                });
            });
        }
    });
    
    // 3. Process PENDING_HISTORY.CSV (Optional)
    if (state.files.history && state.files.history.loaded) {
        state.files.history.data.forEach(row => {
            const dateStr = String(getRowValue(row, ['DATE', 'Date']) || '').trim();
            const desc = String(getRowValue(row, ['DISCRIPTION', 'Description', 'Narration']) || '').trim();
            const type = String(getRowValue(row, ['TYPE', 'Type']) || '').trim().toUpperCase();
            const rawActualDate = String(getRowValue(row, ['ACTUAL DATE', 'ActualDate']) || '').trim();
            const creditStr = String(getRowValue(row, ['CREDIT TRN', 'Credit']) || '').replace(/,/g, '').trim();
            const debitStr = String(getRowValue(row, ['DEBIT TRN', 'Debit']) || '').replace(/,/g, '').trim();
            
            if (!dateStr || !desc) return; // Skip empty rows
            
            const credit = parseFloat(creditStr) || 0;
            const debit = parseFloat(debitStr) || 0;
            
            const flagStr = String(getRowValue(row, ['FLAG', 'Flag']) || '').trim();
            const autoFlag = flagStr || getAutoFlag(desc);
            
            let actualDate = rawActualDate || dateStr;
            if (autoFlag === 'NPCI') {
                actualDate = extractDateFromDescription(desc, actualDate);
            }
            
            merged.push({
                id: 'history_' + Math.random().toString(36).substr(2, 9),
                date: dateStr,
                description: desc,
                type: type || 'HISTORICAL',
                actualDate: actualDate || dateStr,
                creditTrn: credit,
                debitTrn: debit,
                refNo: '', // No separate refNo column
                flag: flagStr || getAutoFlag(desc),
                day: '',
                count: 0,
                reconciled: false,
                matchGroupId: null
            });
        });
    }
    
    // Sort combined data by the largest transaction amount
    merged.sort((a, b) => {
        const amtA = Math.max(a.creditTrn, a.debitTrn);
        const amtB = Math.max(b.creditTrn, b.debitTrn);
        return amtB - amtA;
    });
    
    // Assign serial counts
    merged.forEach((item, index) => {
        item.count = index + 1;
    });
    
    state.mergedData = merged;
    
    // Update net total balance
    let totalBal = 0;
    Object.values(state.files).forEach(f => {
        if (f.loaded) totalBal += f.balance;
    });
    
    valTotalBal.textContent = formatCurrency(totalBal);
    if (totalBal < 0) {
        valTotalBal.classList.add('negative-balance');
    } else {
        valTotalBal.classList.remove('negative-balance');
    }
    
    // Enable exports
    btnExportPending.disabled = false;
    btnExportReconciled.disabled = false;
    
    state.currentPage = 1;
    renderTable();
}// Text Similarity Token Jaccard calculation with keyword boost
function getDescriptionSimilarity(s1, s2) {
    if (!s1 || !s2) return 0;
    
    const lowerS1 = s1.toLowerCase();
    const lowerS2 = s2.toLowerCase();
    
    // Extract dates in DD/MM/YYYY or DD-MM-YYYY or DD/MM/YY format
    const extractDates = (str) => {
        const dates = [];
        const regex = /\b(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{2,4})\b/g;
        let match;
        while ((match = regex.exec(str)) !== null) {
            let d = match[1].padStart(2, '0');
            let m = match[2].padStart(2, '0');
            let y = match[3];
            if (y.length === 2) {
                y = '20' + y;
            }
            dates.push(`${d}/${m}/${y}`);
        }
        return dates;
    };
    
    const dates1 = extractDates(s1);
    const dates2 = extractDates(s2);
    
    // Conflicting dates check: if both have dates in narration but they are different, similarity is 0
    if (dates1.length > 0 && dates2.length > 0) {
        const hasCommonDate = dates1.some(d => dates2.includes(d));
        if (!hasCommonDate) {
            return 0;
        }
    }
    
    // Check for SUBMEMBER ACH and ACH keywords matching (using boundary check to avoid KACHOT, RANACHHO)
    const achRegex = /(?:^|[^a-z])ach(?:$|[^a-z])/i;
    const submemberAchRegex = /(?:^|[^a-z])submember\s+ach(?:$|[^a-z])/i;
    
    const hasSubmemberACH = submemberAchRegex.test(lowerS1) || submemberAchRegex.test(lowerS2);
    const hasACH = achRegex.test(lowerS1) && achRegex.test(lowerS2);
    if (hasSubmemberACH && hasACH) {
        return 0.90; // Boost similarity to auto-reconcile
    }
    
    const hasCommonDateInDesc = dates1.some(d => dates2.includes(d));
    if (hasCommonDateInDesc) {
        return 0.90; // Boost similarity if they share a specific date in narration
    }
    
    // Normalize strings: lowercase and replace special characters with spaces
    const norm1 = s1.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
    const norm2 = s2.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
    
    // Space-insensitive matching: e.g. RAHUL JASHVANTRAI POPAT matching RAHULJASHVANTRAIPOPAT
    const noSpace1 = norm1.replace(/\s+/g, '');
    const noSpace2 = norm2.replace(/\s+/g, '');
    
    if (noSpace1 === noSpace2) return 1.0;
    if (noSpace1.length >= 6 && noSpace2.length >= 6) {
        if (noSpace1.includes(noSpace2) || noSpace2.includes(noSpace1)) {
            return 0.85;
        }
    }
    
    if (norm1 === norm2) return 1.0;
    if (norm1.includes(norm2) || norm2.includes(norm1)) return 0.85;
    
    // Tokenization
    const tokens1 = norm1.split(/\s+/).filter(t => t.length > 1);
    const tokens2 = norm2.split(/\s+/).filter(t => t.length > 1);
    
    if (tokens1.length === 0 || tokens2.length === 0) return 0;
    
    // Stop words to ignore for keyword boost
    const stopWords = new Set(['from', 'to', 'for', 'the', 'and', 'in', 'on', 'at', 'by', 'of', 'with', 'payment', 'transfer', 'upi', 'neft', 'rtgs', 'txn', 'ref', 'chq', 'bill', 'deposit', 'online', 'cts', 'ach', 'dr', 'cr', 'imps', 'chg', 'charge', 'charges', 'gst', 'tax', 'cash', 'self', 'salary', 'interest', 'commission', 'loan', 'emi', 'rev', 'rtn', 'return', 'reversal']);
    
    const set1 = new Set(tokens1);
    const set2 = new Set(tokens2);
    
    let intersection = 0;
    let matchingNonStopWords = 0;
    let hasLongMatchingWord = false;
    let hasMediumMatchingWord = false;
    
    set1.forEach(t => {
        if (set2.has(t)) {
            intersection++;
            // Count matching significant non-stop words (length >= 3)
            if (t.length >= 3 && !stopWords.has(t) && !/^\d{4}$/.test(t)) {
                matchingNonStopWords++;
                if (t.length >= 6) {
                    hasLongMatchingWord = true;
                }
                if (t.length >= 4) {
                    hasMediumMatchingWord = true;
                }
            }
        }
    });
    
    // Rule 1: If they share at least 2 common significant words, auto reconcile
    if (matchingNonStopWords >= 2) {
        return 0.90; // High score to auto reconcile
    }
    
    // Rule 2: If they share at least 1 very long significant word (length >= 6, e.g. "kanjibhai"), auto reconcile
    if (hasLongMatchingWord) {
        return 0.85; // Boost score to auto reconcile
    }
    
    // Rule 3: If they share at least 1 medium significant word (length >= 4, e.g. "jalu", "popat"), auto reconcile
    if (hasMediumMatchingWord) {
        return 0.85; // Boost score to auto reconcile
    }
    
    const jaccard = intersection / (set1.size + set2.size - intersection);
    return jaccard;
}
// Helper to get normalized file/account category (e.g. HDFC, 345051, 3493, 3496)
function getFileCategory(item) {
    const rawType = String(item.parentType || item.type || '').toUpperCase();
    if (rawType.includes('HDFC')) return 'HDFC';
    if (rawType.includes('345051')) return '345051';
    if (rawType.includes('3493')) return '3493';
    if (rawType.includes('3496')) return '3496';
    return rawType.replace('PEND-', '');
}

/**
 * Reconciliation pair rules:
 *   ✓  HDFC  ↔  345051
 *   ✓  HDFC  ↔  3493
 *   ✓  HDFC  ↔  3496
 *   ✗  Any same-file pair
 *   ✗  GL ↔ GL pairs (3493↔3496, 345051↔3493, etc.)
 *   → One side of every pair MUST always be HDFC.
 */
function isValidReconPair(a, b) {
    const catA = getFileCategory(a);
    const catB = getFileCategory(b);
    const oneIsHDFC = catA === 'HDFC' || catB === 'HDFC';
    const differentFiles = catA !== catB;
    return oneIsHDFC && differentFiles;
}

// Helper to extract or inherit Ref No for a transaction
function getDisplayRefNo(item) {
    if (!item) return '';
    // 1. If item has its own valid refNo, return it
    if (item.refNo && item.refNo.trim().length >= 4) {
        return item.refNo.trim();
    }
    
    // 2. If reconciled, check if partner has a valid refNo
    if (item.matchGroupId) {
        const partner = state.mergedData.find(t => t.matchGroupId === item.matchGroupId && t.id !== item.id);
        if (partner && partner.refNo && partner.refNo.trim().length >= 4) {
            return partner.refNo.trim();
        }
    }
    
    // 3. Extract from description using UTR / UPI Ref patterns
    const desc = item.description || '';
    
    // Look for alphanumeric strings that resemble UTRs or UPI ref numbers
    // Split by spaces, hyphens, colons, commas, slashes, parentheses, brackets
    const tokens = desc.split(/[\s\-:,\/\(\)\[\]\{\}]/);
    for (let token of tokens) {
        const cleanToken = token.trim();
        // Regex for RTGS/NEFT UTR (typically 12 to 22 alphanumeric characters, starts with 4 letters)
        if (/^[A-Z]{4}[A-Z0-9]{8,18}$/i.test(cleanToken)) {
            return cleanToken;
        }
        // UPI Ref patterns (typically 12 digits starting with 5 or 6 or 4 etc.)
        if (/^\d{12}$/.test(cleanToken)) {
            return cleanToken;
        }
    }
    
    // 4. Fallback lookup: Find any matching transaction in the other ledger (by amount and close date) to borrow its Ref No
    const amt = Math.max(item.creditTrn, item.debitTrn);
    if (amt > 0) {
        const isGL = getFileCategory(item) !== 'HDFC';
        const dateObj = parseDateString(item.date);
        
        const candidate = state.mergedData.find(t => {
            const isCandGL = getFileCategory(t) !== 'HDFC';
            if (isGL === isCandGL) return false; // Must be opposite category (GL vs HDFC)
            
            const candAmt = Math.max(t.creditTrn, t.debitTrn);
            if (Math.abs(amt - candAmt) > 0.05) return false;
            
            const candDate = parseDateString(t.date);
            if (dateObj && candDate && Math.abs(dateObj - candDate) <= 3 * 24 * 60 * 60 * 1000) {
                if (t.refNo && t.refNo.trim().length >= 4) {
                    return true;
                }
            }
            return false;
        });
        
        if (candidate) {
            return candidate.refNo.trim();
        }
    }
    
    return '';
}

// Auto Reconciliation Algorithm
function runAutoReconciliation() {
    if (state.mergedData.length === 0) return;
    pushToUndoStack();
    
    // Deterministic sort of mergedData to ensure 100% consistent matching results
    state.mergedData.sort((a, b) => {
        const dateA = parseDateString(a.date) || 0;
        const dateB = parseDateString(b.date) || 0;
        if (dateA - dateB !== 0) return dateA - dateB;
        
        const amtA = Math.max(a.creditTrn, a.debitTrn);
        const amtB = Math.max(b.creditTrn, b.debitTrn);
        if (Math.abs(amtA - amtB) > 0.001) return amtB - amtA;
        
        const descA = a.description || '';
        const descB = b.description || '';
        if (descA !== descB) return descA.localeCompare(descB);
        
        return a.id.localeCompare(b.id);
    });
    
    console.log("[Recon Debug] Scanning mergedData for SHREENATH:");
    state.mergedData.forEach(item => {
        if (item.description.toUpperCase().includes('SHREENATH') || item.description.toUpperCase().includes('SHREENATHJI')) {
            console.log(`[Recon Debug] Found: id=${item.id}, desc="${item.description}", type=${item.type}, parentType=${item.parentType || 'none'}, credit=${item.creditTrn}, debit=${item.debitTrn}, date=${item.date}, reconciled=${item.reconciled}`);
        }
    });
    
    let matchCount = 0;
    
    // Reset previous auto-matches to start fresh (but PRESERVE manual matches)
    state.mergedData.forEach(item => {
        if (item.matchGroupId && !item.matchGroupId.startsWith('manual_group_')) {
            item.reconciled = false;
            item.matchGroupId = null;
        }
    });
    
    // Pass 0: NPCI Daily Sum Auto-Reconciliation (Only match if total Credits === total Debits)
    const npciTxns = state.mergedData.filter(item => !item.reconciled && item.flag && item.flag.toUpperCase() === 'NPCI');
    
    // Group them by actualDate
    const npciByDate = {};
    npciTxns.forEach(item => {
        const d = item.actualDate; // Group by ACTUAL DATE
        if (!d) return;
        if (!npciByDate[d]) npciByDate[d] = [];
        npciByDate[d].push(item);
    });
    
    Object.keys(npciByDate).forEach(d => {
        const txns = npciByDate[d];
        
        let total_credits = 0;
        let total_debits = 0;
        txns.forEach(item => {
            total_credits += item.creditTrn || 0;
            total_debits += item.debitTrn || 0;
        });
        
        // Only reconcile if Credit and Debit sum match exactly (difference < 0.05)
        if (total_credits > 0 && Math.abs(total_credits - total_debits) < 0.05) {
            const groupId = 'npci_group_daily_' + d.replace(/\//g, '_') + '_' + state.matchGroupCounter;
            state.matchGroupCounter++;
            txns.forEach(item => {
                item.reconciled = true;
                item.matchGroupId = groupId;
                matchCount++;
            });
        }
    });
    
    // Pass 1: Match by HDFC Reference No in other files' descriptions (regardless of column mapping)
    const hdfcTxns = state.mergedData.filter(item => getFileCategory(item) === 'HDFC' && !item.reconciled);
    const otherTxns = state.mergedData.filter(item => getFileCategory(item) !== 'HDFC' && !item.reconciled);
    
    hdfcTxns.forEach(h => {
        const ref = h.refNo ? h.refNo.trim() : '';
        if (ref.length < 4) return; // Only match if reference number is substantial
        
        const lowerRef = ref.toLowerCase();
        const amtH = Math.max(h.creditTrn, h.debitTrn);
        
        // Find a matching transaction in the other files
        // h is always an HDFC entry (filtered above), so no need to guard it separately
        const match = otherTxns.find(o => {
            if (o.reconciled) return false;
            
            // One side must always be HDFC — pair must be valid
            if (!isValidReconPair(h, o)) return false;
            
            // Amounts must match exactly
            const amtO = Math.max(o.creditTrn, o.debitTrn);
            if (Math.abs(amtH - amtO) >= 0.01) return false;
            
            // Must be opposite sides (one Credit and one Debit)
            const isOppSide = (h.creditTrn > 0 && o.debitTrn > 0) || (h.debitTrn > 0 && o.creditTrn > 0);
            if (!isOppSide) return false;
            
            // Description must contain the reference number
            return o.description.toLowerCase().includes(lowerRef);
        });
        
        if (match) {
            const gid = 'group_' + state.matchGroupCounter;
            h.reconciled = true;
            h.matchGroupId = gid;
            h.refNoMatched = true;   // matched via Ref No — highlight badge on both sides
            
            match.reconciled = true;
            match.matchGroupId = gid;
            match.refNoMatched = true; // inherit highlight on partner
            
            state.matchGroupCounter++;
            matchCount++;
        }
    });
    
    // Pass 2: Match by description similarity and date tolerance (grouped by absolute amount)
    const pendingTxns = state.mergedData.filter(item => !item.reconciled);
    const txnsByAmt = {};
    
    pendingTxns.forEach(t => {
        const amt = Math.max(t.creditTrn, t.debitTrn).toFixed(2);
        if (parseFloat(amt) > 0) {
            if (!txnsByAmt[amt]) txnsByAmt[amt] = [];
            txnsByAmt[amt].push(t);
        }
    });
    
    Object.keys(txnsByAmt).forEach(amtStr => {
        const group = txnsByAmt[amtStr];
        if (group.length < 2) return;
        
        const matchedIds = new Set();
        
        for (let i = 0; i < group.length; i++) {
            const a = group[i];
            if (a.reconciled || matchedIds.has(a.id)) continue;
            
            const aDate = parseDateString(a.date);
            
            let bestMatch = null;
            let bestSimilarity = -1;
            
            for (let j = 0; j < group.length; j++) {
                if (i === j) continue;
                const b = group[j];
                if (b.reconciled || matchedIds.has(b.id)) continue;
                
                // Pair rule:
                //   Odd amounts  → any different-file pair is allowed (GL ↔ GL OK)
                //   Round amounts → one side must be HDFC
                const amtVal2 = parseFloat(amtStr);
                const isRound2 = (amtVal2 % 1000 === 0);
                const isOdd2 = (Math.round(amtVal2 * 100) % 100 !== 0) || (amtVal2 % 100 !== 0) || !isRound2;
                if (!isOdd2 && !isValidReconPair(a, b)) continue;
                
                // Must be opposite sides (one Credit and one Debit)
                const isOppositeSide = (a.creditTrn > 0 && b.debitTrn > 0) || (a.debitTrn > 0 && b.creditTrn > 0);
                if (!isOppositeSide) continue;
                
                // Same-file matching rules:
                //   ROUND amounts   → same file NEVER allowed
                //   ODD amounts     → same file allowed only for GL files (345051, 3493, 3496).
                //                     HDFC bank entries are always separate real transactions;
                //                     two HDFC entries must never auto-match each other.
                const isSameFile = getFileCategory(a) === getFileCategory(b);
                const catA = getFileCategory(a);
                const sameFileGlOk = isOdd2 && catA !== 'HDFC';
                if (isSameFile && !sameFileGlOk) continue;
                
                // Description Similarity
                const sim = getDescriptionSimilarity(a.description, b.description);
                if (sim < state.similarityThreshold) continue;
                
                // Date constraint: always enforce a maximum tolerance of 15 days, and enforce state.dateTolerance if similarity is below 0.95
                const bDate = parseDateString(b.date);
                if (aDate && bDate) {
                    const daysDiff = Math.abs(calculateDaysDifference(aDate, bDate));
                    if (daysDiff > 15) continue;
                    if (sim < 0.95 && daysDiff > state.dateTolerance) continue;
                }
                
                if (sim > bestSimilarity) {
                    bestSimilarity = sim;
                    bestMatch = b;
                }
            }
            
            if (bestMatch) {
                a.reconciled = true;
                a.matchGroupId = 'group_' + state.matchGroupCounter;
                
                bestMatch.reconciled = true;
                bestMatch.matchGroupId = 'group_' + state.matchGroupCounter;
                
                state.matchGroupCounter++;
                matchCount++;
                
                matchedIds.add(a.id);
                matchedIds.add(bestMatch.id);
            }
        }
    });
    
    // Pass 3: Match remaining unique odd-figure amounts (grouped by absolute amount)
    const remTxns = state.mergedData.filter(item => !item.reconciled);
    const remTxnsByAmt = {};
    
    remTxns.forEach(t => {
        const amt = Math.max(t.creditTrn, t.debitTrn).toFixed(2);
        if (parseFloat(amt) > 0) {
            if (!remTxnsByAmt[amt]) remTxnsByAmt[amt] = [];
            remTxnsByAmt[amt].push(t);
        }
    });
    
    Object.keys(remTxnsByAmt).forEach(amtStr => {
        const group = remTxnsByAmt[amtStr];
        
        // Only match if there are EXACTLY 2 remaining transactions of this amount
        if (group.length === 2) {
            const a = group[0];
            const b = group[1];
            
            if (a.reconciled || b.reconciled) return;
            
            // Basic gates: opposite sides always required.
            // isDiffFile required only for round amounts; odd amounts may match within the same file.
            const isOppositeSide = (a.creditTrn > 0 && b.debitTrn > 0) || (a.debitTrn > 0 && b.creditTrn > 0);
            const isDiffFile = getFileCategory(a) !== getFileCategory(b);
            if (!isOppositeSide) return;
            
            const amtVal = parseFloat(amtStr);
            const isRound = (amtVal % 1000 === 0);
            const isOddQuick = (Math.round(amtVal * 100) % 100 !== 0) || (amtVal % 100 !== 0) || !isRound;
            // For round amounts, different file is mandatory.
            // For odd amounts, same file is allowed ONLY for GL files (not HDFC):
            //   HDFC bank entries are always separate real transactions —
            //   two HDFC Credit/Debit entries of the same amount must never auto-pair.
            const catP3_a = getFileCategory(a);
            const sameFileP3Ok = isOddQuick && catP3_a !== 'HDFC';
            if (!isDiffFile && !sameFileP3Ok) return;
            
            // Reuse isRound & isOddQuick computed above; also check uniqueness + flag for round figures
            const totalCountOfAmt = state.mergedData.filter(item => {
                const itemAmt = Math.max(item.creditTrn, item.debitTrn).toFixed(2);
                const isNotHistory = !item.id.startsWith('history_') && item.type !== 'history' && item.parentType !== 'history';
                return itemAmt === amtStr && isNotHistory;
            }).length;
            
            let isOdd = isOddQuick;
            if (isRound && totalCountOfAmt === 2) {
                const flagA = (a.flag || 'DAILY').trim().toUpperCase();
                const flagB = (b.flag || 'DAILY').trim().toUpperCase();
                if (flagA === flagB) isOdd = true;
            }
            
            // Pair rule:
            //   Odd amounts  → any pair (same file OK, GL ↔ GL OK)
            //   Round amounts → one side must be HDFC
            if (!isOdd) {
                if (!isValidReconPair(a, b)) return;
            }
            
            if (isOdd) {
                const aDate = parseDateString(a.date);
                const bDate = parseDateString(b.date);
                let dateValid = true;
                if (aDate && bDate) {
                    const daysDiff = Math.abs(calculateDaysDifference(aDate, bDate));
                    if (daysDiff > 15) dateValid = false;
                }
                
                if (dateValid) {
                    const extractDates = (str) => {
                        const dates = [];
                        const regex = /\b(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{2,4})\b/g;
                        let match;
                        while ((match = regex.exec(str)) !== null) {
                            let d = match[1].padStart(2, '0');
                            let m = match[2].padStart(2, '0');
                            let y = match[3];
                            if (y.length === 2) y = '20' + y;
                            dates.push(`${d}/${m}/${y}`);
                        }
                        return dates;
                    };
                    const descDatesA = extractDates(a.description || '');
                    const descDatesB = extractDates(b.description || '');
                    if (descDatesA.length > 0 && descDatesB.length > 0) {
                        const hasCommonDate = descDatesA.some(d => descDatesB.includes(d));
                        if (!hasCommonDate) dateValid = false;
                    }
                }
                
                if (dateValid) {
                    a.reconciled = true;
                    a.matchGroupId = 'group_' + state.matchGroupCounter;
                    
                    b.reconciled = true;
                    b.matchGroupId = 'group_' + state.matchGroupCounter;
                    
                    state.matchGroupCounter++;
                    matchCount++;
                }
            }
        }
    });
    
    if (matchCount > 0) {
        showToast(`ઓટો-રીકન્સિલેશન પૂર્ણ! ${matchCount} જોડકાં (પેર) ઓટો-મેચ થયા.`, 'success');
        saveStateToLocalStorage();
        refreshLedgerCounts();
    } else {
        showToast('કોઈ ઓટો-મેચ મળ્યા નથી. કૃપા કરીને સેટિંગ્સમાં સમાનતા અથવા તારીખ તફાવત ઘટાડો.', 'warning');
    }
    
    state.currentPage = 1;
    renderTable();
    
    // Check if the live difference is now zero (0.00) after auto-reconciliation
    const finalDiff = calculateCurrentDifference();
    if (Math.abs(finalDiff) < 0.01 && state.mergedData.length > 0) {
        showToast('મેળવણી પત્રક પૂર્ણ! ડિફરન્સ શૂન્ય છે.', 'success');
    }
}

// Toggle selection checkbox for manual reconciliation
function toggleSelect(itemId) {
    if (state.selectedIds.has(itemId)) {
        state.selectedIds.delete(itemId);
    } else {
        state.selectedIds.add(itemId);
    }
    
    const checkbox = document.querySelector(`.recon-checkbox[data-id="${itemId}"]`);
    if (checkbox) {
        if (state.selectedIds.has(itemId)) {
            checkbox.classList.add('checked');
        } else {
            checkbox.classList.remove('checked');
        }
    }
    
    // If on NPCI tab, update date checkbox if necessary
    if (state.currentTab === 'npci-pivot') {
        const item = state.mergedData.find(t => t.id === itemId);
        if (item) {
            const dateStr = item.actualDate;
            if (dateStr) {
                if (!state.selectedNpciDates) {
                    state.selectedNpciDates = new Set();
                }
                const txns = state.mergedData.filter(t => t.actualDate === dateStr && t.flag && t.flag.toUpperCase() === 'NPCI');
                const allSelected = txns.length > 0 && txns.every(t => state.selectedIds.has(t.id));
                const dateCheckbox = document.querySelector(`.npci-checkbox-${dateStr.replace(/\//g, '_')}`);
                if (dateCheckbox) {
                    if (allSelected) {
                        dateCheckbox.classList.add('checked');
                        state.selectedNpciDates.add(dateStr);
                    } else {
                        dateCheckbox.classList.remove('checked');
                        state.selectedNpciDates.delete(dateStr);
                    }
                }
            }
        }
    }
    
    updateBulkActionButtons();
    updateBRSLiveWidget();
}

// Update manually edited balance from closing balance card headers
function updateManualBalance(fileType, text) {
    // Parse the number from the text (supporting minus sign, digits and decimal point)
    let cleanText = text.replace(/[^\d.-]/g, '');
    // If multiple decimals or minus signs, clean up
    if ((cleanText.match(/\./g) || []).length > 1) {
        const parts = cleanText.split('.');
        cleanText = parts[0] + '.' + parts.slice(1).join('');
    }
    const newBal = parseFloat(cleanText) || 0;
    
    if (state.files && state.files[fileType]) {
        pushToUndoStack();
        state.files[fileType].balance = newBal;
        saveStateToLocalStorage();
        updateBRSLiveWidget();
        refreshLedgerCounts();
        showToast(`${fileType} બેલેન્સ સફળતાપૂર્વક અપડેટ થયું.`, 'success');
    }
}

// Clear all checkbox selections
function clearAllSelections() {
    state.selectedIds.clear();
    if (state.selectedNpciDates) {
        state.selectedNpciDates.clear();
    }
    updateBulkActionButtons();
    updateBRSLiveWidget();
    renderTable();
    showToast("તમામ પસંદગી સાફ કરવામાં આવી છે.", "info");
}

// Delete a single transaction from the ledger
function deleteTransaction(itemId) {
    if (confirm("શું તમે આ વ્યવહારને કાયમ માટે ડીલીટ કરવા માંગો છો?")) {
        pushToUndoStack();
        state.mergedData = state.mergedData.filter(item => item.id !== itemId);
        state.selectedIds.delete(itemId);
        saveStateToLocalStorage();
        refreshLedgerCounts();
        renderTable();
        showToast('વ્યવહાર સફળતાપૂર્વક ડીલીટ કરવામાં આવ્યો છે.', 'info');
    }
}

// Reconcile a single pending transaction (manual — no file-type restriction)
function reconcileSingle(itemId) {
    pushToUndoStack();
    const item = state.mergedData.find(t => t.id === itemId);
    if (item && !item.reconciled) {
        item.reconciled = true;
        item.matchGroupId = 'manual_group_' + state.matchGroupCounter;
        state.matchGroupCounter++;
        
        saveStateToLocalStorage();
        refreshLedgerCounts();
        renderTable();
        showToast('વ્યવહાર સફળતાપૂર્વક મેળવવામાં (Reconciled) આવ્યો છે.', 'success');
        
        // Check if the live difference is now zero (0.00) after reconciliation
        const finalDiff = calculateCurrentDifference();
        if (Math.abs(finalDiff) < 0.01 && state.mergedData.length > 0) {
            showToast('મેળવણી પત્રક પૂર્ણ! ડિફરન્સ શૂન્ય છે.', 'success');
        }
    }
}

// Unreconcile a single reconciled transaction
function unreconcileSingle(itemId) {
    pushToUndoStack();
    const item = state.mergedData.find(t => t.id === itemId);
    if (item && item.reconciled) {
        if (item.matchGroupId) {
            // Find partner and unreconcile both
            const partner = state.mergedData.find(t => t.matchGroupId === item.matchGroupId && t.id !== item.id);
            item.reconciled = false;
            item.matchGroupId = null;
            if (partner) {
                partner.reconciled = false;
                partner.matchGroupId = null;
            }
        } else {
            item.reconciled = false;
        }
        
        saveStateToLocalStorage();
        refreshLedgerCounts();
        renderTable();
        showToast('વ્યવહાર ફરીથી પેન્ડીંગ કરવામાં આવ્યો છે.', 'info');
    }
}

// Bulk delete all selected transactions
function runBulkDelete() {
    if (state.selectedIds.size === 0) return;
    
    if (confirm(`શું તમે પસંદ કરેલા તમામ ${state.selectedIds.size} વ્યવહારોને કાયમ માટે ડીલીટ કરવા માંગો છો?`)) {
        pushToUndoStack();
        const idsToDelete = new Set(state.selectedIds);
        state.mergedData = state.mergedData.filter(item => !idsToDelete.has(item.id));
        state.selectedIds.clear();
        updateBulkActionButtons();
        saveStateToLocalStorage();
        refreshLedgerCounts();
        renderTable();
        showToast('પસંદ કરેલા વ્યવહારો સફળતાપૂર્વક ડીલીટ કરવામાં આવ્યા છે.', 'info');
    }
}

// Update visibility and text of bulk action button
function updateBulkActionButtons() {
    const bulkDeleteBtn = document.getElementById('btn-bulk-delete');
    const npciPendingBtn = document.getElementById('btn-npci-pending');
    
    if (state.currentTab === 'npci-pivot') {
        if (bulkDeleteBtn) bulkDeleteBtn.style.display = 'none';
        
        let pendingCount = 0;
        let reconciledCount = 0;
        
        state.selectedIds.forEach(id => {
            const item = state.mergedData.find(t => t.id === id);
            if (item) {
                if (item.reconciled) reconciledCount++;
                else pendingCount++;
            }
        });
        
        if (pendingCount > 0 || reconciledCount > 0) {
            bulkActionsContainer.style.display = 'flex';
            
            // Reconcile button
            if (pendingCount > 0) {
                btnBulkAction.style.display = 'inline-flex';
                btnBulkAction.className = 'btn btn-success';
                btnBulkAction.style.backgroundColor = '';
                bulkActionText.textContent = `પસંદ કરેલ NPCI રીકન્સાઇલ કરો (${pendingCount})`;
            } else {
                btnBulkAction.style.display = 'none';
            }
            
            // Pending button
            if (npciPendingBtn) {
                if (reconciledCount > 0) {
                    npciPendingBtn.style.display = 'inline-flex';
                    const npciPendingText = document.getElementById('npci-pending-text');
                    if (npciPendingText) {
                        npciPendingText.textContent = `પસંદ કરેલ NPCI પેન્ડિંગ કરો (${reconciledCount})`;
                    }
                } else {
                    npciPendingBtn.style.display = 'none';
                }
            }
        } else {
            bulkActionsContainer.style.display = 'none';
            if (npciPendingBtn) npciPendingBtn.style.display = 'none';
        }
    } else {
        if (npciPendingBtn) npciPendingBtn.style.display = 'none';
        if (bulkDeleteBtn) bulkDeleteBtn.style.display = 'inline-flex';
        
        if (state.selectedIds.size > 0) {
            bulkActionsContainer.style.display = 'flex';
            btnBulkAction.style.display = 'inline-flex';
            if (state.currentTab === 'pending') {
                btnBulkAction.className = 'btn btn-success';
                btnBulkAction.style.backgroundColor = '';
                bulkActionText.textContent = `પસંદ કરેલ રીકન્સાઇલ કરો (${state.selectedIds.size})`;
            } else {
                btnBulkAction.className = 'btn';
                btnBulkAction.style.backgroundColor = '#f59e0b';
                btnBulkAction.style.color = '#fff';
                bulkActionText.textContent = `પસંદ કરેલ પેન્ડીંગ કરો (${state.selectedIds.size})`;
            }
        } else {
            bulkActionsContainer.style.display = 'none';
        }
    }
}

// Run Bulk Actions (move checked entries)
function runBulkAction() {
    pushToUndoStack();
    if (state.currentTab === 'npci-pivot') {
        let count = 0;
        const pendingSelected = [];
        state.selectedIds.forEach(id => {
            const item = state.mergedData.find(t => t.id === id);
            if (item && !item.reconciled) {
                pendingSelected.push(item);
            }
        });
        
        if (pendingSelected.length === 0) return;
        
        // Group by actualDate
        const byDate = {};
        pendingSelected.forEach(item => {
            if (!byDate[item.actualDate]) byDate[item.actualDate] = [];
            byDate[item.actualDate].push(item);
        });
        
        Object.keys(byDate).forEach(dateStr => {
            const txns = byDate[dateStr];
            const groupId = 'npci_group_' + dateStr.replace(/\//g, '_') + '_' + state.matchGroupCounter;
            state.matchGroupCounter++;
            txns.forEach(item => {
                item.reconciled = true;
                item.matchGroupId = groupId;
                count++;
            });
        });
        
        // Clear selected IDs that are now reconciled
        pendingSelected.forEach(item => state.selectedIds.delete(item.id));
        if (state.selectedNpciDates) state.selectedNpciDates.clear();
        
        updateBulkActionButtons();
        saveStateToLocalStorage();
        refreshLedgerCounts();
        renderTable();
        showToast(`${count} NPCI વ્યવહારો સફળતાપૂર્વક રીકન્સાઇલ કર્યા!`, 'success');
        
        // Check if the live difference is now zero (0.00) after bulk manual reconciliation
        const finalDiff = calculateCurrentDifference();
        if (Math.abs(finalDiff) < 0.01 && state.mergedData.length > 0) {
            showToast('મેળવણી પત્રક પૂર્ણ! ડિફરન્સ શૂન્ય છે.', 'success');
        }
        return;
    }
    
    if (state.selectedIds.size === 0) return;
    
    let count = 0;
    
    if (state.currentTab === 'pending') {
        // Manual reconcile — no file-type restriction; user can reconcile any combination.
        // (Auto-reconcile enforces the HDFC rule; manual is the user's explicit override.)
        const manualGroupId = 'manual_group_' + state.matchGroupCounter;
        state.matchGroupCounter++;
        state.selectedIds.forEach(id => {
            const item = state.mergedData.find(t => t.id === id);
            if (item && !item.reconciled) {
                item.reconciled = true;
                item.matchGroupId = manualGroupId;
                count++;
            }
        });
        showToast(`${count} વ્યવહારો સફળતાપૂર્વક રીકન્સાઇલ કર્યા!`, 'success');
    } else {
        state.selectedIds.forEach(id => {
            const item = state.mergedData.find(t => t.id === id);
            if (item && item.reconciled) {
                if (item.matchGroupId) {
                    const partner = state.mergedData.find(t => t.matchGroupId === item.matchGroupId && t.id !== item.id);
                    item.reconciled = false;
                    item.matchGroupId = null;
                    if (partner) {
                        partner.reconciled = false;
                        partner.matchGroupId = null;
                    }
                } else {
                    item.reconciled = false;
                }
                count++;
            }
        });
        showToast(`${count} વ્યવહારો પેન્ડીંગ લિસ્ટમાં પાછા મોકલ્યા!`, 'info');
    }
    
    state.selectedIds.clear();
    updateBulkActionButtons();
    saveStateToLocalStorage(); // Save manual bulk changes
    refreshLedgerCounts();
    renderTable();
    
    // Check if the live difference is now zero (0.00) after bulk manual reconciliation
    const finalDiff = calculateCurrentDifference();
    if (Math.abs(finalDiff) < 0.01 && state.mergedData.length > 0) {
        showToast('મેળવણી પત્રક પૂર્ણ! ડિફરન્સ શૂન્ય છે.', 'success');
    }
}

// Bulk mark NPCI dates as pending (un-reconciled)
function runNpciPendingAction() {
    let count = 0;
    const reconciledSelected = [];
    state.selectedIds.forEach(id => {
        const item = state.mergedData.find(t => t.id === id);
        if (item && item.reconciled) {
            reconciledSelected.push(item);
        }
    });
    
    if (reconciledSelected.length === 0) return;
    
    reconciledSelected.forEach(item => {
        item.reconciled = false;
        item.matchGroupId = null;
        count++;
    });
    
    // Clear selected IDs that are now pending
    reconciledSelected.forEach(item => state.selectedIds.delete(item.id));
    if (state.selectedNpciDates) state.selectedNpciDates.clear();
    
    updateBulkActionButtons();
    saveStateToLocalStorage();
    refreshLedgerCounts();
    renderTable();
    showToast(`${count} NPCI વ્યવહારો સફળતાપૂર્વક પેન્ડિંગ (વણમેળવાયેલ) તરીકે સેટ કર્યા!`, 'info');
}

// Switch tabs: pending vs reconciled vs NPCI pivot vs rbi
function switchTab(tabName) {
    state.currentTab = tabName;
    tabPending.classList.toggle('active', tabName === 'pending');
    tabReconciled.classList.toggle('active', tabName === 'reconciled');
    if (tabNPCIPivot) tabNPCIPivot.classList.toggle('active', tabName === 'npci-pivot');
    if (tabRbi) tabRbi.classList.toggle('active', tabName === 'rbi');
    
    // Hide/show search filters
    const searchFilters = document.querySelector('.search-filters');
    if (searchFilters) {
        searchFilters.style.display = (tabName === 'npci-pivot' || tabName === 'rbi') ? 'none' : 'flex';
    }
    
    // Show/hide RBI report container vs main table container
    const mainTableContainer = document.getElementById('main-table-container');
    const tableFooter = document.querySelector('.table-footer');
    if (rbiReportContainer && mainTableContainer && tableFooter) {
        if (tabName === 'rbi') {
            rbiReportContainer.style.display = 'block';
            mainTableContainer.style.display = 'none';
            tableFooter.style.display = 'none';
        } else {
            rbiReportContainer.style.display = 'none';
            mainTableContainer.style.display = 'block';
            tableFooter.style.display = 'flex';
        }
    }
    
    state.selectedIds.clear(); // Clear selection on tab change
    if (state.selectedNpciDates) state.selectedNpciDates.clear(); // Clear NPCI selection on tab change
    updateBulkActionButtons();
    state.currentPage = 1;
    renderTable();
}

// Filter and Search data to display
function getFilteredData() {
    return state.mergedData.filter(item => {
        // Tab filter
        const matchesTab = (state.currentTab === 'pending' && !item.reconciled) || 
                             (state.currentTab === 'reconciled' && item.reconciled);
        if (!matchesTab) return false;
        
        // Dropdown Type filter
        if (state.filterType !== 'all' && item.type !== state.filterType) return false;
        
        // Search text box filter
        if (state.searchQuery.trim() !== '') {
            const query = state.searchQuery.toLowerCase().trim();
            const cleanQuery = query.replace(/[\s,₹]/g, '');
            
            const cleanDesc = item.description.toLowerCase().replace(/\s+/g, '');
            const queryNoSpace = query.replace(/\s+/g, '');
            
            const refStr = item.refNo ? item.refNo.toLowerCase() : '';
            const typeStr = item.type ? item.type.toLowerCase() : '';
            
            // Match amounts
            const creditStr = item.creditTrn.toString();
            const debitStr = item.debitTrn.toString();
            const creditFixed = item.creditTrn.toFixed(2);
            const debitFixed = item.debitTrn.toFixed(2);
            
            const matchesAmount = (item.creditTrn > 0 && (creditStr.includes(cleanQuery) || creditFixed.includes(cleanQuery))) ||
                                  (item.debitTrn > 0 && (debitStr.includes(cleanQuery) || debitFixed.includes(cleanQuery)));
            
            const matchText = item.description.toLowerCase().includes(query) ||
                              cleanDesc.includes(queryNoSpace) ||
                              refStr.includes(query) ||
                              typeStr.includes(query) ||
                              item.date.includes(query) ||
                              item.actualDate.includes(query) ||
                              matchesAmount;
            if (!matchText) return false;
        }
        
        return true;
    });
}

// Render Data Table and Pagination
function renderTable() {
    updateBRSLiveWidget();
    // Update badge counts
    const totalPending = state.mergedData.filter(t => !t.reconciled).length;
    const totalReconciled = state.mergedData.filter(t => t.reconciled).length;
    badgePending.textContent = totalPending;
    badgeReconciled.textContent = totalReconciled;
    
    if (state.currentTab === 'npci-pivot') {
        renderNPCIPivotTable();
        return;
    }
    
    const filtered = getFilteredData();
    const totalCount = filtered.length;
    
    // Pagination math
    const totalPages = Math.ceil(totalCount / state.pageSize) || 1;
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    
    const startIndex = (state.currentPage - 1) * state.pageSize;
    const endIndex = Math.min(startIndex + state.pageSize, totalCount);
    
    const pageItems = filtered.slice(startIndex, endIndex);
    
    // Empty state check
    if (pageItems.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="12" class="table-empty-state">
                    <i data-lucide="file-warning"></i>
                    <p>${state.mergedData.length === 0 
                        ? 'કોઈ ડેટા ઉપલબ્ધ નથી. કૃપા કરીને ઉપરની ચાર સીએસવી ફાઈલો અપલોડ કરો.' 
                        : 'શોધ પરિણામોમાં કોઈ મેળ ખાતા વ્યવહારો મળ્યા નથી.'}</p>
                </td>
            </tr>
        `;
        entriesCount.textContent = `0 માંથી 0 વ્યવહારો દર્શાવેલ છે`;
        paginationControls.innerHTML = '';
        lucide.createIcons();
        return;
    }
    
    // Count frequencies of absolute amounts in the filtered list to color duplicates
    const amountCounts = {};
    filtered.forEach(item => {
        const amt = Math.max(item.creditTrn, item.debitTrn).toFixed(2);
        if (parseFloat(amt) > 0) {
            amountCounts[amt] = (amountCounts[amt] || 0) + 1;
        }
    });
    
    // Map duplicate amounts to color hex codes
    const duplicateColors = {};
    let colIdx = 0;
    const colorsPalette = [
        '#8b5cf6', // Violet
        '#3b82f6', // Blue
        '#10b981', // Emerald
        '#ec4899', // Pink
        '#f59e0b', // Amber
        '#06b6d4', // Cyan
        '#14b8a6', // Teal
        '#f43f5e', // Rose
        '#a855f7', // Purple
        '#6366f1'  // Indigo
    ];
    
    Object.keys(amountCounts).forEach(amt => {
        if (amountCounts[amt] > 1) {
            duplicateColors[amt] = colorsPalette[colIdx % colorsPalette.length];
            colIdx++;
        }
    });

    // Render Rows
    let html = '';
    
    // Group index coloring helper for matched pairs
    const groupColors = {};
    let colorIndex = 0;
    
    pageItems.forEach((item, idx) => {
        const amtVal = Math.max(item.creditTrn, item.debitTrn).toFixed(2);
        const dupColor = duplicateColors[amtVal];
        
        let firstTdStyle = '';
        let creditHtml = item.creditTrn > 0 ? formatCurrency(item.creditTrn) : '';
        let debitHtml = item.debitTrn > 0 ? formatCurrency(item.debitTrn) : '';
        
        if (dupColor) {
            // Apply left border to first td to indicate duplicate amount group
            firstTdStyle = `style="border-left: 4px solid ${dupColor}; padding-left: 12px;"`;
            
            // Highlight amounts in capsules
            if (item.creditTrn > 0) {
                creditHtml = `<span style="background-color: ${dupColor}20; border-radius: 6px; padding: 4px 8px; color: var(--success); font-weight: 700; border: 1px solid ${dupColor}40; display: inline-block;">${formatCurrency(item.creditTrn)}</span>`;
            }
            if (item.debitTrn > 0) {
                debitHtml = `<span style="background-color: ${dupColor}20; border-radius: 6px; padding: 4px 8px; color: var(--danger); font-weight: 700; border: 1px solid ${dupColor}40; display: inline-block;">${formatCurrency(item.debitTrn)}</span>`;
            }
        }
        
        // Match status badge
        let statusHtml = '';
        if (item.reconciled) {
            const displayGroupId = item.matchGroupId || 'None';
            if (item.matchGroupId && !item.matchGroupId.startsWith('manual_group_') && !item.matchGroupId.startsWith('npci_group_')) {
                statusHtml = `<span class="flag-badge auto" title="Group: ${displayGroupId}">Auto-Match (${displayGroupId})</span>`;
            } else {
                statusHtml = `<span class="flag-badge manual" title="Group: ${displayGroupId}">Manual (${displayGroupId})</span>`;
            }
        } else {
            statusHtml = `<span class="flag-badge pending">Pending</span>`;
        }
        
        // Highlight class if auto-matched and has group
        let highlightClass = '';
        if (item.matchGroupId) {
            if (groupColors[item.matchGroupId] === undefined) {
                groupColors[item.matchGroupId] = colorIndex % 6;
                colorIndex++;
            }
            highlightClass = `match-pair-highlight-${groupColors[item.matchGroupId]}`;
        }
        
        const typeClass = item.type === 'HDFC' ? 'HDFC' : 'x' + item.type;
        
        const isSelected = state.selectedIds.has(item.id);
        const descHtml = highlightDescription(item.description);
        
        // Highlight/Get REF NO if matched, inherited, or extracted
        const displayRefNo = getDisplayRefNo(item);
        let refNoHtml = displayRefNo || '-';
        if (displayRefNo && item.matchGroupId) {
            if (item.refNoMatched) {
                // Ref-No-based reconciliation: show a distinct highlighted badge on BOTH sides
                refNoHtml = `<span class="ref-no-badge-matched ref-no-badge-reflink" title="Ref No ઉપરથી મેળવણી">🔗 ${displayRefNo}</span>`;
            } else {
                refNoHtml = `<span class="ref-no-badge-matched">${displayRefNo}</span>`;
            }
        }
        
        let separatorClass = '';
        if (item.matchGroupId) {
            const nextItem = pageItems[idx + 1];
            if (!nextItem || nextItem.matchGroupId !== item.matchGroupId) {
                separatorClass = 'group-separator';
            }
        }
        
        html += `
            <tr class="${highlightClass} ${separatorClass}" title="${item.matchGroupId ? 'ઓટો-રીકન્સાઇલ્ડ જોડકું (Group: ' + item.matchGroupId + ')' : ''}">
                <td ${firstTdStyle}>${item.date}</td>
                <td style="font-weight: 500; font-size: 12px; line-height: 1.5; color: var(--text-primary); min-width: 380px; word-break: break-word;">${descHtml}</td>
                <td class="type-cell ${typeClass}">
                    <span contenteditable="true" class="editable-type-cell" style="cursor: pointer; display: inline-block; min-width: 50px;" onblur="updateTransactionType('${item.id}', this.textContent, this)">${item.type}</span>
                </td>
                <td style="text-align: center; vertical-align: middle;">
                    <span contenteditable="true" class="editable-flag-cell" onblur="updateFlagValue('${item.id}', this.textContent)">${item.flag || 'DAILY'}</span>
                </td>
                <td style="font-family: monospace; font-size: 11px; text-align: center; vertical-align: middle;">${refNoHtml}</td>
                <td style="text-align: center; vertical-align: middle;">
                    <div style="display: inline-flex; align-items: center; justify-content: center; gap: 8px;">
                        <div class="recon-checkbox ${isSelected ? 'checked' : ''}" data-id="${item.id}" onclick="toggleSelect('${item.id}')">
                            <i data-lucide="check"></i>
                        </div>
                        ${item.reconciled 
                            ? `<button onclick="event.stopPropagation(); unreconcileSingle('${item.id}')" title="પેન્ડીંગ કરો / Set Pending" style="border: none; background: transparent; color: #f59e0b; cursor: pointer; padding: 4px; display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px;">
                                 <i data-lucide="undo-2" style="width:14px; height:14px;"></i>
                               </button>`
                            : `<button onclick="event.stopPropagation(); reconcileSingle('${item.id}')" title="રીકન્સાઇલ કરો / Reconcile" style="border: none; background: transparent; color: #10b981; cursor: pointer; padding: 4px; display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px;">
                                 <i data-lucide="check-circle-2" style="width:14px; height:14px;"></i>
                               </button>`
                        }
                        <button class="delete-row-btn" onclick="event.stopPropagation(); deleteTransaction('${item.id}')" title="ડીલીટ કરો / Delete" style="border: none; background: transparent; color: var(--danger); cursor: pointer; padding: 4px; border-radius: 4px; transition: all 0.2s ease; display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px;">
                            <i data-lucide="trash-2" style="width:14px; height:14px;"></i>
                        </button>
                    </div>
                </td>
                <td>${item.actualDate}</td>
                <td class="amount-credit" style="text-align: right;">${creditHtml}</td>
                <td class="amount-debit" style="text-align: right;">${debitHtml}</td>
                <td>${statusHtml}</td>
                <td style="text-align: center; font-weight: 600;">${item.day}</td>
                <td style="text-align: center; color: var(--text-muted);">${item.count}</td>
            </tr>
        `;
    });
    
    tableBody.innerHTML = html;
    entriesCount.textContent = `${totalCount} માંથી ${startIndex + 1} થી ${endIndex} વ્યવહારો દર્શાવેલ છે`;
    
    // Render pagination controls
    renderPagination(totalPages);
    lucide.createIcons();
}

// Render NPCI Daily Settlement Pivot Table
function renderNPCIPivotTable() {
    // Hide search filters and bulk actions container when on NPCI tab
    const searchFilters = document.querySelector('.search-filters');
    if (searchFilters) searchFilters.style.display = 'none';
    if (bulkActionsContainer) bulkActionsContainer.style.display = 'none';
    
    if (!state.openNpciDateDetails) {
        state.openNpciDateDetails = new Set();
    }
    
    // Group ONLY pending (unreconciled) NPCI transactions for NPCI Summary
    const npciTxns = state.mergedData.filter(item => item.flag && item.flag.toUpperCase() === 'NPCI' && !item.reconciled);
    
    const npciByDate = {};
    npciTxns.forEach(item => {
        const d = item.actualDate;
        if (!d) return;
        if (!npciByDate[d]) npciByDate[d] = [];
        npciByDate[d].push(item);
    });
    
    const dates = Object.keys(npciByDate).sort((a, b) => {
        const dateA = parseDateString(a) || 0;
        const dateB = parseDateString(b) || 0;
        return dateB - dateA; // Newest first
    });
    
    let html = `
        <table class="recon-table" style="min-width: 100%;">
            <thead>
                <tr>
                    <th style="width: 120px; text-align: center;">ACTUAL DATE</th>
                    <th style="text-align: right;">GL CREDITS SUM (Our GL)</th>
                    <th style="text-align: right;">GL DEBITS SUM (Our GL)</th>
                    <th style="text-align: right;">HDFC CREDITS SUM (Bank)</th>
                    <th style="text-align: right;">HDFC DEBITS SUM (Bank)</th>
                    <th style="text-align: right;">NET DIFFERENCE</th>
                    <th style="width: 100px; text-align: center;">STATUS</th>
                    <th style="width: 100px; text-align: center;">ACTION (RECON)</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    if (dates.length === 0) {
        html += `
            <tr>
                <td colspan="8" class="table-empty-state">
                    <i data-lucide="file-warning"></i>
                    <p>કોઈ પેન્ડીંગ NPCI વ્યવહારો મળ્યા નથી.</p>
                </td>
            </tr>
        `;
    } else {
        dates.forEach((d, idx) => {
            const txns = npciByDate[d];
            const dateId = 'npci_date_' + idx;
            const isExpanded = state.openNpciDateDetails.has(d) || state.openNpciDateDetails.has(dateId);
            
            const hdfcSide = txns.filter(item => getFileCategory(item) === 'HDFC');
            const glSide = txns.filter(item => getFileCategory(item) !== 'HDFC');
            
            let hdfc_credits = 0;
            let hdfc_debits = 0;
            hdfcSide.forEach(item => {
                hdfc_credits += item.creditTrn || 0;
                hdfc_debits += item.debitTrn || 0;
            });
            
            let gl_credits = 0;
            let gl_debits = 0;
            glSide.forEach(item => {
                gl_credits += item.creditTrn || 0;
                gl_debits += item.debitTrn || 0;
            });
            
            // Calculate Difference
            const total_credits = gl_credits + hdfc_credits;
            const total_debits = gl_debits + hdfc_debits;
            const netDiffVal = total_credits - total_debits;
            const netDiffAbs = Math.abs(netDiffVal);
            
            let diffColor = 'var(--success)';
            let diffDisplay = '₹ 0.00';
            
            if (netDiffAbs >= 0.01) {
                if (netDiffVal > 0) {
                    diffColor = 'var(--success)';
                    diffDisplay = formatCurrency(netDiffAbs) + ' CR';
                } else {
                    diffColor = 'var(--danger)';
                    diffDisplay = formatCurrency(netDiffAbs) + ' DR';
                }
            }
            
            const isAllReconciled = txns.every(t => t.reconciled);
            const statusHtml = isAllReconciled 
                ? '<span class="flag-badge auto">Reconciled</span>' 
                : '<span class="flag-badge pending">Pending</span>';
                
            // Generate rows for sub-table
            let subTableRows = '';
            txns.forEach(item => {
                const typeClass = item.type === 'HDFC' ? 'HDFC' : 'x' + item.type;
                const recStatus = item.reconciled 
                    ? '<span style="color:var(--success); font-weight:700;">Reconciled</span>' 
                    : '<span style="color:var(--text-muted);">Pending</span>';
                const isSubSelected = state.selectedIds.has(item.id);
                
                subTableRows += `
                    <tr style="border-bottom: 1px solid rgba(255,255,255,0.03);">
                        <td style="padding: 6px 4px; text-align: center; vertical-align: middle; width: 50px;">
                            <div class="recon-checkbox ${isSubSelected ? 'checked' : ''}" data-id="${item.id}" onclick="toggleSelect('${item.id}')">
                                <i data-lucide="check"></i>
                            </div>
                        </td>
                        <td class="type-cell ${typeClass}" style="padding: 6px 4px;">${item.type}</td>
                        <td style="padding: 6px 4px; color: var(--text-primary); max-width: 400px; word-break: break-word;">${item.description}</td>
                        <td style="padding: 6px 4px; text-align: right; color: var(--success); font-family: monospace;">${item.creditTrn > 0 ? formatCurrency(item.creditTrn) : '-'}</td>
                        <td style="padding: 6px 4px; text-align: right; color: var(--danger); font-family: monospace;">${item.debitTrn > 0 ? formatCurrency(item.debitTrn) : '-'}</td>
                        <td style="padding: 6px 4px; text-align: center;">
                            <span contenteditable="true" class="editable-flag-cell" style="min-width: 60px;" onblur="updateNPCIFlag('${item.id}', this.textContent)">${item.flag || 'NPCI'}</span>
                        </td>
                        <td style="padding: 6px 4px; text-align: center;">
                            <span contenteditable="true" class="editable-flag-cell" style="min-width: 80px;" onblur="updateNPCIActualDate('${item.id}', this.textContent)">${item.actualDate}</span>
                        </td>
                        <td style="padding: 6px 4px; text-align: center;">${recStatus}</td>
                    </tr>
                `;
            });
            
            html += `
                <tr style="background: ${isAllReconciled ? 'rgba(16, 185, 129, 0.03)' : 'transparent'}; border-bottom: 1px solid rgba(255,255,255,0.05);">
                    <td style="text-align: center; font-weight: 600; padding: 12px 8px;">
                        <span style="cursor: pointer; color: #a855f7; display: inline-flex; align-items: center; gap: 4px;" onclick="toggleNpciRowDetails('${dateId}', '${d}')">
                            <i data-lucide="chevron-right" class="npci-toggle-icon-${dateId}" style="width:12px; height:12px; transition: transform 0.2s; transform: ${isExpanded ? 'rotate(90deg)' : 'rotate(0deg)'};"></i> 
                            ${d}
                        </span>
                    </td>
                    <td style="text-align: right; color: var(--success); font-family: monospace; font-weight: 600; padding: 12px 8px;">${gl_credits > 0 ? formatCurrency(gl_credits) : '-'}</td>
                    <td style="text-align: right; color: var(--danger); font-family: monospace; font-weight: 600; padding: 12px 8px;">${gl_debits > 0 ? formatCurrency(gl_debits) : '-'}</td>
                    <td style="text-align: right; color: var(--success); font-family: monospace; font-weight: 600; padding: 12px 8px;">${hdfc_credits > 0 ? formatCurrency(hdfc_credits) : '-'}</td>
                    <td style="text-align: right; color: var(--danger); font-family: monospace; font-weight: 600; padding: 12px 8px;">${hdfc_debits > 0 ? formatCurrency(hdfc_debits) : '-'}</td>
                    <td style="text-align: right; font-family: monospace; font-weight: 700; padding: 12px 8px; color: ${diffColor};">
                        ${diffDisplay}
                    </td>
                    <td style="text-align: center; padding: 12px 8px;">${statusHtml}</td>
                    <td style="text-align: center; vertical-align: middle; padding: 12px 8px;">
                        <div class="recon-checkbox npci-date-checkbox npci-checkbox-${d.replace(/\//g, '_')} ${state.selectedNpciDates && state.selectedNpciDates.has(d) ? 'checked' : ''}" onclick="toggleNpciDateSelection('${d}')">
                            <i data-lucide="check"></i>
                        </div>
                    </td>
                </tr>
                <!-- Sub-table row containing detailed NPCI transactions -->
                <tr id="npci-details-${dateId}" style="display: ${isExpanded ? 'table-row' : 'none'}; background: rgba(255,255,255,0.015);">
                    <td colspan="8" style="padding: 12px 24px;">
                        <div style="border-left: 3px solid #a855f7; padding-left: 16px; background: rgba(255,255,255,0.005); padding-top: 8px; padding-bottom: 8px; border-radius: 0 8px 8px 0;">
                            <h4 style="font-size: 11px; margin-bottom: 8px; color: #a855f7; font-weight: 700; display: flex; align-items: center; gap: 4px;">
                                <i data-lucide="list"></i> ${d} ના વ્યવહારોની વિગત (NPCI Transactions for ${d})
                            </h4>
                            <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                                <thead>
                                    <tr style="border-bottom: 1px solid rgba(255,255,255,0.1); color: var(--text-muted); font-weight: 600;">
                                        <th style="padding: 6px 4px; text-align: center; width: 50px;">SELECT</th>
                                        <th style="padding: 6px 4px; text-align: left; width: 80px;">TYPE</th>
                                        <th style="padding: 6px 4px; text-align: left;">DESCRIPTION</th>
                                        <th style="padding: 6px 4px; text-align: right; width: 120px;">CREDIT</th>
                                        <th style="padding: 6px 4px; text-align: right; width: 120px;">DEBIT</th>
                                        <th style="padding: 6px 4px; text-align: center; width: 100px;">FLAG (Editable)</th>
                                        <th style="padding: 6px 4px; text-align: center; width: 150px;">ACTUAL DATE (Editable)</th>
                                        <th style="padding: 6px 4px; text-align: center; width: 100px;">STATUS</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${subTableRows}
                                </tbody>
                            </table>
                        </div>
                    </td>
                </tr>
            `;
        });
    }
    
    html += `
            </tbody>
        </table>
    `;
    
    tableBody.innerHTML = html;
    entriesCount.textContent = `કુલ ${dates.length} તારીખોના પેન્ડીંગ NPCI વ્યવહારોનું પીવોટ ટેબલ`;
    paginationControls.innerHTML = '';
    lucide.createIcons();
    updateBulkActionButtons();
}

// Toggle daily NPCI date reconciliation
function toggleNPCIDateReconciliation(dateStr, isCurrentlyReconciled) {
    const targetState = !isCurrentlyReconciled;
    const npciTxns = state.mergedData.filter(item => item.actualDate === dateStr && item.flag && item.flag.toUpperCase() === 'NPCI');
    
    if (targetState) {
        // Reconcile all
        const groupId = 'npci_group_' + dateStr.replace(/\//g, '_') + '_' + state.matchGroupCounter;
        state.matchGroupCounter++;
        npciTxns.forEach(item => {
            item.reconciled = true;
            item.matchGroupId = groupId;
        });
        showToast(`${dateStr} ના તમામ NPCI વ્યવહારો સફળતાપૂર્વક રીકન્સાઇલ કરવામાં આવ્યા.`, 'success');
    } else {
        // Unreconcile all
        npciTxns.forEach(item => {
            item.reconciled = false;
            item.matchGroupId = null;
        });
        showToast(`${dateStr} ના તમામ NPCI વ્યવહારો પેન્ડીંગ કરવામાં આવ્યા.`, 'info');
    }
    
    saveStateToLocalStorage();
    refreshLedgerCounts();
    renderTable();
    
    // Check if the live difference is now zero (0.00) after pivot reconciliation
    const finalDiff = calculateCurrentDifference();
    if (Math.abs(finalDiff) < 0.01 && state.mergedData.length > 0) {
        showToast('મેળવણી પત્રક પૂર્ણ! ડિફરન્સ શૂન્ય છે.', 'success');
    }
}

// Toggle checkbox selection for daily NPCI dates
function toggleNpciDateSelection(dateStr) {
    if (!state.selectedNpciDates) {
        state.selectedNpciDates = new Set();
    }
    
    const txns = state.mergedData.filter(item => item.actualDate === dateStr && item.flag && item.flag.toUpperCase() === 'NPCI' && !item.reconciled);
    
    if (state.selectedNpciDates.has(dateStr)) {
        state.selectedNpciDates.delete(dateStr);
        txns.forEach(item => state.selectedIds.delete(item.id));
    } else {
        state.selectedNpciDates.add(dateStr);
        txns.forEach(item => state.selectedIds.add(item.id));
    }
    
    // Toggle checkmark styling on checkbox element without full table rebuild
    const checkbox = document.querySelector(`.npci-checkbox-${dateStr.replace(/\//g, '_')}`);
    if (checkbox) {
        if (state.selectedNpciDates.has(dateStr)) {
            checkbox.classList.add('checked');
        } else {
            checkbox.classList.remove('checked');
        }
    }
    
    // Also toggle checkboxes of individual entries for this date
    txns.forEach(item => {
        const subCheckbox = document.querySelector(`.recon-checkbox[data-id="${item.id}"]`);
        if (subCheckbox) {
            if (state.selectedIds.has(item.id)) {
                subCheckbox.classList.add('checked');
            } else {
                subCheckbox.classList.remove('checked');
            }
        }
    });
    
    updateBulkActionButtons();
    updateBRSLiveWidget();
}

// Toggle visibility of detailed transactions inside NPCI summary pivot table
function toggleNpciRowDetails(dateId, dateStr) {
    if (!state.openNpciDateDetails) {
        state.openNpciDateDetails = new Set();
    }
    
    const key = dateStr || dateId;
    const row = document.getElementById(`npci-details-${dateId}`);
    const icon = document.querySelector(`.npci-toggle-icon-${dateId}`);
    
    if (state.openNpciDateDetails.has(key)) {
        state.openNpciDateDetails.delete(key);
        if (row) row.style.display = 'none';
        if (icon) icon.style.transform = 'rotate(0deg)';
    } else {
        state.openNpciDateDetails.add(key);
        if (row) row.style.display = 'table-row';
        if (icon) icon.style.transform = 'rotate(90deg)';
    }
}

// Update ACTUAL DATE for a specific NPCI transaction
function updateNPCIActualDate(itemId, newDate) {
    const item = state.mergedData.find(t => t.id === itemId);
    if (!item) return;
    
    // Clean and validate date format
    const parsedDate = parseDateString(newDate);
    if (!parsedDate || isNaN(parsedDate.getTime())) {
        showToast('અમાન્ય તારીખ ફોર્મેટ! કૃપા કરીને DD/MM/YYYY ફોર્મેટમાં લખો.', 'error');
        return;
    }
    
    const formattedDate = formatDateToSlash(parsedDate);
    const oldDate = item.actualDate;
    
    // Check if anything actually changed
    if (oldDate === formattedDate) {
        return;
    }
    
    item.actualDate = formattedDate;
    
    if (!state.openNpciDateDetails) {
        state.openNpciDateDetails = new Set();
    }
    
    // Ensure both old and new date detail rows stay open so user sees the entry move!
    if (oldDate) state.openNpciDateDetails.add(oldDate);
    state.openNpciDateDetails.add(formattedDate);
    
    saveStateToLocalStorage();
    refreshLedgerCounts();
    updateBRSLiveWidget();
    renderTable();
    
    showToast(`વ્યવહારની Actual Date બદલીને ${formattedDate} ના લિસ્ટમાં મોકલવામાં આવી.`, 'info');
}

// Update FLAG for a specific NPCI transaction
function updateNPCIFlag(itemId, newFlag) {
    const item = state.mergedData.find(t => t.id === itemId);
    if (!item) return;
    
    const cleanFlag = String(newFlag || '').trim().toUpperCase();
    
    // Check if anything actually changed
    if (item.flag === cleanFlag) {
        return;
    }
    
    item.flag = cleanFlag || 'DAILY';
    
    if (item.flag === 'NPCI') {
        const extracted = extractDateFromDescription(item.description, item.actualDate);
        if (extracted) {
            item.actualDate = extracted;
            if (!state.openNpciDateDetails) state.openNpciDateDetails = new Set();
            state.openNpciDateDetails.add(extracted);
        }
    }
    
    saveStateToLocalStorage();
    refreshLedgerCounts();
    updateBRSLiveWidget();
    renderTable();
}

// Generate pagination button controls
function renderPagination(totalPages) {
    if (totalPages <= 1) {
        paginationControls.innerHTML = '';
        return;
    }
    
    let html = '';
    
    // Back button
    html += `<button class="page-btn" ${state.currentPage === 1 ? 'disabled' : ''} onclick="changePage(${state.currentPage - 1})"><i data-lucide="chevron-left" style="width:12px;height:12px;"></i></button>`;
    
    // Page numbers
    for (let i = 1; i <= totalPages; i++) {
        if (i === 1 || i === totalPages || (i >= state.currentPage - 1 && i <= state.currentPage + 1)) {
            html += `<button class="page-btn ${state.currentPage === i ? 'active' : ''}" onclick="changePage(${i})">${i}</button>`;
        } else if (i === 2 || i === totalPages - 1) {
            html += `<span style="padding: 0 4px; color: var(--text-muted);">...</span>`;
        }
    }
    
    // Next button
    html += `<button class="page-btn" ${state.currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${state.currentPage + 1})"><i data-lucide="chevron-right" style="width:12px;height:12px;"></i></button>`;
    
    paginationControls.innerHTML = html;
}

function changePage(page) {
    state.currentPage = page;
    renderTable();
    // Scroll table container back to top
    document.querySelector('.table-container').scrollTop = 0;
}

// Export data sets (Pending or Reconciled) to actual Excel Sheets (.XLSX)
function exportToExcel(tabType) {
    const listToExport = state.mergedData.filter(item => 
        (tabType === 'pending' && !item.reconciled && !item.matchGroupId) || 
        (tabType === 'reconciled' && (item.reconciled || item.matchGroupId))
    );
    
    if (listToExport.length === 0) {
        showToast('નિકાસ (Export) કરવા માટે કોઈ વ્યવહાર નથી!', 'warning');
        return;
    }
    
    // Sort by specific type order: HDFC -> 345051 -> 3496 -> 3493, then by largest transaction amount descending
    const sortedList = [...listToExport];
    const typeOrder = {
        'HDFC': 1,
        '345051': 2,
        '3496': 3,
        '3493': 4
    };
    
    if (tabType === 'reconciled') {
        const groupMaxAmt = {};
        listToExport.forEach(item => {
            const amt = Math.max(item.creditTrn, item.debitTrn);
            if (item.matchGroupId) {
                groupMaxAmt[item.matchGroupId] = Math.max(groupMaxAmt[item.matchGroupId] || 0, amt);
            }
        });
        
        sortedList.sort((a, b) => {
            const groupIdA = a.matchGroupId || a.id;
            const groupIdB = b.matchGroupId || b.id;
            
            const maxAmtA = a.matchGroupId ? (groupMaxAmt[a.matchGroupId] || 0) : Math.max(a.creditTrn, a.debitTrn);
            const maxAmtB = b.matchGroupId ? (groupMaxAmt[b.matchGroupId] || 0) : Math.max(b.creditTrn, b.debitTrn);
            
            // Rule 1: Sort primarily by group maximum transaction amount descending (largest amount on top)
            if (Math.abs(maxAmtA - maxAmtB) > 0.001) {
                return maxAmtB - maxAmtA;
            }
            
            // Rule 2: If belonging to the same match group (or same single transaction), sort HDFC on top, then Credit before Debit
            if (groupIdA === groupIdB) {
                const isHdfcA = getFileCategory(a) === 'HDFC';
                const isHdfcB = getFileCategory(b) === 'HDFC';
                if (isHdfcA !== isHdfcB) {
                    return isHdfcA ? -1 : 1; // HDFC first
                }
                return b.creditTrn - a.creditTrn; // Credit before Debit
            }
            
            // Rule 3: Group entries of the same amount by their group ID to keep each matched group's items side-by-side
            return groupIdA.localeCompare(groupIdB);
        });
    } else {
        sortedList.sort((a, b) => {
            const typeA = getFileCategory(a);
            const typeB = getFileCategory(b);
            const orderA = typeOrder[typeA] || 99;
            const orderB = typeOrder[typeB] || 99;
            
            if (orderA !== orderB) {
                return orderA - orderB;
            }
            
            const amtA = Math.max(a.creditTrn, a.debitTrn);
            const amtB = Math.max(b.creditTrn, b.debitTrn);
            return amtB - amtA;
        });
    }
    
    // Helper to format numeric values to exactly two decimal places (0.00)
    function formatDoubleDigit(val) {
        if (val === null || val === undefined || isNaN(val)) return '0.00';
        return parseFloat(val).toFixed(2);
    }

    // Helper to format cells with styles and Excel number types
    function c(val, style = null) {
        let cell = { v: val === null || val === undefined ? '' : val };
        
        if (typeof val === 'number') {
            cell.t = 'n';
            cell.z = '#,##0.00';
        }
        
        if (style) {
            cell.s = style;
        }
        
        if (!style && typeof val !== 'number') {
            return val === null || val === undefined ? '' : val;
        }
        
        return cell;
    }
    
    // Styles (All with black borders for direct printing layout)
    const titleStyle = {
        font: { name: 'Arial', sz: 14, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' }
    };
    
    const subtitleStyle = {
        font: { name: 'Arial', sz: 11, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' }
    };
    
    const borderThin = { style: 'thin', color: { rgb: '000000' } };
    
    const labelStyle = {
        font: { name: 'Arial', sz: 9 },
        border: {
            bottom: borderThin
        }
    };
    
    const labelBoldStyle = {
        font: { name: 'Arial', sz: 9, bold: true },
        border: {
            bottom: borderThin
        }
    };
    
    const noBorder = {
        font: { name: 'Arial', sz: 9 }
    };
    
    const noBorderBold = {
        font: { name: 'Arial', sz: 9, bold: true }
    };
    
    const headerStyle = {
        font: { name: 'Arial', sz: 9, bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        },
        fill: { fgColor: { rgb: '000000' } }
    };

    const boxStyle = {
        font: { name: 'Arial', sz: 9 },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };

    const boxStyleWrap = {
        font: { name: 'Arial', sz: 9 },
        alignment: { wrapText: true, vertical: 'center' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };

    const sectionSubtitleStyle = {
        font: { name: 'Arial', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '000000' } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };

    const totalHeaderStyle = {
        font: { name: 'Arial', sz: 9, bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        fill: { fgColor: { rgb: '000000' } },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };

    const totalHeaderRightStyle = {
        font: { name: 'Arial', sz: 9, bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { horizontal: 'right', vertical: 'center' },
        fill: { fgColor: { rgb: '000000' } },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };

    const footerBankStyle = {
        font: { name: 'Arial', sz: 11, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' }
    };

    const footerRoleStyle = {
        font: { name: 'Arial', sz: 10, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' }
    };
    
    const boxRight = {
        font: { name: 'Arial', sz: 9 },
        alignment: { horizontal: 'right' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };
    
    const boxCenter = {
        font: { name: 'Arial', sz: 9 },
        alignment: { horizontal: 'center' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };
    
    const boxBold = {
        font: { name: 'Arial', sz: 9, bold: true },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };
    
    const boxBoldRight = {
        font: { name: 'Arial', sz: 9, bold: true },
        alignment: { horizontal: 'right' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };
    
    const dataCell = {
        font: { name: 'Arial', sz: 9 },
        alignment: { wrapText: true, vertical: 'center' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };
    
    const dataCellRight = {
        font: { name: 'Arial', sz: 9 },
        alignment: { horizontal: 'right' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };
    
    const dataCellCenter = {
        font: { name: 'Arial', sz: 9 },
        alignment: { horizontal: 'center' },
        border: {
            top: borderThin,
            bottom: borderThin,
            left: borderThin,
            right: borderThin
        }
    };
    
    // Structure worksheet based on tab type
    let worksheet;
    
    if (tabType === 'pending') {
        const hdfcDates = state.mergedData
            .filter(item => item.type === 'HDFC' || item.parentType === 'HDFC')
            .map(item => parseDateString(item.date))
            .filter(d => d !== null);
        
        let fromDateStr = '';
        let toDateStr = '';
        if (hdfcDates.length > 0) {
            const minDate = new Date(Math.min(...hdfcDates));
            const maxDate = new Date(Math.max(...hdfcDates));
            fromDateStr = formatDateOnly(minDate);
            toDateStr = formatDateOnly(maxDate);
        }
        
        const hdfc_bal = state.files.HDFC.loaded ? state.files.HDFC.balance : 0;
        const gl_345051_bal = state.files['345051'].loaded ? state.files['345051'].balance : 0;
        const gl_3493_bal = state.files['3493'].loaded ? Math.abs(state.files['3493'].balance) : 0;
        const gl_3496_bal = state.files['3496'].loaded ? Math.abs(state.files['3496'].balance) : 0;
        
        // Sum up pending credits/debits for GLs (345051, 3493, 3496)
        let gl_pending_credits = 0;
        let gl_pending_debits = 0;
        let hdfc_pending_credits = 0;
        let hdfc_pending_debits = 0;
        
        state.mergedData.forEach(item => {
            if (!item.reconciled && !item.matchGroupId) {
                const cat = getFileCategory(item);
                if (cat === 'HDFC') {
                    hdfc_pending_credits += item.creditTrn;
                    hdfc_pending_debits += item.debitTrn;
                } else {
                    gl_pending_credits += item.creditTrn;
                    gl_pending_debits += item.debitTrn;
                }
            }
        });
        
        // BRS calculations
        const val_row5 = gl_345051_bal;
        const val_row6 = val_row5 - gl_pending_credits + gl_pending_debits;
        const val_row7 = val_row6 - hdfc_pending_credits + hdfc_pending_debits;
        const diff_any = hdfc_bal + val_row7; // Net difference
        const other_gls_sum = gl_3496_bal - gl_3493_bal;
        const total_unreconciled_diff = diff_any - other_gls_sum;
        
        const reconDateStr = state.currentDate || new Date().toISOString().slice(0, 10);
        let formattedCurrentDate = reconDateStr;
        const dateObj = parseDateString(reconDateStr);
        if (dateObj) {
            const day = String(dateObj.getDate()).padStart(2, '0');
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const year = dateObj.getFullYear();
            formattedCurrentDate = `${day}/${month}/${year}`;
        }
        
        // Build 2D array of rows matching BRS template
        const sheetData = [
            [c('THE JUNAGADH COMMERCIAL CO-OPERATIVE BANK LTD.', titleStyle), '', '', '', '', '', '', ''],
            [c(`HDFC RECONCILIATION STATEMENT AS ON ${formattedCurrentDate}`, subtitleStyle), '', '', '', '', '', '', ''],
            [c('HDFC BANK BALANCE', noBorderBold), '', '', '', '', '', '', c(hdfc_bal, boxBoldRight)],
            [c('OUR BANK GL BALANCE', noBorderBold), '', '', '', '', '', '', c(gl_345051_bal, boxBoldRight)],
            ['', '', '', '', '', '', '', ''], // Spacer row (empty, no borders)
            
            // Row 7: Table Headers
            [c('Date', headerStyle), c('Description', headerStyle), c('Side', headerStyle), c('Sign', headerStyle), c('Credit Pending', headerStyle), c('Debit Pending', headerStyle), '', c('Net Balance', headerStyle)],
            
            // Row 8: Our Site row
            [c(fromDateStr, boxCenter), c('-', boxStyle), c('Our Site', boxCenter), c('', boxCenter), c('', boxCenter), c('', boxCenter), '', c(val_row5, boxRight)],
            
            // Row 9: Transaction at our site but not credit/Debit by Bank
            [c('', boxCenter), c('Transaction at our site but not credit/Debit by Bank', boxStyleWrap), c('Opp.Bank', boxCenter), c('+', boxCenter), c(gl_pending_credits, boxRight), c(gl_pending_debits, boxRight), '', c(val_row6, boxRight)],
            
            // Row 10: Transaction at hdfc Bank but not effect in Our GL Balance
            [c('', boxCenter), c('Transaction at hdfc Bank but not effect in Our GL Balance', boxStyleWrap), c('Our Bank', boxCenter), c('+', boxCenter), c(hdfc_pending_credits, boxRight), c(hdfc_pending_debits, boxRight), '', c(val_row7, boxRight)],
            
            // Row 11: Net Balance as per our side GL Statment
            ['', c('Net Balance as per our side GL Statment', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), '', c(val_row7, boxRight)],
            
            // Row 12: Net Balance as per Other Side (HDFC BANK)
            ['', c('Net Balance as per Other Side (HDFC BANK)', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), '', c(hdfc_bal, boxRight)],
            
            // Row 13: Difference if any
            ['', c('Difference if any', labelBoldStyle), c('', labelBoldStyle), c('', labelBoldStyle), c('', labelBoldStyle), c('', labelBoldStyle), '', c(diff_any, boxBoldRight)],
            
            // Row 14: GL 3493
            ['', c('GL 3493', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), '', c(gl_3493_bal, boxRight)],
            
            // Row 15: GL 3496
            ['', c('GL 3496', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), c('', labelStyle), '', c(gl_3496_bal, boxRight)],
            
            // Row 16: TOTAL UNRECONCILED DIFFERENCE
            ['', c('TOTAL UNRECONCILED DIFFERENCE', labelBoldStyle), c('', labelBoldStyle), c('', labelBoldStyle), c('', labelBoldStyle), c('', labelBoldStyle), '', c(total_unreconciled_diff, boxBoldRight)],
            
            [], // Row 17 (blank spacer)
            []  // Row 18 (blank spacer)
        ];
        
        // ----------------------------------------------------
        // SECTION 1: Transaction at our site but not credit/Debit by Bank
        // ----------------------------------------------------
        const glList = listToExport.filter(item => getFileCategory(item) !== 'HDFC');
        
        // Sort Section 1 by FLAG alphabetically, then DATE chronologically ascending (oldest first)
        glList.sort((a, b) => {
            const flagA = (a.flag || 'DAILY').trim().toUpperCase();
            const flagB = (b.flag || 'DAILY').trim().toUpperCase();
            if (flagA !== flagB) {
                return flagA.localeCompare(flagB);
            }
            const dateA = parseDateString(a.date) || 0;
            const dateB = parseDateString(b.date) || 0;
            return dateA - dateB;
        });
        
        let gl_total_credits = 0;
        let gl_total_debits = 0;
        glList.forEach(item => {
            gl_total_credits += item.creditTrn || 0;
            gl_total_debits += item.debitTrn || 0;
        });
        
        sheetData.push(
            [c('Transaction at our site but not credit/Debit by Bank', sectionSubtitleStyle), '', '', '', '', '', '', ''],
            [c('DATE', headerStyle), c('DISCRIPTION', headerStyle), c('TYPE', headerStyle), c('FLAG', headerStyle), c('REF NO', headerStyle), c('ACTUAL DATE', headerStyle), c('CREDIT TRN', headerStyle), c('DEBIT TRN', headerStyle)]
        );
        
        glList.forEach(item => {
            sheetData.push([
                c(item.date, dataCellCenter),
                c(item.description, dataCell),
                c(item.type, dataCellCenter),
                c(item.flag, dataCellCenter),
                c(getDisplayRefNo(item), dataCellCenter),
                c(item.actualDate, dataCellCenter),
                item.creditTrn > 0 ? c(item.creditTrn, dataCellRight) : c('', dataCellRight),
                item.debitTrn > 0 ? c(item.debitTrn, dataCellRight) : c('', dataCellRight)
            ]);
        });
        
        sheetData.push([
            c('TOTAL', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c(gl_total_credits, totalHeaderRightStyle),
            c(gl_total_debits, totalHeaderRightStyle)
        ]);
        
        // Spacer
        sheetData.push([]);
        sheetData.push([]);
        
        // ----------------------------------------------------
        // SECTION 2: Transaction at HDFC Bank but not effect in Our GL Balance
        // ----------------------------------------------------
        const hdfcList = listToExport.filter(item => getFileCategory(item) === 'HDFC');
        
        // Sort Section 2 by FLAG alphabetically, then DATE chronologically ascending (oldest first)
        hdfcList.sort((a, b) => {
            const flagA = (a.flag || 'DAILY').trim().toUpperCase();
            const flagB = (b.flag || 'DAILY').trim().toUpperCase();
            if (flagA !== flagB) {
                return flagA.localeCompare(flagB);
            }
            const dateA = parseDateString(a.date) || 0;
            const dateB = parseDateString(b.date) || 0;
            return dateA - dateB;
        });
        
        let hdfc_total_credits = 0;
        let hdfc_total_debits = 0;
        hdfcList.forEach(item => {
            hdfc_total_credits += item.creditTrn || 0;
            hdfc_total_debits += item.debitTrn || 0;
        });
        
        sheetData.push(
            [c('Transaction at hdfc Bank but not effect in Our GL Balance', sectionSubtitleStyle), '', '', '', '', '', '', ''],
            [c('DATE', headerStyle), c('DISCRIPTION', headerStyle), c('TYPE', headerStyle), c('FLAG', headerStyle), c('REF NO', headerStyle), c('ACTUAL DATE', headerStyle), c('CREDIT TRN', headerStyle), c('DEBIT TRN', headerStyle)]
        );
        
        hdfcList.forEach(item => {
            sheetData.push([
                c(item.date, dataCellCenter),
                c(item.description, dataCell),
                c(item.type, dataCellCenter),
                c(item.flag, dataCellCenter),
                c(getDisplayRefNo(item), dataCellCenter),
                c(item.actualDate, dataCellCenter),
                item.creditTrn > 0 ? c(item.creditTrn, dataCellRight) : c('', dataCellRight),
                item.debitTrn > 0 ? c(item.debitTrn, dataCellRight) : c('', dataCellRight)
            ]);
        });
        
        sheetData.push([
            c('TOTAL', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c('', totalHeaderStyle),
            c(hdfc_total_credits, totalHeaderRightStyle),
            c(hdfc_total_debits, totalHeaderRightStyle)
        ]);
        
        // Add signature block at the bottom
        sheetData.push(
            [],
            [],
            [c('THE JUNAGADH COMMERCIAL CO-OPERATIVE BANK LTD.', footerBankStyle), '', '', '', '', '', '', ''],
            [],
            ['', c('OFFICER', footerRoleStyle), '', '', '', c('MANAGER', footerRoleStyle), '', '']
        );
        
        worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        
        const merges = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } }
        ];
        
        sheetData.forEach((row, idx) => {
            if (row && row[0] && typeof row[0] === 'object' && row[0].v) {
                const str = String(row[0].v);
                if (str === 'Transaction at our site but not credit/Debit by Bank' || 
                    str === 'Transaction at hdfc Bank but not effect in Our GL Balance' || 
                    str === 'THE JUNAGADH COMMERCIAL CO-OPERATIVE BANK LTD.') {
                    merges.push({ s: { r: idx, c: 0 }, e: { r: idx, c: 7 } });
                }
            }
            if (row && row[1] && typeof row[1] === 'object' && row[1].v === 'OFFICER') {
                merges.push({ s: { r: idx, c: 1 }, e: { r: idx, c: 2 } });
            }
            if (row && row[5] && typeof row[5] === 'object' && row[5].v === 'MANAGER') {
                merges.push({ s: { r: idx, c: 5 }, e: { r: idx, c: 6 } });
            }
        });
        
        worksheet['!merges'] = merges;
    } else {
        // Reconciled download doesn't need BRS Summary, keep original simple format but add Bank Title
        const sheetData = [
            [c('THE JUNAGADH COMMERCIAL CO-OPERATIVE BANK LTD.', titleStyle), '', '', '', '', '', '', ''],
            [c('RECONCILED TRANSACTIONS REPORT', subtitleStyle), '', '', '', '', '', '', ''],
            [],
            [c('DATE', headerStyle), c('DISCRIPTION', headerStyle), c('TYPE', headerStyle), c('FLAG', headerStyle), c('REF NO', headerStyle), c('ACTUAL DATE', headerStyle), c('CREDIT TRN', headerStyle), c('DEBIT TRN', headerStyle)]
        ];
        
        sortedList.forEach(item => {
            sheetData.push([
                c(item.date, dataCellCenter),
                c(item.description, dataCell),
                c(item.type, dataCellCenter),
                c(item.flag, dataCellCenter),
                c(getDisplayRefNo(item), dataCellCenter),
                c(item.actualDate, dataCellCenter),
                item.creditTrn > 0 ? c(item.creditTrn, dataCellRight) : c('', dataCellRight),
                item.debitTrn > 0 ? c(item.debitTrn, dataCellRight) : c('', dataCellRight)
            ]);
        });
        
        sheetData.push(
            [],
            [],
            [c('THE JUNAGADH COMMERCIAL CO-OPERATIVE BANK LTD.', footerBankStyle), '', '', '', '', '', '', ''],
            [],
            ['', c('OFFICER', footerRoleStyle), '', '', '', c('MANAGER', footerRoleStyle), '', '']
        );
        
        worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        
        const merges = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } }
        ];
        
        sheetData.forEach((row, idx) => {
            if (row && row[0] && typeof row[0] === 'object' && row[0].v === 'THE JUNAGADH COMMERCIAL CO-OPERATIVE BANK LTD.') {
                merges.push({ s: { r: idx, c: 0 }, e: { r: idx, c: 7 } });
            }
            if (row && row[1] && typeof row[1] === 'object' && row[1].v === 'OFFICER') {
                merges.push({ s: { r: idx, c: 1 }, e: { r: idx, c: 2 } });
            }
            if (row && row[5] && typeof row[5] === 'object' && row[5].v === 'MANAGER') {
                merges.push({ s: { r: idx, c: 5 }, e: { r: idx, c: 6 } });
            }
        });
        
        worksheet['!merges'] = merges;
    }
    
    // Set page setup for printer (A4, Margins, Gridlines, Fit to Width)
    worksheet['!pageSetup'] = {
        paperSize: 9, // A4
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0
    };
    
    worksheet['!margins'] = {
        left: 1.0,  // 1.0 inch
        right: 0.5, // 0.5 inch
        top: 0.5,   // 0.5 inch
        bottom: 0.5, // 0.5 inch
        header: 0.3,
        footer: 0.3
    };
    
    worksheet['!views'] = [
        { showGridLines: true }
    ];
    
    // Explicit Column Widths to fit A4 paper nicely without clipping
    if (tabType === 'pending') {
        worksheet['!cols'] = [
            { wch: 12 }, // Date (Col A)
            { wch: 45 }, // Description (Col B)
            { wch: 12 }, // Side (Col C)
            { wch: 10 }, // Sign (Col D)
            { wch: 18 }, // Credit Pending (Col E)
            { wch: 18 }, // Debit Pending (Col F)
            { wch: 18 }  // Net Balance (Col G)
        ];
    } else {
        worksheet['!cols'] = [
            { wch: 12 }, // Date (Col A)
            { wch: 45 }, // Description (Col B)
            { wch: 12 }, // Type (Col C)
            { wch: 12 }, // Flag (Col D)
            { wch: 20 }, // REF NO (Col E)
            { wch: 14 }, // Actual Date (Col F)
            { wch: 18 }, // Credit Trn (Col G)
            { wch: 18 }  // Debit Trn (Col H)
        ];
    }
    
    const workbook = XLSX.utils.book_new();
    const sheetName = tabType === 'pending' ? 'Pending Transactions' : 'Reconciled Transactions';
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    
    // Generate filename using the reconciliation date in the state
    const reconDateStr = state.currentDate || new Date().toISOString().slice(0, 10);
    let formattedReconDate = reconDateStr;
    const dateObj = parseDateString(reconDateStr);
    if (dateObj) {
        const day = String(dateObj.getDate()).padStart(2, '0');
        const month = String(dateObj.getMonth() + 1).padStart(2, '0');
        const year = dateObj.getFullYear();
        formattedReconDate = `${day}-${month}-${year}`;
    }
    const filename = `${tabType === 'pending' ? 'Pending' : 'Reconciled'}_Recon_${formattedReconDate}.xlsx`;
    
    // Trigger download
    XLSX.writeFile(workbook, filename);
    showToast(`${filename} સફળતાપૂર્વક ડાઉનલોડ થઇ ગઇ.`, 'success');
}

// Start a new day: Keep pending entries, discard reconciled ones, and clear file upload cards
function startNewDay() {
    if (state.mergedData.length === 0) {
        showToast('કોઈ વ્યવહાર ઉપલબ્ધ નથી!', 'warning');
        return;
    }
    
    // Ask for confirmation to prevent accidental clicks
    const confirmMsg = "શું તમે ખરેખર નવો દિવસ શરૂ કરવા માંગો છો?\nઆનાથી આજના રીકન્સાઇલ થયેલા વ્યવહારો સાફ થઈ જશે અને માત્ર બાકી (Pending) વ્યવહારો જ સેવ રહેશે.";
    if (!confirm(confirmMsg)) {
        return;
    }

    // 1. Create backup of previous state
    const backup = {
        mergedData: state.mergedData,
        files: state.files,
        currentDate: state.currentDate,
        matchGroupCounter: state.matchGroupCounter
    };
    localStorage.setItem('recon_new_day_backup', JSON.stringify(backup));
    btnUndoNewDay.style.display = 'block';

    // Keep only pending transactions and mark them as carried over from previous day
    state.mergedData = state.mergedData.filter(item => !item.reconciled);
    state.mergedData.forEach(item => {
        item.isCarriedOver = true;
    });
    
    // Retain previous day's closing balance for carry-forward when holiday has no transactions
    const prevHdfcBal = state.files.HDFC.loaded ? state.files.HDFC.balance : (state.files.HDFC.previousBalance || 0);
    const prev345051Bal = state.files['345051'].loaded ? state.files['345051'].balance : (state.files['345051'].previousBalance || 0);
    const prev3493Bal = state.files['3493'].loaded ? state.files['3493'].balance : (state.files['3493'].previousBalance || 0);
    const prev3496Bal = state.files['3496'].loaded ? state.files['3496'].balance : (state.files['3496'].previousBalance || 0);

    // Reset file loaded states so new day's files can be uploaded, carrying over previous day's balance
    state.files = {
        HDFC: { loaded: false, data: [], balance: 0, previousBalance: prevHdfcBal, rawName: '' },
        345051: { loaded: false, data: [], balance: 0, previousBalance: prev345051Bal, rawName: '' },
        3493: { loaded: false, data: [], balance: 0, previousBalance: prev3493Bal, rawName: '' },
        3496: { loaded: false, data: [], balance: 0, previousBalance: prev3496Bal, rawName: '' },
        history: { loaded: false, data: [], balance: 0, previousBalance: 0, rawName: '' }
    };
    
    state.selectedIds.clear();
    updateBulkActionButtons();
    state.currentPage = 1;
    state.matchGroupCounter = 0;
    
    // Clear file inputs
    fileInput.value = '';
    
    // Reset file card UIs
    const types = ['HDFC', '345051', '3493', '3496', 'history'];
    types.forEach(t => {
        const card = document.getElementById(`status-${t.toLowerCase()}`);
        if (card) {
            card.classList.remove('loaded');
            card.querySelector('.file-indicator').className = 'file-indicator empty';
            card.querySelector('.file-meta').textContent = t === 'history' ? 'અપલોડ બાકી છે (વૈકલ્પિક)' : 'અપલોડ બાકી છે';
            const trashBtn = card.querySelector('.clear-file-btn');
            if (trashBtn) trashBtn.style.display = 'none';
        }
        
        if (t !== 'history') {
            const balVal = document.getElementById(`val-${t.toLowerCase()}-bal`);
            if (balVal) {
                balVal.textContent = '₹ 0.00';
                balVal.classList.remove('negative-balance');
            }
        }
    });
    
    valTotalBal.textContent = '₹ 0.00';
    valTotalBal.classList.remove('negative-balance');
    
    // Increment date in input by 1 day
    const currentDateObj = parseDateString(state.currentDate);
    if (currentDateObj) {
        currentDateObj.setDate(currentDateObj.getDate() + 1);
        const yyyy = currentDateObj.getFullYear();
        const mm = String(currentDateObj.getMonth() + 1).padStart(2, '0');
        const dd = String(currentDateObj.getDate()).padStart(2, '0');
        state.currentDate = `${yyyy}-${mm}-${dd}`;
        currentDateInput.value = state.currentDate;
    }
    
    // Save to local storage
    saveStateToLocalStorage();
    
    // Refresh ledger counts
    refreshLedgerCounts();
    renderTable();
    
    showToast(`નવો દિવસ શરૂ થયો છે! ગઈકાલના પેન્ડીંગ વ્યવહારો સચવાયેલા છે. નવી HDFC અને GL ફાઇલો અપલોડ કરો.`, 'success');
}

// Restore backup if New Day was clicked by mistake
function undoNewDay() {
    const backupStr = localStorage.getItem('recon_new_day_backup');
    if (!backupStr) {
        showToast('પુનઃસ્થાપિત કરવા માટે કોઈ બેકઅપ મળ્યો નથી!', 'warning');
        return;
    }
    
    try {
        const backup = JSON.parse(backupStr);
        state.mergedData = backup.mergedData;
        state.files = backup.files;
        state.currentDate = backup.currentDate;
        state.matchGroupCounter = backup.matchGroupCounter;
        
        // Update input field
        currentDateInput.value = state.currentDate;
        
        // Remove backup from storage
        localStorage.removeItem('recon_new_day_backup');
        btnUndoNewDay.style.display = 'none';
        
        // Update UIs
        saveStateToLocalStorage();
        
        // Reset file UIs
        Object.keys(state.files).forEach(fileType => {
            const card = document.getElementById(`status-${fileType.toLowerCase()}`);
            if (card) {
                if (state.files[fileType].loaded) {
                    card.classList.add('loaded');
                    card.querySelector('.file-indicator').className = 'file-indicator success';
                    const rowCount = state.files[fileType].rowCount !== undefined 
                        ? state.files[fileType].rowCount 
                        : (state.files[fileType].data ? state.files[fileType].data.length : 0);
                    card.querySelector('.file-meta').textContent = `સક્રિય | ${rowCount} રોઝ | ${state.files[fileType].rawName}`;
                    const trashBtn = card.querySelector('.clear-file-btn');
                    if (trashBtn) trashBtn.style.display = 'flex';
                } else {
                    card.classList.remove('loaded');
                    card.querySelector('.file-indicator').className = 'file-indicator empty';
                    card.querySelector('.file-meta').textContent = fileType === 'history' ? 'અપલોડ બાકી છે (વૈકલ્પિક)' : 'અપલોડ બાકી છે';
                    const trashBtn = card.querySelector('.clear-file-btn');
                    if (trashBtn) trashBtn.style.display = 'none';
                }
            }
            
            if (fileType !== 'history') {
                const balVal = document.getElementById(`val-${fileType.toLowerCase()}-bal`);
                if (balVal) {
                    const balance = state.files[fileType].balance;
                    balVal.textContent = formatCurrency(balance);
                    if (balance < 0) {
                        balVal.classList.add('negative-balance');
                    } else {
                        balVal.classList.remove('negative-balance');
                    }
                }
            }
        });
        
        // Update net total balance
        let totalBal = 0;
        Object.values(state.files).forEach(f => {
            if (f.loaded) totalBal += f.balance;
        });
        valTotalBal.textContent = formatCurrency(totalBal);
        if (totalBal < 0) {
            valTotalBal.classList.add('negative-balance');
        } else {
            valTotalBal.classList.remove('negative-balance');
        }
        
        // Check loaded files to enable/disable buttons
        checkAllFilesLoaded();
        
        refreshLedgerCounts();
        renderTable();
        showToast('છેલ્લી પરિસ્થિતિ સફળતાપૂર્વક પાછી મેળવી લેવામાં આવી છે (Undo successful)!', 'success');
    } catch (e) {
        showToast('બેકઅપ રીસ્ટોર કરવામાં ભૂલ આવી.', 'error');
        console.error(e);
    }
}

// Reset App state
function resetApp() {
    state.files = {
        HDFC: { loaded: false, data: [], balance: 0, rawName: '' },
        345051: { loaded: false, data: [], balance: 0, rawName: '' },
        3493: { loaded: false, data: [], balance: 0, rawName: '' },
        3496: { loaded: false, data: [], balance: 0, rawName: '' },
        history: { loaded: false, data: [], balance: 0, rawName: '' }
    };
    state.mergedData = [];
    state.selectedIds.clear();
    updateBulkActionButtons();
    state.currentPage = 1;
    state.matchGroupCounter = 0;
    
    // Clear local storage
    localStorage.removeItem('recon_merged_data');
    localStorage.removeItem('recon_match_counter');
    localStorage.removeItem('recon_files_meta');
    
    // Reset optional history card UI
    const cardHistory = document.getElementById('status-history');
    if (cardHistory) {
        cardHistory.classList.remove('loaded');
        cardHistory.querySelector('.file-indicator').className = 'file-indicator empty';
        cardHistory.querySelector('.file-meta').textContent = 'અપલોડ બાકી છે (વૈકલ્પિક)';
    }
    
    // Reset file input
    fileInput.value = '';
    
    // Reset file indicator UIs
    const types = ['HDFC', '345051', '3493', '3496'];
    types.forEach(t => {
        const card = document.getElementById(`status-${t.toLowerCase()}`);
        card.classList.remove('loaded');
        card.querySelector('.file-indicator').className = 'file-indicator empty';
        card.querySelector('.file-meta').textContent = 'અપલોડ બાકી છે';
        
        document.getElementById(`val-${t.toLowerCase()}-bal`).textContent = '₹ 0.00';
        document.getElementById(`val-${t.toLowerCase()}-bal`).classList.remove('negative-balance');
    });
    
    valTotalBal.textContent = '₹ 0.00';
    valTotalBal.classList.remove('negative-balance');
    
    // Disable buttons
    btnAutoReconcile.disabled = true;
    btnExportPending.disabled = true;
    btnExportReconciled.disabled = true;
    
    renderTable();
    showToast('સિસ્ટમ રીસેટ કરવામાં આવી છે. નવી ફાઇલો અપલોડ કરો.', 'info');
}

// Read and parse uploaded Excel reconciliation sheet (.xlsx)
function processExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Check if it's a reconciled report (contains reconciled sheet)
            let isReconciledExcel = false;
            workbook.SheetNames.forEach(sheetName => {
                if (sheetName.toLowerCase().includes('reconciled') || sheetName.includes('મેળવણી')) {
                    isReconciledExcel = true;
                }
            });
            
            if (isReconciledExcel) {
                // Clear state for fresh load of past reconciled report
                state.mergedData = [];
                const reconciledTxnsList = [];
                
                workbook.SheetNames.forEach(sheetName => {
                    if (sheetName.toLowerCase().includes('reconciled') || sheetName.includes('મેળવણી')) {
                        const sheet = workbook.Sheets[sheetName];
                        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                        
                        rows.forEach(row => {
                            if (!row || row.length < 3) return;
                            
                            const dateStr = String(row[0] || '').trim();
                            const desc = String(row[1] || '').trim();
                            const type = String(row[2] || '').trim().toUpperCase();
                            
                            if (!dateStr || !desc || !type) return;
                            if (dateStr.toLowerCase().includes('date') || dateStr.toLowerCase().includes('statement') || dateStr.toLowerCase().includes('bank')) return;
                            if (desc.toLowerCase().includes('transaction') || desc.toLowerCase().includes('particular') || desc.toLowerCase().includes('description')) return;
                            if (dateStr === 'TOTAL' || desc === 'TOTAL') return;
                            
                            let parsedDate = formatDateToSlash(dateStr);
                            if (!isNaN(dateStr) && Number(dateStr) > 30000) {
                                const excelDate = new Date((Number(dateStr) - 25569) * 86400 * 1000);
                                parsedDate = formatDateToSlash(excelDate);
                            }
                            
                            const credit = parseFloat(String(row[6] || '').replace(/,/g, '')) || 0;
                            const debit = parseFloat(String(row[7] || '').replace(/,/g, '')) || 0;
                            
                            if (credit === 0 && debit === 0) return;
                            
                            reconciledTxnsList.push({
                                date: parsedDate,
                                description: desc,
                                type: type,
                                credit: credit,
                                debit: debit,
                                refNo: String(row[4] || '').trim(),
                                flag: String(row[3] || 'DAILY').trim().toUpperCase(),
                                actualDate: formatDateToSlash(String(row[5] || parsedDate).trim())
                            });
                        });
                    }
                });
                
                // Pair them up
                for (let i = 0; i < reconciledTxnsList.length; i += 2) {
                    const a = reconciledTxnsList[i];
                    const b = reconciledTxnsList[i + 1];
                    const groupId = 'recovered_excel_group_' + state.matchGroupCounter;
                    state.matchGroupCounter++;
                    
                    if (a) {
                        state.mergedData.push({
                            id: 'history_' + Math.random().toString(36).substr(2, 9),
                            date: a.date,
                            description: a.description,
                            type: a.type,
                            actualDate: a.actualDate || a.date,
                            creditTrn: a.credit,
                            debitTrn: a.debit,
                            refNo: a.refNo || '',
                            flag: a.flag || 'DAILY',
                            day: 0,
                            count: 0,
                            reconciled: true,
                            matchGroupId: groupId
                        });
                    }
                    if (b) {
                        state.mergedData.push({
                            id: 'history_' + Math.random().toString(36).substr(2, 9),
                            date: b.date,
                            description: b.description,
                            type: b.type,
                            actualDate: b.actualDate || b.date,
                            creditTrn: b.credit,
                            debitTrn: b.debit,
                            refNo: b.refNo || '',
                            flag: b.flag || 'DAILY',
                            day: 0,
                            count: 0,
                            reconciled: true,
                            matchGroupId: groupId
                        });
                    }
                }
                
                saveStateToLocalStorage();
                refreshLedgerCounts();
                switchTab('reconciled');
                showToast(`મેળવણી અહેવાલ સફળતાપૂર્વક આયાત કર્યો! ${reconciledTxnsList.length} મેળવાયેલા વ્યવહારો લોડ થયા.`, 'success');
                return;
            }
            
            // 1. Reset all previous auto-matches to start fresh
            state.mergedData.forEach(item => {
                if (item.matchGroupId) {
                    item.reconciled = false;
                    item.matchGroupId = null;
                }
            });
            
            // 2. Remove all old history items
            state.mergedData = state.mergedData.filter(item => 
                item.id && !item.id.startsWith('history_') && item.type !== 'history' && item.parentType !== 'history'
            );
            
            const pendingTxnsList = [];
            
            workbook.SheetNames.forEach(sheetName => {
                // Skip reconciled sheets
                if (sheetName.toLowerCase().includes('reconciled') || sheetName.includes('મેળવણી')) return;
                
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                rows.forEach(row => {
                    if (!row || row.length < 3) return;
                    
                    const dateStr = String(row[0] || '').trim();
                    const desc = String(row[1] || '').trim();
                    const type = String(row[2] || '').trim().toUpperCase();
                    
                    if (!dateStr || !desc || !type) return;
                    if (dateStr.toLowerCase().includes('date') || dateStr.toLowerCase().includes('statement') || dateStr.toLowerCase().includes('bank')) return;
                    if (desc.toLowerCase().includes('transaction') || desc.toLowerCase().includes('description') || desc.toLowerCase().includes('particular')) return;
                    if (desc.toLowerCase().includes('total unreconciled') || desc.toLowerCase().includes('net balance')) return;
                    if (dateStr === 'TOTAL' || desc === 'TOTAL') return;
                    
                    let parsedDate = formatDateToSlash(dateStr);
                    if (!isNaN(dateStr) && Number(dateStr) > 30000) {
                        const excelDate = new Date((Number(dateStr) - 25569) * 86400 * 1000);
                        parsedDate = formatDateToSlash(excelDate);
                    }
                    
                    const credit = parseFloat(String(row[6] || '').replace(/,/g, '')) || 0;
                    const debit = parseFloat(String(row[7] || '').replace(/,/g, '')) || 0;
                    
                    if (credit === 0 && debit === 0) return;
                    
                    const excelFlag = String(row[3] || 'DAILY').trim().toUpperCase();
                    let finalActualDate = formatDateToSlash(String(row[5] || parsedDate).trim());
                    if (excelFlag === 'NPCI') {
                        const descDate = extractDateFromDescription(desc);
                        if (descDate) {
                            finalActualDate = descDate;
                        }
                    }
                    
                    pendingTxnsList.push({
                        date: parsedDate,
                        description: desc,
                        type: type,
                        credit: credit,
                        debit: debit,
                        refNo: String(row[4] || '').trim(),
                        flag: excelFlag,
                        actualDate: finalActualDate
                    });
                });
            });
            
            console.log('[Recon Excel Import] Excel pending count:', pendingTxnsList.length);
            
            if (pendingTxnsList.length === 0) {
                showToast('અપલોડ કરેલ એક્સેલ ફાઇલમાં કોઈ માન્ય પેન્ડિંગ વ્યવહારો મળ્યા નથી!', 'error');
                return;
            }
            
            let matchedCount = 0;
            let importedCount = 0;
            const normalizeDesc = (str) => String(str || '').toLowerCase().replace(/[\s\-_,\.\/]/g, '').trim();
            const excelMatched = new Array(pendingTxnsList.length).fill(false);
            
            // Reconcile current transactions against the Excel pending list
            state.mergedData.forEach(item => {
                // If it's already reconciled, skip
                if (item.reconciled) return;
                
                const itemAmt = Math.max(item.creditTrn, item.debitTrn);
                const isItemCredit = item.creditTrn > 0;
                
                const matchIndex = pendingTxnsList.findIndex((excel, idx) => {
                    if (excelMatched[idx]) return false;
                    
                    const excelAmt = Math.max(excel.credit, excel.debit);
                    const isExcelCredit = excel.credit > 0;
                    
                    if (Math.abs(itemAmt - excelAmt) > 0.05) return false;
                    if (isItemCredit !== isExcelCredit) return false;
                    
                    const itemDescNorm = normalizeDesc(item.description);
                    const excelDescNorm = normalizeDesc(excel.description);
                    
                    if (itemDescNorm.includes(excelDescNorm) || excelDescNorm.includes(itemDescNorm) || getDescriptionSimilarity(item.description, excel.description) > 0.6) {
                        return true;
                    }
                    return false;
                });
                
                if (matchIndex !== -1) {
                    item.reconciled = false;
                    item.matchGroupId = null;
                    excelMatched[matchIndex] = true;
                    matchedCount++;
                }
            });
            
            // Add any Excel pending transactions that were NOT matched
            pendingTxnsList.forEach((excel, idx) => {
                if (excelMatched[idx]) return;
                
                state.mergedData.push({
                    id: 'history_' + idx,
                    date: excel.date,
                    description: excel.description,
                    type: excel.type,
                    actualDate: excel.actualDate || excel.date,
                    creditTrn: excel.credit,
                    debitTrn: excel.debit,
                    refNo: excel.refNo || '',
                    flag: excel.flag || 'DAILY',
                    day: 0,
                    count: 0,
                    reconciled: false,
                    matchGroupId: null
                });
                importedCount++;
            });
            
            saveStateToLocalStorage();
            refreshLedgerCounts();
            renderTable();
            updateBRSLiveWidget();
            
            showToast(`એક્સેલ મેળવણી પત્રક આયાત કર્યું! ${matchedCount} વ્યવહારો મેચ થયા અને ${importedCount} જૂના પેન્ડિંગ વ્યવહારો ઉમેરાયા.`, 'success');
        } catch (err) {
            console.error(err);
            showToast('એક્સેલ ફાઇલ રીડ કરવામાં ભૂલ આવી: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Automatically sync BRS state with Pending_Recon_20-07-2026.xlsx on startup if available (DISABLED)
function autoSyncWithExcel() {
    return; // Disabled per user request to only use uploaded files
    
    console.log('[Recon Startup] Checking for local Pending_Recon_20-07-2026.xlsx...');
    
    fetch('./Pending_Recon_20-07-2026.xlsx')
        .then(response => {
            if (!response.ok) throw new Error('Excel file not found locally');
            return response.arrayBuffer();
        })
        .then(arrayBuffer => {
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 1. Reset all previous auto-matches to start fresh
            state.mergedData.forEach(item => {
                if (item.matchGroupId) {
                    item.reconciled = false;
                    item.matchGroupId = null;
                }
            });
            
            // 2. Remove all old history items
            state.mergedData = state.mergedData.filter(item => 
                item.id && !item.id.startsWith('history_') && item.type !== 'history' && item.parentType !== 'history'
            );
            
            const pendingTxnsList = [];
            
            workbook.SheetNames.forEach(sheetName => {
                // Skip reconciled sheets
                if (sheetName.toLowerCase().includes('reconciled') || sheetName.includes('મેળવણી')) return;
                
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                rows.forEach(row => {
                    if (!row || row.length < 3) return;
                    
                    const dateStr = String(row[0] || '').trim();
                    const desc = String(row[1] || '').trim();
                    const type = String(row[2] || '').trim().toUpperCase();
                    
                    if (!dateStr || !desc || !type) return;
                    if (dateStr.toLowerCase().includes('date') || dateStr.toLowerCase().includes('statement') || dateStr.toLowerCase().includes('bank')) return;
                    if (desc.toLowerCase().includes('transaction') || desc.toLowerCase().includes('description') || desc.toLowerCase().includes('particular')) return;
                    if (desc.toLowerCase().includes('total unreconciled') || desc.toLowerCase().includes('net balance')) return;
                    if (dateStr === 'TOTAL' || desc === 'TOTAL') return;
                    
                    let parsedDate = formatDateToSlash(dateStr);
                    if (!isNaN(dateStr) && Number(dateStr) > 30000) {
                        const excelDate = new Date((Number(dateStr) - 25569) * 86400 * 1000);
                        parsedDate = formatDateToSlash(excelDate);
                    }
                    
                    const credit = parseFloat(String(row[6] || '').replace(/,/g, '')) || 0;
                    const debit = parseFloat(String(row[7] || '').replace(/,/g, '')) || 0;
                    
                    if (credit === 0 && debit === 0) return;
                    
                    pendingTxnsList.push({
                        date: parsedDate,
                        description: desc,
                        type: type,
                        credit: credit,
                        debit: debit,
                        refNo: String(row[4] || '').trim(),
                        flag: String(row[3] || 'DAILY').trim().toUpperCase(),
                        actualDate: formatDateToSlash(String(row[5] || parsedDate).trim())
                    });
                });
            });
            
            if (pendingTxnsList.length === 0) return;
            
            let matchedCount = 0;
            let importedCount = 0;
            const normalizeDesc = (str) => String(str || '').toLowerCase().replace(/[\s\-_,\.\/]/g, '').trim();
            const excelMatched = new Array(pendingTxnsList.length).fill(false);
            
            state.mergedData.forEach(item => {
                if (item.reconciled) return;
                
                const itemAmt = Math.max(item.creditTrn, item.debitTrn);
                const isItemCredit = item.creditTrn > 0;
                
                const matchIndex = pendingTxnsList.findIndex((excel, idx) => {
                    if (excelMatched[idx]) return false;
                    
                    const excelAmt = Math.max(excel.credit, excel.debit);
                    const isExcelCredit = excel.credit > 0;
                    
                    if (Math.abs(itemAmt - excelAmt) > 0.05) return false;
                    if (isItemCredit !== isExcelCredit) return false;
                    
                    const itemDescNorm = normalizeDesc(item.description);
                    const excelDescNorm = normalizeDesc(excel.description);
                    
                    if (itemDescNorm.includes(excelDescNorm) || excelDescNorm.includes(itemDescNorm) || getDescriptionSimilarity(item.description, excel.description) > 0.6) {
                        return true;
                    }
                    return false;
                });
                
                if (matchIndex !== -1) {
                    item.reconciled = false;
                    item.matchGroupId = null;
                    excelMatched[matchIndex] = true;
                    matchedCount++;
                }
            });
            
            pendingTxnsList.forEach((excel, idx) => {
                if (excelMatched[idx]) return;
                
                state.mergedData.push({
                    id: 'history_' + idx,
                    date: excel.date,
                    description: excel.description,
                    type: excel.type,
                    actualDate: excel.actualDate || excel.date,
                    creditTrn: excel.credit,
                    debitTrn: excel.debit,
                    flag: excel.flag || 'DAILY',
                    day: 0,
                    count: 0,
                    reconciled: false,
                    matchGroupId: null
                });
                importedCount++;
            });
            
            saveStateToLocalStorage();
            refreshLedgerCounts();
            renderTable();
            updateBRSLiveWidget();
            
            sessionStorage.setItem('recon_excel_synced', 'true');
            showToast(`દૈનિક મેળવણી પત્રક આપોઆપ મેળવાઈ ગયું! ${matchedCount} વ્યવહારો પેન્ડિંગ અને ${importedCount} જૂના પેન્ડિંગ વ્યવહારો ઉમેરાયા.`, 'success');
        })
        .catch(err => {
            console.log('[Recon Startup] Auto-sync skipped:', err.message);
        });
}

// Clear a specific uploaded file and its transactions from the ledger
function clearFile(fileType) {
    fileType = fileType.toLowerCase();
    
    if (fileType === 'history') {
        state.mergedData = state.mergedData.filter(item => !item.id.startsWith('history_'));
        state.files.history = { loaded: false, data: [], balance: 0, rawName: '' };
    } else {
        const typeStr = fileType.toUpperCase(); // e.g. HDFC, 345051, 3493, 3496
        state.mergedData = state.mergedData.filter(item => {
            const itemType = item.parentType || item.type;
            return itemType !== typeStr && !item.id.startsWith(fileType + '_');
        });
        state.files[fileType.toUpperCase()] = { loaded: false, data: [], balance: 0, rawName: '' };
    }
    
    // Save to local storage
    saveStateToLocalStorage();
    
    // Reset file input
    fileInput.value = '';
    
    // Update balance UI
    if (fileType !== 'history') {
        const balVal = document.getElementById(`val-${fileType}-bal`);
        if (balVal) {
            balVal.textContent = '₹ 0.00';
            balVal.classList.remove('negative-balance');
        }
    }
    
    // Update total balance UI
    let totalBal = 0;
    Object.values(state.files).forEach(f => {
        if (f.loaded) totalBal += f.balance;
    });
    valTotalBal.textContent = formatCurrency(totalBal);
    if (totalBal < 0) {
        valTotalBal.classList.add('negative-balance');
    } else {
        valTotalBal.classList.remove('negative-balance');
    }
    
    // Reset Card UI
    const card = document.getElementById(`status-${fileType}`);
    if (card) {
        card.classList.remove('loaded');
        card.querySelector('.file-indicator').className = 'file-indicator empty';
        card.querySelector('.file-meta').textContent = fileType === 'history' ? 'અપલોડ બાકી છે (વૈકલ્પિક)' : 'અપલોડ બાકી છે';
        
        // Hide delete icon
        const trashBtn = card.querySelector('.clear-file-btn');
        if (trashBtn) trashBtn.style.display = 'none';
    }
    
    // If no files are loaded, disable buttons
    let anyLoaded = Object.values(state.files).some(f => f.loaded);
    if (!anyLoaded) {
        btnAutoReconcile.disabled = true;
        btnExportPending.disabled = true;
        btnExportReconciled.disabled = true;
    }
    
    refreshLedgerCounts();
    renderTable();
    showToast(`${fileType === 'history' ? 'પેન્ડીંગ ઇતિહાસ' : fileType.toUpperCase() + '.CSV'} ફાઇલ સફળતાપૂર્વક દૂર કરવામાં આવી છે અને તમે નવી ફાઇલ અપલોડ કરી શકો છો.`, 'info');
}

// Global variables for RBI Report state
state.rbiData = null;
state.rbiReportingDateStr = '-';

function processRbiExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            let foundPendingSheet = false;
            let rows = [];
            
            // Look for pending or unreconciled sheet
            workbook.SheetNames.forEach(sheetName => {
                if (sheetName.toLowerCase().includes('pending') || sheetName.includes('બાકી')) {
                    const sheet = workbook.Sheets[sheetName];
                    rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                    foundPendingSheet = true;
                }
            });
            
            if (!foundPendingSheet && workbook.SheetNames.length > 0) {
                // fallback to first sheet
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            }
            
            if (rows.length === 0) {
                showToast('પેન્ડિંગ એક્સેલ ફાઇલ ખાલી છે અથવા યોગ્ય નથી!', 'error');
                return;
            }
            
            let reportingDate = new Date();
            let gl3493_bal = 0;
            let gl3496_bal = 0;
            const txns = [];
            
            rows.forEach(row => {
                if (!row) return;
                const cell1 = String(row[0] || '').trim();
                const cell2 = String(row[1] || '').trim();
                
                // Extract statement date
                if (cell1.toLowerCase().includes('statement date') || cell2.toLowerCase().includes('statement date')) {
                    row.forEach(val => {
                        const strVal = String(val || '').trim();
                        if (strVal) {
                            let targetDateStr = strVal;
                            if (strVal.toUpperCase().includes(' TO ')) {
                                const parts = strVal.toUpperCase().split(' TO ');
                                targetDateStr = parts[parts.length - 1].trim();
                            }
                            
                            const parsed = parseDateString(targetDateStr);
                            if (parsed && !isNaN(parsed.getTime())) {
                                reportingDate = parsed;
                            }
                        }
                    });
                }
                
                // Extract balances of 3493 and 3496
                if (cell2.toUpperCase().includes('GL 3493')) {
                    const cleanVal = String(row[6] || '0').replace(/[^\d\.-]/g, '').trim();
                    gl3493_bal = parseFloat(cleanVal) || 0;
                }
                if (cell2.toUpperCase().includes('GL 3496')) {
                    const cleanVal = String(row[6] || '0').replace(/[^\d\.-]/g, '').trim();
                    gl3496_bal = parseFloat(cleanVal) || 0;
                }
                
                // Parse transaction rows
                if (row.length >= 3) {
                    const dateStr = String(row[0] || '').trim();
                    const desc = String(row[1] || '').trim();
                    const type = String(row[2] || '').trim().toUpperCase();
                    
                    if (!dateStr || !desc || !type) return;
                    if (dateStr.toLowerCase().includes('date') || dateStr.toLowerCase().includes('statement') || dateStr.toLowerCase().includes('bank')) return;
                    if (desc.toLowerCase().includes('transaction') || desc.toLowerCase().includes('particular') || desc.toLowerCase().includes('description')) return;
                    if (desc.toLowerCase().includes('total unreconciled') || desc.toLowerCase().includes('net balance')) return;
                    if (dateStr === 'TOTAL' || desc === 'TOTAL') return;
                    
                    let parsedDate = parseDateString(dateStr);
                    if (!isNaN(dateStr) && Number(dateStr) > 30000) {
                        parsedDate = new Date((Number(dateStr) - 25569) * 86400 * 1000);
                    }
                    
                    if (!parsedDate || isNaN(parsedDate.getTime())) return;
                    
                    const cleanCredit = String(row[5] || '').replace(/[^\d\.-]/g, '').trim();
                    const credit = parseFloat(cleanCredit) || 0;
                    
                    const cleanDebit = String(row[6] || '').replace(/[^\d\.-]/g, '').trim();
                    const debit = parseFloat(cleanDebit) || 0;
                    
                    if (credit === 0 && debit === 0) return;
                    
                    txns.push({
                        date: parsedDate, // Use booking DATE column for ageing as requested
                        description: desc,
                        type: type,
                        credit: credit,
                        debit: debit
                    });
                }
            });
            
            // Format Reporting Date for display
            const dayPart = String(reportingDate.getDate()).padStart(2, '0');
            const monthPart = String(reportingDate.getMonth() + 1).padStart(2, '0');
            const yearPart = reportingDate.getFullYear();
            state.rbiReportingDateStr = `${dayPart}-${monthPart}-${yearPart}`;
            
            // Initialize Buckets
            const buckets = [
                { name: '1 TO 90 DAYS', creditCount: 0, creditAmt: 0, debitCount: 0, debitAmt: 0 },
                { name: '91 TO 180 DAYS', creditCount: 0, creditAmt: 0, debitCount: 0, debitAmt: 0 },
                { name: '181 TO 365 DAYS', creditCount: 0, creditAmt: 0, debitCount: 0, debitAmt: 0 },
                { name: '1 YEAR TO 2 YEAR', creditCount: 0, creditAmt: 0, debitCount: 0, debitAmt: 0 },
                { name: 'ABOVE 2 YEAR', creditCount: 0, creditAmt: 0, debitCount: 0, debitAmt: 0 }
            ];
            
            txns.forEach(t => {
                const daysDiff = Math.max(0, Math.floor((reportingDate - t.date) / (1000 * 60 * 60 * 24)));
                
                let bIdx = 4; // Above 2 years
                if (daysDiff <= 90) bIdx = 0;
                else if (daysDiff <= 180) bIdx = 1;
                else if (daysDiff <= 365) bIdx = 2;
                else if (daysDiff <= 730) bIdx = 3;
                
                if (t.credit > 0) {
                    buckets[bIdx].creditCount++;
                    buckets[bIdx].creditAmt += t.credit;
                } else if (t.debit > 0) {
                    buckets[bIdx].debitCount++;
                    buckets[bIdx].debitAmt += t.debit;
                }
            });
            
            // Add GL 3496 closing balance to 1 to 90 days Credit Amount
            buckets[0].creditAmt += gl3496_bal;
            
            // Add GL 3493 closing balance to 1 to 90 days Debit Amount
            buckets[0].debitAmt += gl3493_bal;
            
            state.rbiData = {
                reportingDateStr: state.rbiReportingDateStr,
                gl3493_bal: gl3493_bal,
                gl3496_bal: gl3496_bal,
                buckets: buckets
            };
            
            // Render table
            rbiStatementDateLabel.textContent = `અહેવાલ તારીખ (Reporting Date): ${state.rbiReportingDateStr} (GL 3493 Bal: ${formatCurrency(gl3493_bal)} | GL 3496 Bal: ${formatCurrency(gl3496_bal)})`;
            renderRbiTable();
            rbiResultArea.style.display = 'block';
            showToast('RBI ત્રિમાસિક રિપોર્ટ સફળતાપૂર્વક તૈયાર થયો છે!', 'success');
            
        } catch (err) {
            console.error(err);
            showToast('એક્સેલ ફાઇલ વાંચવામાં ભૂલ આવી: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderRbiTable() {
    if (!state.rbiData) return;
    
    let html = '';
    let totalCreditCount = 0;
    let totalCreditAmt = 0;
    let totalDebitCount = 0;
    let totalDebitAmt = 0;
    
    state.rbiData.buckets.forEach((b, idx) => {
        totalCreditCount += b.creditCount;
        totalCreditAmt += b.creditAmt;
        totalDebitCount += b.debitCount;
        totalDebitAmt += b.debitAmt;
        
        html += `
            <tr>
                <td style="text-align: center; font-weight: 600;">${idx + 1}</td>
                <td style="font-weight: 600;">${b.name}</td>
                <td style="text-align: right; font-weight: 600; color: #10b981;">${b.creditCount}</td>
                <td style="text-align: right; font-weight: 600; color: #10b981;">${formatCurrency(b.creditAmt)}</td>
                <td style="text-align: right; font-weight: 600; color: #ef4444;">${b.debitCount}</td>
                <td style="text-align: right; font-weight: 600; color: #ef4444;">${formatCurrency(b.debitAmt)}</td>
            </tr>
        `;
    });
    
    // Total row
    html += `
        <tr style="background: rgba(168, 85, 247, 0.08); border-top: 2px solid var(--border-color);">
            <td style="text-align: center; font-weight: 700;">-</td>
            <td style="font-weight: 700; color: #a855f7;">TOTAL</td>
            <td style="text-align: right; font-weight: 700; color: #10b981;">${totalCreditCount}</td>
            <td style="text-align: right; font-weight: 700; color: #10b981;">${formatCurrency(totalCreditAmt)}</td>
            <td style="text-align: right; font-weight: 700; color: #ef4444;">${totalDebitCount}</td>
            <td style="text-align: right; font-weight: 700; color: #ef4444;">${formatCurrency(totalDebitAmt)}</td>
        </tr>
    `;
    
    rbiTableBody.innerHTML = html;
}

function exportRbiReportToExcel() {
    if (!state.rbiData) return;
    
    // Create styles for clean presentation
    const titleStyle = { font: { name: 'Arial', size: 14, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "111827" } }, alignment: { horizontal: "center", vertical: "center" } };
    const headerStyle = { font: { name: 'Arial', size: 10, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "1F2937" } }, alignment: { horizontal: "center", vertical: "center" }, border: { top: { style: 'thin', color: { rgb: "D1D5DB" } }, bottom: { style: 'medium', color: { rgb: "111827" } }, left: { style: 'thin', color: { rgb: "D1D5DB" } }, right: { style: 'thin', color: { rgb: "D1D5DB" } } } };
    
    const dataStyleCenter = { font: { name: 'Arial', size: 10 }, alignment: { horizontal: "center", vertical: "center" }, border: { top: { style: 'thin', color: { rgb: "E5E7EB" } }, bottom: { style: 'thin', color: { rgb: "E5E7EB" } }, left: { style: 'thin', color: { rgb: "E5E7EB" } }, right: { style: 'thin', color: { rgb: "E5E7EB" } } } };
    const dataStyleLeft = { font: { name: 'Arial', size: 10, bold: true }, alignment: { horizontal: "left", vertical: "center" }, border: { top: { style: 'thin', color: { rgb: "E5E7EB" } }, bottom: { style: 'thin', color: { rgb: "E5E7EB" } }, left: { style: 'thin', color: { rgb: "E5E7EB" } }, right: { style: 'thin', color: { rgb: "E5E7EB" } } } };
    const dataStyleRight = { font: { name: 'Arial', size: 10 }, numFmt: '#,##0.00', alignment: { horizontal: "right", vertical: "center" }, border: { top: { style: 'thin', color: { rgb: "E5E7EB" } }, bottom: { style: 'thin', color: { rgb: "E5E7EB" } }, left: { style: 'thin', color: { rgb: "E5E7EB" } }, right: { style: 'thin', color: { rgb: "E5E7EB" } } } };
    
    const totalStyleLeft = { font: { name: 'Arial', size: 10, bold: true }, alignment: { horizontal: "left", vertical: "center" }, fill: { fgColor: { rgb: "F3F4F6" } }, border: { top: { style: 'medium', color: { rgb: "111827" } }, bottom: { style: 'double', color: { rgb: "111827" } }, left: { style: 'thin', color: { rgb: "D1D5DB" } }, right: { style: 'thin', color: { rgb: "D1D5DB" } } } };
    const totalStyleRight = { font: { name: 'Arial', size: 10, bold: true }, numFmt: '#,##0.00', alignment: { horizontal: "right", vertical: "center" }, fill: { fgColor: { rgb: "F3F4F6" } }, border: { top: { style: 'medium', color: { rgb: "111827" } }, bottom: { style: 'double', color: { rgb: "111827" } }, left: { style: 'thin', color: { rgb: "D1D5DB" } }, right: { style: 'thin', color: { rgb: "D1D5DB" } } } };
    const totalStyleCenter = { font: { name: 'Arial', size: 10, bold: true }, alignment: { horizontal: "center", vertical: "center" }, fill: { fgColor: { rgb: "F3F4F6" } }, border: { top: { style: 'medium', color: { rgb: "111827" } }, bottom: { style: 'double', color: { rgb: "111827" } }, left: { style: 'thin', color: { rgb: "D1D5DB" } }, right: { style: 'thin', color: { rgb: "D1D5DB" } } } };
    
    const c = (v, s) => ({ v: v, t: typeof v === 'number' ? 'n' : 's', s: s });
    
    const sheetData = [
        [c('THE JUNAGADH COMMERCIAL CO-OPERATIVE BANK LTD.', titleStyle), '', '', '', '', ''],
        [c(`RBI PENDING AGEING REPORT (As of ${state.rbiData.reportingDateStr})`, titleStyle), '', '', '', '', ''],
        [c(`GL 3493 Balance: ${state.rbiData.gl3493_bal.toFixed(2)}  |  GL 3496 Balance: ${state.rbiData.gl3496_bal.toFixed(2)}`, { font: { name: 'Arial', size: 10, italic: true }, alignment: { horizontal: "center" } }), '', '', '', '', ''],
        [],
        [
            c('SR NO', headerStyle),
            c('AGEING CATEGORY', headerStyle),
            c('CREDIT RECORD', headerStyle),
            c('CREDIT AMOUNT', headerStyle),
            c('DEBIT RECORD', headerStyle),
            c('DEBIT AMOUNT', headerStyle)
        ]
    ];
    
    let totalCreditCount = 0;
    let totalCreditAmt = 0;
    let totalDebitCount = 0;
    let totalDebitAmt = 0;
    
    state.rbiData.buckets.forEach((b, idx) => {
        totalCreditCount += b.creditCount;
        totalCreditAmt += b.creditAmt;
        totalDebitCount += b.debitCount;
        totalDebitAmt += b.debitAmt;
        
        sheetData.push([
            c(idx + 1, dataStyleCenter),
            c(b.name, dataStyleLeft),
            c(b.creditCount, dataStyleCenter),
            c(b.creditAmt, dataStyleRight),
            c(b.debitCount, dataStyleCenter),
            c(b.debitAmt, dataStyleRight)
        ]);
    });
    
    // Add total row
    sheetData.push([
        c('', totalStyleCenter),
        c('TOTAL', totalStyleLeft),
        c(totalCreditCount, totalStyleCenter),
        c(totalCreditAmt, totalStyleRight),
        c(totalDebitCount, totalStyleCenter),
        c(totalDebitAmt, totalStyleRight)
    ]);
    
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    
    // Merges
    worksheet['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 5 } }
    ];
    
    worksheet['!cols'] = [
        { wch: 10 }, // Sr No
        { wch: 25 }, // Ageing Category
        { wch: 18 }, // Credit Record
        { wch: 22 }, // Credit Amount
        { wch: 18 }, // Debit Record
        { wch: 22 }  // Debit Amount
    ];
    
    worksheet['!pageSetup'] = {
        paperSize: 9, // A4
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0
    };
    worksheet['!margins'] = {
        left: 1.0,  // 1.0 inch
        right: 0.5, // 0.5 inch
        top: 0.5,   // 0.5 inch
        bottom: 0.5, // 0.5 inch
        header: 0.3,
        footer: 0.3
    };
    worksheet['!views'] = [{ showGridLines: true }];
    
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'RBI Ageing Report');
    
    const filename = `RBI_Quarterly_Ageing_${state.rbiData.reportingDateStr}.xlsx`;
    XLSX.writeFile(workbook, filename);
    showToast(`RBI રીપોર્ટ ${filename} ડાઉનલોડ થઈ ગયો!`, 'success');
}
