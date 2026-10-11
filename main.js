import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs/dist/exceljs.min.js';
import { sysgenInitialLoanData } from './sysgen_rpo_loans_data.js';

// Application State
const state = {
  lang: 'en', // 'en' | 'tl'
  activeTheme: 'theme-wood',
  currentView: 'excel', // 'excel' | 'onlyLoan' | 'noLoan' | 'overall' | 'compute'
  loanFilterMode: 'formal', // 'formal' (24 active institutional loans) | 'all' (123 all deductions)
  records: [], // Normalized objects
  headers: [], // Raw Column keys
  searchQuery: '',
  currentPage: 1,
  pageSize: 50,
  selectedIndices: new Set(),
  inspectingRecord: null
};

// Bilingual Translations Dictionary (English / Tagalog)
const i18n = {
  en: {
    splashGreeting: 'Welcome, <span class="highlight-gold">Ivy Salamanca</span>',
    splashRole: 'Authorized Personnel | Loans & Payroll Computation Division',
    splashLangLabel: 'Language / Wika:',
    warningTitle: 'CONFIDENTIALITY & USAGE NOTICE',
    warningDesc: 'Strictly for Sysgen RPO internal loan computations only. All formulas, deductions, and employee payroll records are proprietary and confidential.',
    btnLogin: 'Login & Open Workspace',
    creditLabel: 'Engineered & Developed by',
    badgeEngine: 'Loans Engine',
    headerSubtitle: 'Payroll Loan Computations & Excel Intelligence Suite',
    langLabel: 'Language:',
    themeLabel: 'Theme:',
    navExcel: 'All Records (Excel)',
    navOnlyLoan: 'Only Loans',
    navNoLoan: 'All Without Loans',
    navOverall: 'Overall Split View',
    navOverallSub: 'Top / Bottom',
    navCompute: 'Compute & Analytics',
    navComputeSub: 'Summary',
    loanFilterLabel: 'Category:',
    filterFormalLoans: 'Institutional Loans Only (24)',
    filterAllDeductions: 'All Payroll Deductions (123)',
    btnUpload: 'Upload Excel (.xlsx)',
    btnDownloadOnlyLoans: 'Download Only Loans (.xlsx)',
    btnDownloadAll: 'Download All (.xlsx)',
    searchPlaceholder: 'Search employee name, account, or loan amount...',
    showingAll: 'All Payroll Records',
    showingOnlyLoans: 'Active Institutional Loans Only (Zero Loans Excluded)',
    showingOnlyLoansAll: 'All Deductions & Charges (Zero Loans Excluded)',
    showingNoLoans: 'Employees Without Active Loans',
    showingOverall: 'Overall Split View (No Loans on Top, Active Loans on Bottom)',
    showingCompute: 'Compute & Loan Analytics Dashboard',
    pageSizeLabel: 'Show:',
    btnCopyExport: 'Copy / Export',
    exportJson: 'Copy Visible as JSON',
    exportTsv: 'Copy as TSV (Paste to Excel/Sheets)',
    exportCsv: 'Copy as CSV',
    exportMd: 'Copy as Markdown Table',
    copySelected: 'Copy Selected',
    selected: 'selected',
    splitTopBadge: 'TOP SECTION',
    splitTopTitle: 'No Active Loans / Deductions',
    splitBottomBadge: 'BOTTOM SECTION',
    splitBottomTitle: 'Active Loans & Deductions',
    kpiTotalDeductions: 'Total Cutoff Loan Deductions',
    kpiHeadcount: 'Total Workforce Headcount',
    kpiLoanPercentage: 'with active loan deductions',
    kpiHighestCat: 'Highest Loan Category',
    kpiAvgDeduction: 'Average Deduction Per Borrower',
    kpiAvgDeductionSub: 'Per payroll period',
    breakdownTitle: 'Breakdown by Loan & Contribution Category',
    calcTitle: 'Interactive Loan Amortization Calculator for Ivy Salamanca',
    calcBadge: 'Live Simulator',
    calcLoanType: 'Loan Type',
    calcPrincipal: 'Principal Amount (PHP)',
    calcTerm: 'Term (Months)',
    calcRate: 'Annual Interest Rate (%)',
    calcMonthlyLabel: 'Monthly Deduction',
    calcSemiMonthlyLabel: 'Semi-Monthly Cutoff (15/30)',
    calcInterestLabel: 'Total Interest',
    calcPayableLabel: 'Total Payable Amount',
    topBorrowersTitle: 'Top Loan Deductions in Current Payroll',
    topBorrowersBadge: 'Ranked',
    colRow: '#',
    colAction: 'Action',
    colStatus: 'Loan Status',
    colLastName: 'Last Name',
    colFirstName: 'First Name',
    colPayrollDate: 'Payroll Date',
    colActiveLoans: 'Active Loans List',
    colTotalLoanDeduction: 'TOTAL LOAN DEDUCTION',
    grandTotal: 'GRAND TOTAL',
    borrowersCount: 'borrowers',
    noActiveDeductions: 'No active deductions',
    scrollNavTitle: 'Quick Scroll:',
    btnScrollLeft: 'Left',
    btnScrollRight: 'Right',
    jumpNames: 'Names',
    jumpLoanList: 'Loan List',
    jumpInstitutions: 'SSS / HDMF / AUB',
    jumpTotal: 'Total Deduction',
    scrollTip: 'Tip: Shift + Mouse Wheel or use Left/Right buttons',
    emptyTitle: 'No Payroll File Loaded',
    emptyDesc: 'Drop your Excel (.xlsx) file here or click below to upload. The system will automatically compute all loan deductions, SSS, HDMF, AUB, and rankings.',
    btnEmptyUpload: 'Upload Excel (.xlsx) File',
    btnDemoData: 'Or Load Sample Demo Data',
    emptyHint: 'Supports standard Sysgen RPO cutoffs and custom payroll loan spreadsheets.',
    btnInstallApp: 'Install App'
  },
  tl: {
    splashGreeting: 'Maligayang Pagdating, <span class="highlight-gold">Ivy Salamanca</span>',
    splashRole: 'Awtorisadong Tauhan | Dibisyon ng Komputasyon ng Loan at Payroll',
    splashLangLabel: 'Wika / Language:',
    warningTitle: 'BABALA SA PAGKUMPIDENSIYAL AT PAGGAMIT',
    warningDesc: 'Para lamang sa internal na komputasyon ng Sysgen RPO. Lahat ng formula, bawas, at talaan ng empleyado ay pagmamay-ari at kumpidensiyal.',
    btnLogin: 'Mag-login at Buksan ang Workspace',
    creditLabel: 'Idinisenyo at Binuo ni',
    badgeEngine: 'Makina ng Loan',
    headerSubtitle: 'Komputasyon ng Loan sa Payroll at Sistema ng Excel',
    langLabel: 'Wika:',
    themeLabel: 'Tema:',
    navExcel: 'Lahat ng Talaan (Excel)',
    navOnlyLoan: 'May Loan Lamang',
    navNoLoan: 'Walang Loan',
    navOverall: 'Pinaghating Tanawin',
    navOverallSub: 'Taas / Ibaba',
    navCompute: 'Komputasyon at Analytics',
    navComputeSub: 'Buod',
    loanFilterLabel: 'Kategorya:',
    filterFormalLoans: 'May Active Loan Lamang (24)',
    filterAllDeductions: 'Lahat ng Bawas (123)',
    btnUpload: 'Mag-upload ng Excel (.xlsx)',
    btnDownloadOnlyLoans: 'I-download ang May Loan Lamang (.xlsx)',
    btnDownloadAll: 'I-download ang Lahat (.xlsx)',
    searchPlaceholder: 'Mag-search ng apelyido, pangalan, o halaga ng loan...',
    showingAll: 'Lahat ng Payroll Records',
    showingOnlyLoans: 'May Active Loan Lamang (Hindi Kasama ang Walang Loan)',
    showingOnlyLoansAll: 'Lahat ng Bawas at Singil (Hindi Kasama ang Walang Loan)',
    showingNoLoans: 'Mga Empleyadong Walang Active Loan',
    showingOverall: 'Pinaghating Tanawin (Walang Loan sa Taas, May Loan sa Ibaba)',
    showingCompute: 'Dashboard ng Komputasyon at Loan Analytics',
    pageSizeLabel: 'Ipakita:',
    btnCopyExport: 'Kopyahin / I-export',
    exportJson: 'Kopyahin ang Nakikita bilang JSON',
    exportTsv: 'Kopyahin bilang TSV (I-paste sa Excel/Sheets)',
    exportCsv: 'Kopyahin bilang CSV',
    exportMd: 'Kopyahin bilang Markdown Table',
    copySelected: 'Kopyahin ang Napili',
    selected: 'napili',
    splitTopBadge: 'TAAS NA BAHAGI',
    splitTopTitle: 'Walang Active Loan / Bawas',
    splitBottomBadge: 'IBABANG BAHAGI',
    splitBottomTitle: 'May Active Loan at Bawas',
    kpiTotalDeductions: 'Kabuuang Bawas sa Loan ngayong Cutoff',
    kpiHeadcount: 'Kabuuang Bilang ng Empleyado',
    kpiLoanPercentage: 'ang may bawas sa payroll',
    kpiHighestCat: 'Pinakamataas na Kategorya ng Loan',
    kpiAvgDeduction: 'Average na Bawas Bawat Borrower',
    kpiAvgDeductionSub: 'Bawat panahon ng payroll',
    breakdownTitle: 'Breakdown Ayon sa Uri ng Loan at Kontribusyon',
    calcTitle: 'Interactive Loan Calculator para kay Ivy Salamanca',
    calcBadge: 'Live Simulator',
    calcLoanType: 'Uri ng Loan',
    calcPrincipal: 'Halaga ng Principal (PHP)',
    calcTerm: 'Term (Buwan)',
    calcRate: 'Taunang Interes (%)',
    calcMonthlyLabel: 'Buwanang Bawas',
    calcSemiMonthlyLabel: 'Kada Cutoff (15/30)',
    calcInterestLabel: 'Kabuuang Interes',
    calcPayableLabel: 'Kabuuang Babayaran',
    topBorrowersTitle: 'Nangungunang Bawas sa Loan sa Kasalukuyang Payroll',
    topBorrowersBadge: 'Niraranggo',
    colRow: '#',
    colAction: 'Aksyon',
    colStatus: 'Katayuan ng Loan',
    colLastName: 'Apelyido',
    colFirstName: 'Pangalan',
    colPayrollDate: 'Petsa ng Payroll',
    colActiveLoans: 'Listahan ng Active Loans',
    colTotalLoanDeduction: 'KABUUANG BAWAS SA LOAN',
    grandTotal: 'KABUUANG TOTAL',
    borrowersCount: 'may loan',
    noActiveDeductions: 'Walang active deductions',
    scrollNavTitle: 'Mabilis na Pag-scroll:',
    btnScrollLeft: 'Kaliwa',
    btnScrollRight: 'Kanan',
    jumpNames: 'Pangalan',
    jumpLoanList: 'Listahan ng Loan',
    jumpInstitutions: 'SSS / HDMF / AUB',
    jumpTotal: 'Kabuuang Bawas',
    scrollTip: 'Tip: Shift + Mouse Wheel o gamitin ang Kaliwa/Kanan',
    emptyTitle: 'Walang Payroll File na Naka-load',
    emptyDesc: 'I-drop ang iyong Excel (.xlsx) file dito o mag-click sa ibaba upang mag-upload. Awtomatikong kukuwentahin ang lahat ng bawas sa loan, SSS, HDMF, AUB, at ranggo.',
    btnEmptyUpload: 'Mag-upload ng Excel (.xlsx) File',
    btnDemoData: 'O Subukan ang Sample Demo Data',
    emptyHint: 'Sumusuporta sa karaniwang Sysgen RPO cutoff at iba pang payroll loan spreadsheet.',
    btnInstallApp: 'I-install ang App'
  }
};

// DOM References
const splashOverlay = document.getElementById('splashOverlay');
const splashLogo = document.getElementById('splashLogo');
const btnLogin = document.getElementById('btnLogin');
const btnLogout = document.getElementById('btnLogout');
const app = document.getElementById('app');
const headerBrandLogo = document.getElementById('headerBrandLogo');

// Empty State & App Install
const emptyDropzoneState = document.getElementById('emptyDropzoneState');
const btnEmptyUpload = document.getElementById('btnEmptyUpload');
const btnLoadDemoData = document.getElementById('btnLoadDemoData');
const btnInstallApp = document.getElementById('btnInstallApp');
const tableScrollHelper = document.getElementById('tableScrollHelper');
const mainPaginationFooter = document.getElementById('mainPaginationFooter');

// Language Buttons
const btnLangEn = document.getElementById('btnLangEn');
const btnLangTl = document.getElementById('btnLangTl');
const btnSplashLangEn = document.getElementById('btnSplashLangEn');
const btnSplashLangTl = document.getElementById('btnSplashLangTl');

// Only Loans Mode Toggle Elements
const onlyLoanModeToggle = document.getElementById('onlyLoanModeToggle');
const btnFilterFormalLoans = document.getElementById('btnFilterFormalLoans');
const btnFilterAllDeductions = document.getElementById('btnFilterAllDeductions');

// View Panels
const viewStandardTable = document.getElementById('viewStandardTable');
const viewOverallSplit = document.getElementById('viewOverallSplit');
const viewCompute = document.getElementById('viewCompute');

// Navigation Buttons
const navBtnExcel = document.getElementById('navBtnExcel');
const navBtnOnlyLoan = document.getElementById('navBtnOnlyLoan');
const navBtnNoLoan = document.getElementById('navBtnNoLoan');
const navBtnOverall = document.getElementById('navBtnOverall');
const navBtnCompute = document.getElementById('navBtnCompute');
const navButtons = [navBtnExcel, navBtnOnlyLoan, navBtnNoLoan, navBtnOverall, navBtnCompute];

// Counters & Tags
const countAllTotal = document.getElementById('countAllTotal');
const countOnlyLoan = document.getElementById('countOnlyLoan');
const countNoLoan = document.getElementById('countNoLoan');
const viewStatusTag = document.getElementById('viewStatusTag');

// Search & Controls
const globalSearchInput = document.getElementById('globalSearchInput');
const btnResetSearch = document.getElementById('btnResetSearch');
const controlsToolbar = document.getElementById('controlsToolbar');
const pageSizeWrap = document.getElementById('pageSizeWrap');
const pageSizeSelect = document.getElementById('pageSizeSelect');

// Batch Selection
const selectionBatchBox = document.getElementById('selectionBatchBox');
const selectedRowsCount = document.getElementById('selectedRowsCount');
const btnCopySelectedBatch = document.getElementById('btnCopySelectedBatch');

// Export Menu
const btnExportMenu = document.getElementById('btnExportMenu');
const exportDropdownMenu = document.getElementById('exportDropdownMenu');
const btnDownloadOnlyLoansXlsx = document.getElementById('btnDownloadOnlyLoansXlsx');
const btnExportAllXlsx = document.getElementById('btnExportAllXlsx');
const fileInputGlobal = document.getElementById('fileInputGlobal');

// Main Grid & Scroll Helper
const mainTableScrollContainer = document.getElementById('mainTableScrollContainer');
const btnScrollStepLeft = document.getElementById('btnScrollStepLeft');
const btnScrollStepRight = document.getElementById('btnScrollStepRight');
const btnJumpToNames = document.getElementById('btnJumpToNames');
const btnJumpToActiveList = document.getElementById('btnJumpToActiveList');
const btnJumpToInstitutions = document.getElementById('btnJumpToInstitutions');
const btnJumpToTotal = document.getElementById('btnJumpToTotal');
const scrollMiniThumb = document.getElementById('scrollMiniThumb');
const scrollMiniIndicator = document.querySelector('.scroll-mini-indicator');
const txtScrollNavTitle = document.getElementById('txtScrollNavTitle');
const txtBtnScrollLeft = document.getElementById('txtBtnScrollLeft');
const txtBtnScrollRight = document.getElementById('txtBtnScrollRight');
const txtScrollTipLabel = document.getElementById('txtScrollTipLabel');
const mainGridHead = document.getElementById('mainGridHead');
const mainGridBody = document.getElementById('mainGridBody');
const mainGridFoot = document.getElementById('mainGridFoot');
const mainPaginationInfo = document.getElementById('mainPaginationInfo');
const mainPageCounter = document.getElementById('mainPageCounter');
const btnPrev = document.getElementById('btnPrev');
const btnNext = document.getElementById('btnNext');

// Overall Split View Elements
const splitNoLoanCount = document.getElementById('splitNoLoanCount');
const splitNoLoanBody = document.getElementById('splitNoLoanBody');
const splitWithLoanCount = document.getElementById('splitWithLoanCount');
const splitWithLoanBody = document.getElementById('splitWithLoanBody');
const splitWithLoanFoot = document.getElementById('splitWithLoanFoot');
const btnCopyNoLoanTsv = document.getElementById('btnCopyNoLoanTsv');
const btnCopyNoLoanJson = document.getElementById('btnCopyNoLoanJson');
const btnCopyWithLoanTsv = document.getElementById('btnCopyWithLoanTsv');
const btnCopyWithLoanJson = document.getElementById('btnCopyWithLoanJson');

// Compute Dashboard Elements
const computeTotalDeductions = document.getElementById('computeTotalDeductions');
const computeActiveBorrowersCount = document.getElementById('computeActiveBorrowersCount');
const computeTotalHeadcount = document.getElementById('computeTotalHeadcount');
const computeLoanPercentage = document.getElementById('computeLoanPercentage');
const computeHighestCategory = document.getElementById('computeHighestCategory');
const computeHighestCategorySum = document.getElementById('computeHighestCategorySum');
const computeAvgDeduction = document.getElementById('computeAvgDeduction');
const categorySummaryGrid = document.getElementById('categorySummaryGrid');
const topBorrowersList = document.getElementById('topBorrowersList');
const btnCopyComputeBreakdown = document.getElementById('btnCopyComputeBreakdown');

// Calculator Elements
const calcLoanType = document.getElementById('calcLoanType');
const calcPrincipal = document.getElementById('calcPrincipal');
const calcMonths = document.getElementById('calcMonths');
const calcInterestRate = document.getElementById('calcInterestRate');
const calcMonthlyDeduction = document.getElementById('calcMonthlyDeduction');
const calcSemiMonthlyDeduction = document.getElementById('calcSemiMonthlyDeduction');
const calcTotalInterest = document.getElementById('calcTotalInterest');
const calcTotalPayable = document.getElementById('calcTotalPayable');
const btnCopyCalcResult = document.getElementById('btnCopyCalcResult');

// Modal Elements
const rowModal = document.getElementById('rowModal');
const modalRowBadge = document.getElementById('modalRowBadge');
const modalRowTitle = document.getElementById('modalRowTitle');
const modalRowContent = document.getElementById('modalRowContent');
const btnModalCopyAll = document.getElementById('btnModalCopyAll');
const btnModalCopyTsv = document.getElementById('btnModalCopyTsv');
const btnModalClose = document.getElementById('btnModalClose');
const btnModalDone = document.getElementById('btnModalDone');

// Toast Container
const toastContainer = document.getElementById('toastContainer');

// Utility: Format Currency in Philippine Pesos (PHP)
function formatPeso(val) {
  const num = Math.abs(parseFloat(val) || 0);
  return 'PHP ' + num.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Utility: Show Clean Toast (No Emojis)
function showToast(msg) {
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerHTML = `
    <svg viewBox="0 0 24 24" width="18" height="18" stroke="currentColor" stroke-width="2.5" fill="none" style="color: var(--gold-primary); flex-shrink: 0;"><polyline points="20 6 9 17 4 12"></polyline></svg>
    <span>${msg}</span>
  `;
  toastContainer.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.25s ease';
    setTimeout(() => toast.remove(), 250);
  }, 2600);
}

// Utility: Copy to Clipboard
async function copyToClipboard(text, msg) {
  try {
    await navigator.clipboard.writeText(text);
    showToast(msg || (state.lang === 'tl' ? 'Kinopya sa clipboard!' : 'Copied to clipboard!'));
  } catch (err) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    showToast(msg || (state.lang === 'tl' ? 'Kinopya sa clipboard!' : 'Copied to clipboard!'));
  }
}

// -----------------------------------------------------------------------------
// NORMALIZE EMPLOYEE & LOAN RECORDS
// -----------------------------------------------------------------------------
function normalizeRecord(raw, index) {
  const keys = Object.keys(raw || {});

  // Helper to find key by pattern variations
  const findVal = (patterns) => {
    for (const p of patterns) {
      const matchKey = keys.find(k => k.trim().toLowerCase() === p.toLowerCase());
      if (matchKey && raw[matchKey] !== undefined && raw[matchKey] !== null) {
        return { key: matchKey, val: raw[matchKey] };
      }
    }
    return null;
  };

  // 1. Employee Name detection
  const lnMatch = findVal(['Last Name', 'LastName', 'Last_Name', 'Surname', 'Family Name', 'Apelyido']);
  const fnMatch = findVal(['First Name', 'FirstName', 'First_Name', 'Given Name', 'Pangalan']);

  let lastName = lnMatch ? String(lnMatch.val).trim() : '';
  let firstName = fnMatch ? String(fnMatch.val).trim() : '';

  if (!lastName && !firstName) {
    const fullMatch = findVal(['Employee Name', 'Employee', 'Full Name', 'Name', 'Emp Name', 'Pangalan ng Empleyado']);
    if (fullMatch && fullMatch.val) {
      const nameStr = String(fullMatch.val).trim();
      if (nameStr.includes(',')) {
        const parts = nameStr.split(',');
        lastName = parts[0].trim();
        firstName = parts.slice(1).join(' ').trim();
      } else if (nameStr.includes(' ')) {
        const parts = nameStr.split(' ');
        lastName = parts[parts.length - 1].trim();
        firstName = parts.slice(0, -1).join(' ').trim();
      } else {
        lastName = nameStr;
      }
    }
  }

  // 2. Payroll Date detection
  const dateMatch = findVal(['Payroll Date', 'Date', 'Payroll Period', 'Period', 'Cutoff', 'Pay Date', 'Petsa']);
  let payrollDate = '2026-09-30';
  if (dateMatch && dateMatch.val) {
    if (dateMatch.val instanceof Date) {
      payrollDate = dateMatch.val.toISOString().split('T')[0];
    } else {
      payrollDate = String(dateMatch.val).split('T')[0];
    }
  }

  // 3. Known Standard Sysgen Fields
  let sssSalary = Math.abs(Number(raw['SSS - Salary Loan']) || 0);
  let sssCalamity = Math.abs(Number(raw['SSS - Calamity Loan']) || 0);
  let hdmfLoan = Math.abs(Number(raw['HDMF Loan']) || 0);
  let hdmfMp2 = Math.abs(Number(raw['HDMF - MP2']) || 0);
  let hdmfCalamity = Math.abs(Number(raw['HDMF - Calamity Loan']) || 0);
  let hdmfAddEE = Math.abs(Number(raw['HDMF - additional contribution EE']) || 0);
  let cashAdvances = Math.abs(Number(raw['Cash Advances']) || 0);
  let aubLoan = Math.abs(Number(raw['AUB Loan']) || 0);
  let hmoShare = Math.abs(Number(raw['HMO - Share']) || 0);
  let sssEmergency = Math.abs(Number(raw['SSS - Emergency Loan']) || 0);
  let otherCharges = Math.abs(Number(raw['Other charges and deduction']) || 0);
  let adjContrib = Math.abs(Number(raw['Adjustment on contributions']) || 0);

  let formalLoansTotal = sssSalary + sssCalamity + hdmfLoan + hdmfMp2 + hdmfCalamity + aubLoan + cashAdvances + sssEmergency;
  let totalAllDeductions = formalLoansTotal + otherCharges + hdmfAddEE + hmoShare + adjContrib;
  const activeLoanItems = [];

  // Known field keys list for exclusion from arbitrary scan
  const knownFields = [
    'SSS - Salary Loan', 'SSS - Calamity Loan', 'HDMF Loan', 'HDMF - MP2',
    'HDMF - Calamity Loan', 'HDMF - additional contribution EE', 'Cash Advances',
    'AUB Loan', 'HMO - Share', 'SSS - Emergency Loan', 'Other charges and deduction',
    'Adjustment on contributions'
  ];

  // 4. Dynamic Column Scan for Any Uploaded Custom Excel Layouts
  keys.forEach(k => {
    const kTrim = k.trim();
    const kLower = kTrim.toLowerCase();

    // Skip metadata columns
    if (
      kLower.includes('last name') || kLower.includes('first name') ||
      kLower.includes('employee') || kLower.includes('surname') ||
      kLower.includes('payroll date') || kLower === 'date' ||
      kLower.includes('column p') || kLower === 'id' || kLower.includes('emp id') ||
      kLower === '_index' || kLower === 'status' || kLower === 'remarks'
    ) {
      return;
    }

    if (knownFields.includes(kTrim)) return;

    const val = raw[k];
    const num = Number(val);
    if (val !== undefined && val !== null && !isNaN(num) && val !== '' && typeof val !== 'boolean') {
      const absVal = Math.abs(num);
      if (absVal > 0) {
        const isFormal = kLower.includes('loan') || kLower.includes('sss') || kLower.includes('hdmf') ||
                         kLower.includes('pag-ibig') || kLower.includes('pagibig') || kLower.includes('aub') ||
                         kLower.includes('calamity') || kLower.includes('advance') || kLower.includes('emergency');

        totalAllDeductions += absVal;
        if (isFormal) formalLoansTotal += absVal;

        activeLoanItems.push({ name: kTrim, amount: absVal, isFormal });
      }
    }
  });

  // Populate known active loan items
  if (aubLoan > 0 && !activeLoanItems.some(i => i.name === 'AUB Loan')) activeLoanItems.push({ name: 'AUB Loan', amount: aubLoan, isFormal: true });
  if (sssSalary > 0 && !activeLoanItems.some(i => i.name === 'SSS Salary Loan')) activeLoanItems.push({ name: 'SSS Salary Loan', amount: sssSalary, isFormal: true });
  if (hdmfLoan > 0 && !activeLoanItems.some(i => i.name === 'HDMF Loan')) activeLoanItems.push({ name: 'HDMF Loan', amount: hdmfLoan, isFormal: true });
  if (hdmfMp2 > 0 && !activeLoanItems.some(i => i.name === 'HDMF MP2')) activeLoanItems.push({ name: 'HDMF MP2', amount: hdmfMp2, isFormal: true });
  if (sssCalamity > 0 && !activeLoanItems.some(i => i.name === 'SSS Calamity')) activeLoanItems.push({ name: 'SSS Calamity', amount: sssCalamity, isFormal: true });
  if (hdmfCalamity > 0 && !activeLoanItems.some(i => i.name === 'HDMF Calamity')) activeLoanItems.push({ name: 'HDMF Calamity', amount: hdmfCalamity, isFormal: true });
  if (cashAdvances > 0 && !activeLoanItems.some(i => i.name === 'Cash Advance')) activeLoanItems.push({ name: 'Cash Advance', amount: cashAdvances, isFormal: true });
  if (sssEmergency > 0 && !activeLoanItems.some(i => i.name === 'SSS Emergency')) activeLoanItems.push({ name: 'SSS Emergency', amount: sssEmergency, isFormal: true });
  if (otherCharges > 0 && !activeLoanItems.some(i => i.name === 'Other Deductions')) activeLoanItems.push({ name: 'Other Deductions', amount: otherCharges, isFormal: false });
  if (adjContrib > 0 && !activeLoanItems.some(i => i.name === 'Adj on Contrib')) activeLoanItems.push({ name: 'Adj on Contrib', amount: adjContrib, isFormal: false });

  const hasLoan = formalLoansTotal > 0 || otherCharges > 0 || totalAllDeductions > 0;
  const hasFormalLoan = formalLoansTotal > 0;

  const activeLoansSummaryText = activeLoanItems.length > 0
    ? activeLoanItems.map(item => `${item.name}: ${formatPeso(item.amount)}`).join(' | ')
    : 'None';

  return {
    _index: index,
    lastName: lastName || String(raw['Last Name'] || '').trim(),
    firstName: firstName || String(raw['First Name'] || '').trim(),
    fullName: `${lastName || raw['Last Name'] || ''}, ${firstName || raw['First Name'] || ''}`.replace(/^,\s*|\s*,$/g, '').trim(),
    payrollDate,
    sssSalary,
    sssCalamity,
    hdmfLoan,
    hdmfMp2,
    hdmfCalamity,
    hdmfAddEE,
    cashAdvances,
    aubLoan,
    hmoShare,
    sssEmergency,
    otherCharges,
    adjContrib,
    formalLoansTotal,
    totalAllDeductions,
    hasLoan,
    hasFormalLoan,
    activeLoanItems,
    activeLoansSummaryText,
    raw: { ...raw }
  };
}

function initDataset(dataArray) {
  state.records = (dataArray || []).map((r, i) => normalizeRecord(r, i));
  if (dataArray && dataArray.length > 0) {
    state.headers = Object.keys(dataArray[0]);
  } else {
    state.headers = [];
  }
  state.selectedIndices.clear();
  state.currentPage = 1;
  updateGlobalCounters();
  renderCurrentView();
  computeDashboardMetrics();
}

function updateGlobalCounters() {
  const total = state.records.length;
  const withFormalLoan = state.records.filter(r => r.hasFormalLoan).length;
  const withAnyLoan = state.records.filter(r => r.hasLoan).length;
  const activeOnlyLoanCount = state.loanFilterMode === 'formal' ? withFormalLoan : withAnyLoan;
  const noLoan = total - withFormalLoan;

  countAllTotal.textContent = total;
  countOnlyLoan.textContent = activeOnlyLoanCount;
  countNoLoan.textContent = noLoan;
  splitNoLoanCount.textContent = noLoan;
  splitWithLoanCount.textContent = withFormalLoan;

  if (btnFilterFormalLoans) {
    btnFilterFormalLoans.textContent = state.lang === 'tl'
      ? `May Active Loan Lamang (${withFormalLoan})`
      : `Institutional Loans Only (${withFormalLoan})`;
    btnFilterFormalLoans.classList.toggle('active', state.loanFilterMode === 'formal');
  }
  if (btnFilterAllDeductions) {
    btnFilterAllDeductions.textContent = state.lang === 'tl'
      ? `Lahat ng Bawas (${withAnyLoan})`
      : `All Payroll Deductions (${withAnyLoan})`;
    btnFilterAllDeductions.classList.toggle('active', state.loanFilterMode === 'all');
  }
}

// -----------------------------------------------------------------------------
// FILTERING & PAGINATION
// -----------------------------------------------------------------------------
function getFilteredRecordsForCurrentView() {
  let list = [...state.records];

  // In "Only Loan" view: ONLY people who have actual loans are listed!
  if (state.currentView === 'onlyLoan') {
    if (state.loanFilterMode === 'formal') {
      list = list.filter(r => r.hasFormalLoan);
    } else {
      list = list.filter(r => r.hasLoan);
    }
  } else if (state.currentView === 'noLoan') {
    list = list.filter(r => !r.hasFormalLoan);
  }

  // Search filter
  const q = state.searchQuery.trim().toLowerCase();
  if (q) {
    list = list.filter(r => {
      if (r.fullName.toLowerCase().includes(q)) return true;
      if (r.payrollDate.toLowerCase().includes(q)) return true;
      if (String(r.formalLoansTotal).includes(q)) return true;
      if (String(r.totalAllDeductions).includes(q)) return true;
      if (r.activeLoansSummaryText.toLowerCase().includes(q)) return true;
      return false;
    });
  }

  return list;
}

// -----------------------------------------------------------------------------
// VIEW NAVIGATION & SWITCHING
// -----------------------------------------------------------------------------
function switchView(viewName) {
  state.currentView = viewName;
  state.currentPage = 1;

  navButtons.forEach(btn => {
    if (btn.dataset.view === viewName) btn.classList.add('active');
    else btn.classList.remove('active');
  });

  const t = i18n[state.lang];

  if (viewName === 'overall') {
    viewStandardTable.classList.add('hidden');
    viewOverallSplit.classList.remove('hidden');
    viewCompute.classList.add('hidden');
    controlsToolbar.classList.remove('hidden');
    pageSizeWrap.classList.add('hidden');
    onlyLoanModeToggle.classList.add('hidden');
    viewStatusTag.innerHTML = `${state.lang === 'tl' ? 'Ipinapakita' : 'Showing'}: <strong>${t.showingOverall}</strong>`;
    renderOverallSplit();
  } else if (viewName === 'compute') {
    viewStandardTable.classList.add('hidden');
    viewOverallSplit.classList.add('hidden');
    viewCompute.classList.remove('hidden');
    controlsToolbar.classList.add('hidden');
    onlyLoanModeToggle.classList.add('hidden');
    viewStatusTag.innerHTML = `${state.lang === 'tl' ? 'Ipinapakita' : 'Showing'}: <strong>${t.showingCompute}</strong>`;
    computeDashboardMetrics();
  } else {
    viewStandardTable.classList.remove('hidden');
    viewOverallSplit.classList.add('hidden');
    viewCompute.classList.add('hidden');
    controlsToolbar.classList.remove('hidden');
    pageSizeWrap.classList.remove('hidden');

    let title = t.showingAll;
    if (viewName === 'onlyLoan') {
      title = state.loanFilterMode === 'formal' ? t.showingOnlyLoans : t.showingOnlyLoansAll;
      onlyLoanModeToggle.classList.remove('hidden');
    } else {
      onlyLoanModeToggle.classList.add('hidden');
    }
    if (viewName === 'noLoan') title = t.showingNoLoans;
    viewStatusTag.innerHTML = `${state.lang === 'tl' ? 'Ipinapakita' : 'Showing'}: <strong>${title}</strong>`;
    renderMainGrid();
  }
}

navButtons.forEach(btn => {
  btn.addEventListener('click', () => switchView(btn.dataset.view));
});

// Only Loan Mode Filter Listeners
if (btnFilterFormalLoans) {
  btnFilterFormalLoans.addEventListener('click', () => {
    state.loanFilterMode = 'formal';
    state.currentPage = 1;
    updateGlobalCounters();
    viewStatusTag.innerHTML = `${state.lang === 'tl' ? 'Ipinapakita' : 'Showing'}: <strong>${i18n[state.lang].showingOnlyLoans}</strong>`;
    renderMainGrid();
  });
}

if (btnFilterAllDeductions) {
  btnFilterAllDeductions.addEventListener('click', () => {
    state.loanFilterMode = 'all';
    state.currentPage = 1;
    updateGlobalCounters();
    viewStatusTag.innerHTML = `${state.lang === 'tl' ? 'Ipinapakita' : 'Showing'}: <strong>${i18n[state.lang].showingOnlyLoansAll}</strong>`;
    renderMainGrid();
  });
}

// -----------------------------------------------------------------------------
// RENDER MAIN DATA GRID
// -----------------------------------------------------------------------------
function renderMainGrid() {
  if (state.records.length === 0) {
    if (emptyDropzoneState) emptyDropzoneState.classList.remove('hidden');
    if (tableScrollHelper) tableScrollHelper.classList.add('hidden');
    if (mainTableScrollContainer) mainTableScrollContainer.classList.add('hidden');
    if (mainPaginationFooter) mainPaginationFooter.classList.add('hidden');
    return;
  } else {
    if (emptyDropzoneState) emptyDropzoneState.classList.add('hidden');
    if (tableScrollHelper) tableScrollHelper.classList.remove('hidden');
    if (mainTableScrollContainer) mainTableScrollContainer.classList.remove('hidden');
    if (mainPaginationFooter) mainPaginationFooter.classList.remove('hidden');
  }

  const filtered = getFilteredRecordsForCurrentView();
  const totalEntries = filtered.length;
  const t = i18n[state.lang];
  const isFormalOnly = state.loanFilterMode === 'formal';

  mainGridHead.innerHTML = '';
  mainGridFoot.innerHTML = '';
  const trH = document.createElement('tr');

  // Checkbox th
  const thSelect = document.createElement('th');
  thSelect.className = 'th-freeze th-freeze-1';
  thSelect.innerHTML = `<input type="checkbox" id="selectAllCheckbox" title="Select All" />`;
  trH.appendChild(thSelect);

  // Row # th
  const thNum = document.createElement('th');
  thNum.className = 'th-freeze th-freeze-2';
  thNum.textContent = t.colRow;
  trH.appendChild(thNum);

  // Actions th
  const thAct = document.createElement('th');
  thAct.className = 'th-freeze th-freeze-3';
  thAct.textContent = t.colAction;
  trH.appendChild(thAct);

  // SVG copy icon (zero emojis)
  const svgCopyIcon = `<svg viewBox="0 0 24 24" width="13" height="13" stroke="currentColor" stroke-width="2" fill="none"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>`;

  if (state.currentView === 'onlyLoan') {
    // SPECIAL ACCURATE ONLY LOAN COLUMNS:
    // Exclude zeroes, keep accurate active loan records, end with TOTAL column!
    const customCols = [
      { key: 'lastName', label: t.colLastName },
      { key: 'firstName', label: t.colFirstName },
      { key: 'payrollDate', label: t.colPayrollDate },
      { key: 'activeLoans', label: t.colActiveLoans },
      { key: 'sssSalary', label: 'SSS Salary Loan' },
      { key: 'sssCalamity', label: 'SSS Calamity Loan' },
      { key: 'hdmfLoan', label: 'HDMF Loan' },
      { key: 'hdmfMp2', label: 'HDMF MP2' },
      { key: 'hdmfCalamity', label: 'HDMF Calamity' },
      { key: 'aubLoan', label: 'AUB Loan' },
      { key: 'cashAdvances', label: 'Cash Advances' }
    ];

    if (!isFormalOnly) {
      customCols.push({ key: 'otherCharges', label: 'Other Charges' });
    }

    customCols.push({ key: 'totalDeduction', label: t.colTotalLoanDeduction });

    customCols.forEach(col => {
      const th = document.createElement('th');
      if (col.key === 'lastName') {
        th.className = 'th-freeze th-freeze-4';
      } else if (col.key === 'firstName') {
        th.className = 'th-freeze th-freeze-5 th-freeze-last';
      } else if (col.key === 'totalDeduction') {
        th.className = 'th-numeric th-total-deduction';
      } else if (['sssSalary', 'sssCalamity', 'hdmfLoan', 'hdmfMp2', 'hdmfCalamity', 'aubLoan', 'cashAdvances', 'otherCharges'].includes(col.key)) {
        th.className = 'th-numeric';
      }
      th.innerHTML = `
        <div class="th-content">
          <span>${col.label}</span>
          <button class="btn-th-copy" title="Copy Column ${col.label}">${svgCopyIcon}</button>
        </div>
      `;
      th.querySelector('.btn-th-copy').addEventListener('click', (e) => {
        e.stopPropagation();
        const vals = filtered.map(r => {
          if (col.key === 'totalDeduction') return isFormalOnly ? r.formalLoansTotal : r.totalAllDeductions;
          if (col.key === 'activeLoans') return r.activeLoansSummaryText;
          return r[col.key] || '';
        }).join('\n');
        copyToClipboard(vals, `Copied column [${col.label}]!`);
      });
      trH.appendChild(th);
    });

  } else {
    // STANDARD VIEW COLUMNS:
    const thLoanStat = document.createElement('th');
    thLoanStat.className = 'th-freeze th-freeze-4 th-freeze-last';
    thLoanStat.textContent = t.colStatus;
    trH.appendChild(thLoanStat);

    state.headers.forEach(hKey => {
      const th = document.createElement('th');
      th.innerHTML = `
        <div class="th-content">
          <span>${hKey}</span>
          <button class="btn-th-copy" title="Copy Column ${hKey}">${svgCopyIcon}</button>
        </div>
      `;
      th.querySelector('.btn-th-copy').addEventListener('click', (e) => {
        e.stopPropagation();
        const colVals = filtered.map(item => item.raw[hKey] !== undefined ? item.raw[hKey] : '').join('\n');
        copyToClipboard(colVals, `Copied column [${hKey}]!`);
      });
      trH.appendChild(th);
    });
  }

  mainGridHead.appendChild(trH);

  // Select all checkbox
  const chkAll = trH.querySelector('#selectAllCheckbox');
  if (chkAll) {
    chkAll.addEventListener('change', (e) => {
      const isChecked = e.target.checked;
      filtered.forEach(r => {
        if (isChecked) state.selectedIndices.add(r._index);
        else state.selectedIndices.delete(r._index);
      });
      updateSelectionBox();
      renderMainGridBody(filtered);
    });
  }

  renderMainGridBody(filtered);
}

function renderMainGridBody(filtered) {
  mainGridBody.innerHTML = '';
  mainGridFoot.innerHTML = '';
  const total = filtered.length;
  const t = i18n[state.lang];

  if (total === 0) {
    mainGridFoot.classList.add('hidden');
    const trEmpty = document.createElement('tr');
    trEmpty.innerHTML = `
      <td colspan="15" style="text-align: center; padding: 3rem; color: var(--text-dim);">
        ${state.lang === 'tl' ? 'Walang talaang tumutugma sa filter o search query.' : 'No records match the current filter or search query.'}
      </td>
    `;
    mainGridBody.appendChild(trEmpty);
    mainPaginationInfo.textContent = '0 records';
    mainPageCounter.textContent = 'Page 0 of 0';
    btnPrev.disabled = true;
    btnNext.disabled = true;
    return;
  }

  // Pagination calculation
  let startIdx = 0;
  let endIdx = total;
  let totalPages = 1;

  if (state.pageSize !== 'all') {
    const size = parseInt(state.pageSize, 10);
    totalPages = Math.ceil(total / size);
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    if (state.currentPage < 1) state.currentPage = 1;

    startIdx = (state.currentPage - 1) * size;
    endIdx = Math.min(startIdx + size, total);
  } else {
    state.currentPage = 1;
  }

  mainPaginationInfo.textContent = `${state.lang === 'tl' ? 'Ipinapakita ang' : 'Showing'} ${startIdx + 1} ${state.lang === 'tl' ? 'hanggang' : 'to'} ${endIdx} ${state.lang === 'tl' ? 'ng' : 'of'} ${total} entries`;
  mainPageCounter.textContent = `Page ${state.currentPage} of ${totalPages}`;
  btnPrev.disabled = state.currentPage <= 1;
  btnNext.disabled = state.currentPage >= totalPages;

  const pageSlice = filtered.slice(startIdx, endIdx);

  pageSlice.forEach((item, sliceIdx) => {
    const tr = document.createElement('tr');
    const isSelected = state.selectedIndices.has(item._index);
    if (isSelected) tr.classList.add('selected');
    if (item.hasLoan) tr.classList.add('has-loan');

    // 1. Checkbox
    const tdChk = document.createElement('td');
    tdChk.className = 'td-freeze td-freeze-1';
    tdChk.innerHTML = `<input type="checkbox" ${isSelected ? 'checked' : ''} />`;
    tdChk.querySelector('input').addEventListener('change', (e) => {
      if (e.target.checked) {
        state.selectedIndices.add(item._index);
        tr.classList.add('selected');
      } else {
        state.selectedIndices.delete(item._index);
        tr.classList.remove('selected');
      }
      updateSelectionBox();
    });
    tr.appendChild(tdChk);

    // 2. Row #
    const tdNum = document.createElement('td');
    tdNum.className = 'td-freeze td-freeze-2';
    tdNum.textContent = startIdx + sliceIdx + 1;
    tr.appendChild(tdNum);

    // 3. Action Buttons (Copy / Inspect)
    const tdAct = document.createElement('td');
    tdAct.className = 'td-freeze td-freeze-3';
    tdAct.innerHTML = `
      <div style="display: flex; align-items: center; gap: 0.25rem;">
        <button class="btn-row-copy-fast" title="Copy Row as JSON">Copy</button>
        <button class="btn-row-zoom" title="Inspect Details">Zoom</button>
      </div>
    `;
    tdAct.querySelector('.btn-row-copy-fast').addEventListener('click', (e) => {
      e.stopPropagation();
      copyToClipboard(JSON.stringify(item.raw, null, 2), `Copied row for ${item.fullName}!`);
    });
    tdAct.querySelector('.btn-row-zoom').addEventListener('click', (e) => {
      e.stopPropagation();
      openRowModal(item);
    });
    tr.appendChild(tdAct);

    if (state.currentView === 'onlyLoan') {
      // SPECIAL ONLY LOAN RENDERING:
      // Show ONLY active loans, replace zeroes with clean dashes '—'

      // Last Name
      const tdLast = document.createElement('td');
      tdLast.className = 'td-freeze td-freeze-4';
      tdLast.innerHTML = `<strong>${escapeHtml(item.lastName)}</strong>`;
      tr.appendChild(tdLast);

      // First Name
      const tdFirst = document.createElement('td');
      tdFirst.className = 'td-freeze td-freeze-5 td-freeze-last';
      tdFirst.textContent = item.firstName;
      tr.appendChild(tdFirst);

      // Payroll Date
      const tdDate = document.createElement('td');
      tdDate.style.fontFamily = 'JetBrains Mono, monospace';
      tdDate.style.fontSize = '0.78rem';
      tdDate.textContent = item.payrollDate;
      tr.appendChild(tdDate);

      // Active Loans Tag List (Only the loans they actually have!)
      const tdActiveList = document.createElement('td');
      if (item.activeLoanItems.length > 0) {
        const pills = item.activeLoanItems.map(loan => `<span class="loan-tag-pill">${loan.name}: ${formatPeso(loan.amount)}</span>`).join(' ');
        tdActiveList.innerHTML = `<div class="active-loans-tags">${pills}</div>`;
      } else {
        tdActiveList.innerHTML = `<span style="color: var(--text-dim); font-size: 0.75rem;">${t.noActiveDeductions}</span>`;
      }
      tr.appendChild(tdActiveList);

      // Specific Loan Columns: SSS Salary, SSS Calamity, HDMF Loan, HDMF MP2, HDMF Calamity, AUB Loan, Cash Advances
      const isFormalOnly = state.loanFilterMode === 'formal';
      const loanCols = [
        item.sssSalary,
        item.sssCalamity,
        item.hdmfLoan,
        item.hdmfMp2,
        item.hdmfCalamity,
        item.aubLoan,
        item.cashAdvances
      ];
      if (!isFormalOnly) {
        loanCols.push(item.otherCharges);
      }

      loanCols.forEach(val => {
        const td = document.createElement('td');
        if (val > 0) {
          td.className = 'cell-amount has-value';
          td.textContent = formatPeso(val);
        } else {
          // Zero loan -> Don't clutter! Display clean dash!
          td.className = 'cell-amount';
          td.innerHTML = `<span class="cell-zero-dash">—</span>`;
        }
        tr.appendChild(td);
      });

      // Total Deduction at end (Sa Dulo!)
      const tdTotal = document.createElement('td');
      const activeTotal = isFormalOnly ? item.formalLoansTotal : item.totalAllDeductions;
      tdTotal.className = 'cell-amount td-total-col';
      tdTotal.innerHTML = `<span class="total-loan-cell">${formatPeso(activeTotal)}</span>`;
      tr.appendChild(tdTotal);

    } else {
      // Standard View Rendering
      const tdStat = document.createElement('td');
      tdStat.className = 'td-freeze td-freeze-4 td-freeze-last';
      if (item.hasLoan) {
        tdStat.innerHTML = `<span class="deduction-pill deduction-active">${formatPeso(item.totalAllDeductions)}</span>`;
      } else {
        tdStat.innerHTML = `<span class="deduction-pill deduction-zero">PHP 0.00</span>`;
      }
      tr.appendChild(tdStat);

      state.headers.forEach(hKey => {
        const td = document.createElement('td');
        const val = item.raw[hKey];
        const displayVal = val !== undefined && val !== null ? val : '—';

        if (typeof val === 'number') {
          td.className = 'cell-amount' + (val !== 0 ? ' has-value' : '');
          td.textContent = val !== 0 ? val.toLocaleString('en-US') : '0';
        } else {
          td.innerHTML = `<span class="cell-interactive">${escapeHtml(String(displayVal))}</span>`;
        }

        td.addEventListener('click', () => {
          copyToClipboard(String(displayVal), `Copied [${hKey}]: "${displayVal}"`);
        });

        tr.appendChild(td);
      });
    }

    mainGridBody.appendChild(tr);
  });

  // RENDER GRAND TOTAL ROW IN TFOOT (FOR ONLY LOAN VIEW)
  if (state.currentView === 'onlyLoan') {
    mainGridFoot.classList.remove('hidden');
    const isFormalOnly = state.loanFilterMode === 'formal';
    const sumSssSalary = filtered.reduce((acc, r) => acc + r.sssSalary, 0);
    const sumSssCalamity = filtered.reduce((acc, r) => acc + r.sssCalamity, 0);
    const sumHdmfLoan = filtered.reduce((acc, r) => acc + r.hdmfLoan, 0);
    const sumHdmfMp2 = filtered.reduce((acc, r) => acc + r.hdmfMp2, 0);
    const sumHdmfCalamity = filtered.reduce((acc, r) => acc + r.hdmfCalamity, 0);
    const sumAubLoan = filtered.reduce((acc, r) => acc + r.aubLoan, 0);
    const sumCashAdvances = filtered.reduce((acc, r) => acc + r.cashAdvances, 0);
    const sumOtherCharges = filtered.reduce((acc, r) => acc + r.otherCharges, 0);
    const grandTotalAll = filtered.reduce((acc, r) => acc + (isFormalOnly ? r.formalLoansTotal : r.totalAllDeductions), 0);

    const trFoot = document.createElement('tr');
    let footColsHtml = `
      <td colspan="5" class="td-freeze td-freeze-footer" style="text-align: right; text-transform: uppercase; font-weight: 800;">
        ${t.grandTotal} (${filtered.length} ${t.borrowersCount}):
      </td>
      <td style="text-align: center; color: var(--text-dim); font-family: 'JetBrains Mono', monospace;">—</td>
      <td style="text-align: center; color: var(--text-dim); font-family: 'JetBrains Mono', monospace;">—</td>
      <td class="cell-amount">${sumSssSalary > 0 ? formatPeso(sumSssSalary) : '—'}</td>
      <td class="cell-amount">${sumSssCalamity > 0 ? formatPeso(sumSssCalamity) : '—'}</td>
      <td class="cell-amount">${sumHdmfLoan > 0 ? formatPeso(sumHdmfLoan) : '—'}</td>
      <td class="cell-amount">${sumHdmfMp2 > 0 ? formatPeso(sumHdmfMp2) : '—'}</td>
      <td class="cell-amount">${sumHdmfCalamity > 0 ? formatPeso(sumHdmfCalamity) : '—'}</td>
      <td class="cell-amount">${sumAubLoan > 0 ? formatPeso(sumAubLoan) : '—'}</td>
      <td class="cell-amount">${sumCashAdvances > 0 ? formatPeso(sumCashAdvances) : '—'}</td>
    `;
    if (!isFormalOnly) {
      footColsHtml += `<td class="cell-amount">${sumOtherCharges > 0 ? formatPeso(sumOtherCharges) : '—'}</td>`;
    }
    footColsHtml += `<td class="cell-amount td-total-col" style="color: #ffffff; background: linear-gradient(135deg, #f59e0b, #b45309); font-weight: 800; font-size: 0.95rem; box-shadow: 0 0 12px rgba(245, 158, 11, 0.5);">${formatPeso(grandTotalAll)}</td>`;
    trFoot.innerHTML = footColsHtml;
    mainGridFoot.appendChild(trFoot);
  } else {
    mainGridFoot.classList.add('hidden');
  }

  if (mainTableScrollContainer) {
    requestAnimationFrame(() => {
      mainTableScrollContainer.dispatchEvent(new Event('scroll'));
    });
  }
}

// -----------------------------------------------------------------------------
// RENDER OVERALL SPLIT VIEW
// -----------------------------------------------------------------------------
function renderOverallSplit() {
  const q = state.searchQuery.trim().toLowerCase();
  const t = i18n[state.lang];

  // Top Table: No Loans
  let noLoanList = state.records.filter(r => !r.hasLoan);
  if (q) noLoanList = noLoanList.filter(r => r.fullName.toLowerCase().includes(q));

  splitNoLoanBody.innerHTML = '';
  if (noLoanList.length === 0) {
    splitNoLoanBody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 1.5rem; color: var(--text-dim);">${state.lang === 'tl' ? 'Walang talaan.' : 'No records.'}</td></tr>`;
  } else {
    noLoanList.forEach((r, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="text-align: center; font-family: 'JetBrains Mono', monospace; color: var(--text-dim);">${idx + 1}</td>
        <td>
          <button class="btn-row-copy-fast" data-idx="${r._index}">Copy</button>
          <button class="btn-row-zoom" data-idx="${r._index}">Zoom</button>
        </td>
        <td><strong>${escapeHtml(r.lastName)}</strong></td>
        <td>${escapeHtml(r.firstName)}</td>
        <td style="font-family: 'JetBrains Mono', monospace;">${r.payrollDate}</td>
        <td><span class="badge badge-green">${t.noActiveDeductions}</span></td>
      `;
      tr.querySelector('.btn-row-copy-fast').addEventListener('click', () => copyToClipboard(JSON.stringify(r.raw, null, 2), `Copied row for ${r.fullName}!`));
      tr.querySelector('.btn-row-zoom').addEventListener('click', () => openRowModal(r));
      splitNoLoanBody.appendChild(tr);
    });
  }

  // Bottom Table: With Loans
  let withLoanList = state.records.filter(r => r.hasLoan);
  if (q) withLoanList = withLoanList.filter(r => r.fullName.toLowerCase().includes(q));

  splitWithLoanBody.innerHTML = '';
  splitWithLoanFoot.innerHTML = '';

  if (withLoanList.length === 0) {
    splitWithLoanBody.innerHTML = `<tr><td colspan="12" style="text-align: center; padding: 1.5rem; color: var(--text-dim);">${state.lang === 'tl' ? 'Walang talaan.' : 'No records.'}</td></tr>`;
  } else {
    withLoanList.forEach((r, idx) => {
      const tr = document.createElement('tr');
      tr.classList.add('has-loan');

      const pills = r.activeLoanItems.map(loan => `<span class="loan-tag-pill">${loan.name}: ${formatPeso(loan.amount)}</span>`).join(' ');

      tr.innerHTML = `
        <td style="text-align: center; font-family: 'JetBrains Mono', monospace; color: var(--text-dim);">${idx + 1}</td>
        <td>
          <button class="btn-row-copy-fast" data-idx="${r._index}">Copy</button>
          <button class="btn-row-zoom" data-idx="${r._index}">Zoom</button>
        </td>
        <td><strong>${escapeHtml(r.lastName)}</strong></td>
        <td>${escapeHtml(r.firstName)}</td>
        <td><strong class="text-gold" style="font-family: 'JetBrains Mono', monospace;">${formatPeso(r.totalAllDeductions)}</strong></td>
        <td><div class="active-loans-tags">${pills}</div></td>
        <td style="font-family: 'JetBrains Mono', monospace;">${r.sssSalary > 0 ? formatPeso(r.sssSalary) : '—'}</td>
        <td style="font-family: 'JetBrains Mono', monospace;">${r.sssCalamity > 0 ? formatPeso(r.sssCalamity) : '—'}</td>
        <td style="font-family: 'JetBrains Mono', monospace;">${r.hdmfLoan > 0 ? formatPeso(r.hdmfLoan) : '—'}</td>
        <td style="font-family: 'JetBrains Mono', monospace;">${r.hdmfMp2 > 0 ? formatPeso(r.hdmfMp2) : '—'}</td>
        <td style="font-family: 'JetBrains Mono', monospace; font-weight: 700; color: #f87171;">${r.aubLoan > 0 ? formatPeso(r.aubLoan) : '—'}</td>
        <td style="font-family: 'JetBrains Mono', monospace;">${r.otherCharges > 0 ? formatPeso(r.otherCharges) : '—'}</td>
      `;
      tr.querySelector('.btn-row-copy-fast').addEventListener('click', () => copyToClipboard(JSON.stringify(r.raw, null, 2), `Copied row for ${r.fullName}!`));
      tr.querySelector('.btn-row-zoom').addEventListener('click', () => openRowModal(r));
      splitWithLoanBody.appendChild(tr);
    });

    // Foot sum for Split Bottom Table
    const grandSum = withLoanList.reduce((acc, r) => acc + r.totalAllDeductions, 0);
    const trF = document.createElement('tr');
    trF.innerHTML = `
      <td colspan="4" style="text-align: right; text-transform: uppercase; font-weight: 800;">${t.grandTotal}:</td>
      <td style="background: var(--gold-primary); color: #000000;">${formatPeso(grandSum)}</td>
      <td colspan="7"></td>
    `;
    splitWithLoanFoot.appendChild(trF);
  }
}

// Split View Copy Buttons
btnCopyNoLoanTsv.addEventListener('click', () => {
  const noLoan = state.records.filter(r => !r.hasLoan);
  const tsv = ['Last Name\tFirst Name\tPayroll Date\tLoan Status', ...noLoan.map(r => `${r.lastName}\t${r.firstName}\t${r.payrollDate}\tNo Loan`)].join('\n');
  copyToClipboard(tsv, `Copied ${noLoan.length} no-loan records as TSV!`);
});

btnCopyNoLoanJson.addEventListener('click', () => {
  const noLoan = state.records.filter(r => !r.hasLoan).map(r => r.raw);
  copyToClipboard(JSON.stringify(noLoan, null, 2), `Copied no-loan list as JSON!`);
});

btnCopyWithLoanTsv.addEventListener('click', () => {
  const withLoan = state.records.filter(r => r.hasLoan);
  const header = 'Last Name\tFirst Name\tTotal Deductions\tActive Loans List\tSSS Salary\tSSS Calamity\tHDMF Loan\tHDMF MP2\tAUB Loan\tOther Charges';
  const lines = withLoan.map(r => `${r.lastName}\t${r.firstName}\t${r.totalAllDeductions}\t"${r.activeLoansSummaryText}"\t${r.sssSalary}\t${r.sssCalamity}\t${r.hdmfLoan}\t${r.hdmfMp2}\t${r.aubLoan}\t${r.otherCharges}`);
  copyToClipboard([header, ...lines].join('\n'), `Copied ${withLoan.length} loan records as TSV!`);
});

btnCopyWithLoanJson.addEventListener('click', () => {
  const withLoan = state.records.filter(r => r.hasLoan).map(r => r.raw);
  copyToClipboard(JSON.stringify(withLoan, null, 2), `Copied loan records as JSON!`);
});

// -----------------------------------------------------------------------------
// COMPUTE DASHBOARD METRICS
// -----------------------------------------------------------------------------
function computeDashboardMetrics() {
  const totalEmployees = state.records.length;
  const borrowers = state.records.filter(r => r.hasLoan);
  const totalDeductionsSum = state.records.reduce((acc, r) => acc + r.totalAllDeductions, 0);
  const avgDeduction = borrowers.length > 0 ? totalDeductionsSum / borrowers.length : 0;
  const percentage = totalEmployees > 0 ? ((borrowers.length / totalEmployees) * 100).toFixed(1) : 0;
  const t = i18n[state.lang];

  const categories = [
    { key: 'aubLoan', name: 'AUB Bank Loan' },
    { key: 'sssSalary', name: 'SSS Salary Loan' },
    { key: 'hdmfMp2', name: 'HDMF MP2 Savings' },
    { key: 'hdmfLoan', name: 'HDMF Loan (Pag-IBIG)' },
    { key: 'sssCalamity', name: 'SSS Calamity Loan' },
    { key: 'hdmfCalamity', name: 'HDMF Calamity Loan' },
    { key: 'otherCharges', name: 'Other Charges & Deduction' },
    { key: 'cashAdvances', name: 'Cash Advances' },
    { key: 'adjContrib', name: 'Adjustment on Contributions' },
    { key: 'hmoShare', name: 'HMO Employee Share' },
    { key: 'sssEmergency', name: 'SSS Emergency Loan' }
  ];

  let highestCat = { name: 'AUB Bank Loan', sum: 0 };

  categorySummaryGrid.innerHTML = '';
  categories.forEach(cat => {
    const sum = state.records.reduce((acc, r) => acc + (r[cat.key] || 0), 0);
    const count = state.records.filter(r => (r[cat.key] || 0) > 0).length;

    if (sum > highestCat.sum) {
      highestCat = { name: cat.name, sum };
    }

    const card = document.createElement('div');
    card.className = 'category-pill-card';
    card.innerHTML = `
      <div style="display: flex; align-items: center; justify-content: space-between;">
        <span class="cat-name">${cat.name}</span>
        <span class="badge ${count > 0 ? 'badge-gold' : ''}">${count} ${state.lang === 'tl' ? 'tao' : 'borrowers'}</span>
      </div>
      <div class="cat-amount ${count > 0 ? 'text-gold' : ''}">${formatPeso(sum)}</div>
      <span class="cat-borrowers">${count > 0 ? `${state.lang === 'tl' ? 'Bawas sa' : 'Deducted from'} ${count} accounts` : t.noActiveDeductions}</span>
    `;
    categorySummaryGrid.appendChild(card);
  });

  computeTotalDeductions.textContent = formatPeso(totalDeductionsSum);
  computeActiveBorrowersCount.textContent = `${borrowers.length} Active Borrowers`;
  computeTotalHeadcount.textContent = totalEmployees;
  computeLoanPercentage.textContent = `${percentage}% ${t.kpiLoanPercentage}`;
  computeHighestCategory.textContent = highestCat.name;
  computeHighestCategorySum.textContent = `${formatPeso(highestCat.sum)} total`;
  computeAvgDeduction.textContent = formatPeso(avgDeduction);

  // Top Borrowers Ranking List
  const topList = [...state.records]
    .filter(r => r.totalAllDeductions > 0)
    .sort((a, b) => b.totalAllDeductions - a.totalAllDeductions)
    .slice(0, 10);

  topBorrowersList.innerHTML = '';
  topList.forEach((r, rankIdx) => {
    const item = document.createElement('div');
    item.className = 'top-borrower-item';

    let dominantType = 'Various Deductions';
    if (r.aubLoan > 0) dominantType = `AUB Loan (${formatPeso(r.aubLoan)})`;
    else if (r.sssSalary > 0) dominantType = `SSS Salary (${formatPeso(r.sssSalary)})`;
    else if (r.hdmfLoan > 0) dominantType = `HDMF Loan (${formatPeso(r.hdmfLoan)})`;
    else if (r.hdmfMp2 > 0) dominantType = `HDMF MP2 (${formatPeso(r.hdmfMp2)})`;

    item.innerHTML = `
      <div class="borrower-meta">
        <span class="borrower-rank">${rankIdx + 1}</span>
        <div>
          <div class="borrower-name">${escapeHtml(r.fullName)}</div>
          <div class="borrower-loan-type">${dominantType}</div>
        </div>
      </div>
      <div class="borrower-amount">${formatPeso(r.totalAllDeductions)}</div>
    `;
    item.addEventListener('click', () => openRowModal(r));
    topBorrowersList.appendChild(item);
  });

  updateCalculator();
}

btnCopyComputeBreakdown.addEventListener('click', () => {
  const categories = [
    { key: 'aubLoan', name: 'AUB Bank Loan' },
    { key: 'sssSalary', name: 'SSS Salary Loan' },
    { key: 'hdmfMp2', name: 'HDMF MP2 Savings' },
    { key: 'hdmfLoan', name: 'HDMF Loan (Pag-IBIG)' },
    { key: 'sssCalamity', name: 'SSS Calamity Loan' },
    { key: 'otherCharges', name: 'Other Charges & Deduction' }
  ];

  const header = 'Loan Category\tTotal Deductions\tBorrower Count';
  const lines = categories.map(cat => {
    const sum = state.records.reduce((acc, r) => acc + (r[cat.key] || 0), 0);
    const count = state.records.filter(r => (r[cat.key] || 0) > 0).length;
    return `${cat.name}\t${sum.toFixed(2)}\t${count}`;
  });

  copyToClipboard([header, ...lines].join('\n'), 'Copied Loan Category Breakdown as TSV!');
});

// -----------------------------------------------------------------------------
// LIVE LOAN CALCULATOR
// -----------------------------------------------------------------------------
function updateCalculator() {
  const principal = parseFloat(calcPrincipal.value) || 0;
  const months = parseInt(calcMonths.value, 10) || 12;
  const annualRatePercent = parseFloat(calcInterestRate.value) || 0;

  const totalInterest = principal * (annualRatePercent / 100) * (months / 12);
  const totalPayable = principal + totalInterest;
  const monthlyAmort = months > 0 ? totalPayable / months : 0;
  const semiMonthlyCutoff = monthlyAmort / 2;

  calcMonthlyDeduction.textContent = formatPeso(monthlyAmort);
  calcSemiMonthlyDeduction.textContent = formatPeso(semiMonthlyCutoff);
  calcTotalInterest.textContent = formatPeso(totalInterest);
  calcTotalPayable.textContent = formatPeso(totalPayable);
}

[calcPrincipal, calcMonths, calcInterestRate, calcLoanType].forEach(el => {
  el.addEventListener('input', updateCalculator);
});

btnCopyCalcResult.addEventListener('click', () => {
  const text = `SYSGEN RPO LOAN COMPUTATION (Ivy Salamanca)
Loan Type: ${calcLoanType.value}
Principal Amount: ${formatPeso(calcPrincipal.value)}
Term: ${calcMonths.value} months
Annual Interest Rate: ${calcInterestRate.value}%
Monthly Amortization: ${calcMonthlyDeduction.textContent}
Semi-Monthly Cutoff (15/30): ${calcSemiMonthlyDeduction.textContent}
Total Interest: ${calcTotalInterest.textContent}
Total Payable: ${calcTotalPayable.textContent}
Engineered by: Christian Antero Salamanca`;

  copyToClipboard(text, 'Copied calculation summary to clipboard!');
});

// -----------------------------------------------------------------------------
// DEDICATED ONLY LOANS EXCEL (.XLSX) WITH LOCKED HEADERS
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------
// DEDICATED ONLY LOANS EXCEL (.XLSX) WITH ORIGINAL UPLOADED FORMAT & FROZEN PANES
// -----------------------------------------------------------------------------
btnDownloadOnlyLoansXlsx.addEventListener('click', async () => {
  try {
    const isFormalOnly = state.loanFilterMode === 'formal';
    const borrowers = state.records.filter(r => isFormalOnly ? r.hasFormalLoan : r.hasLoan);
    const t = i18n[state.lang];

    const Excel = ExcelJS.default || ExcelJS;
    const wb = new Excel.Workbook();
    wb.creator = 'Sysgen RPO Loans Engine';
    wb.lastModifiedBy = 'Ivy Salamanca';
    wb.created = new Date();

    const ws = wb.addWorksheet('Active Loans Only', {
      views: [
        {
          state: 'frozen',
          ySplit: 1,
          xSplit: 2,
          topLeftCell: 'C2',
          activePane: 'bottomRight'
        }
      ]
    });

    // 1. Maintain the EXACT original uploaded columns, plus append TOTAL LOAN DEDUCTION at the end
    const totalColTitle = state.lang === 'tl' ? 'KABUUANG BAWAS SA LOAN' : 'TOTAL LOAN DEDUCTION';
    const exportColumns = [...state.headers, totalColTitle];

    ws.columns = exportColumns.map((headerText, i) => ({
      header: headerText,
      key: `col_${i}`
    }));

    // 2. Format & Style Header Row (Locked at top)
    const headerRow = ws.getRow(1);
    headerRow.height = 32;
    headerRow.eachCell((cell, colNum) => {
      const isTotalCol = colNum === exportColumns.length;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: isTotalCol ? 'FF1E293B' : 'FF18181B' }
      };
      cell.font = {
        name: 'Segoe UI',
        size: 10.5,
        bold: true,
        color: { argb: isTotalCol ? 'FFFDE68A' : 'FFFFFFFF' }
      };
      const hText = exportColumns[colNum - 1];
      const isDate = hText.toLowerCase().includes('date');
      const isName = hText.toLowerCase().includes('name');
      cell.alignment = {
        vertical: 'middle',
        horizontal: isName ? 'left' : isDate ? 'center' : 'right',
        wrapText: false
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF374151' } },
        bottom: { style: 'medium', color: { argb: 'FFF59E0B' } },
        left: { style: 'thin', color: { argb: 'FF374151' } },
        right: { style: 'thin', color: { argb: 'FF374151' } }
      };
    });

    // 3. Populate Data Rows (Only employees with loans, exact format)
    borrowers.forEach((b, rowIdx) => {
      const activeTotal = isFormalOnly ? b.formalLoansTotal : b.totalAllDeductions;
      const rowValues = state.headers.map(hKey => {
        const val = b.raw[hKey];
        if (typeof val === 'number') {
          return val !== 0 ? val : null;
        }
        return val !== undefined && val !== null ? val : '';
      });
      rowValues.push(activeTotal);

      const row = ws.addRow(rowValues);
      row.height = 22;

      const isEvenRow = rowIdx % 2 === 0;
      const rowBgColor = isEvenRow ? 'FFFFFFFF' : 'FFF9FAFB';

      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const isTotalCol = colNum === exportColumns.length;
        const hText = exportColumns[colNum - 1];
        const isDate = hText.toLowerCase().includes('date');
        const isName = hText.toLowerCase().includes('name');

        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isTotalCol ? 'FFFEF3C7' : rowBgColor }
        };

        cell.font = {
          name: 'Segoe UI',
          size: 10,
          bold: isTotalCol,
          color: { argb: isTotalCol ? 'FF92400E' : 'FF1F2937' }
        };

        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        };

        if (typeof cell.value === 'number') {
          cell.numFmt = '"₱"#,##0.00;("₱"#,##0.00);"-"';
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (isDate) {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: isName ? 'left' : 'center' };
        }
      });
    });

    // 4. Grand Total Row at Bottom
    const footValues = state.headers.map((hKey, i) => {
      if (i === 0) return state.lang === 'tl' ? 'KABUUANG TOTAL' : 'GRAND TOTAL';
      if (i === 1) return `(${borrowers.length} ${t.borrowersCount})`;
      const sum = borrowers.reduce((acc, r) => {
        const v = r.raw[hKey];
        return acc + (typeof v === 'number' && v > 0 ? v : 0);
      }, 0);
      return sum > 0 ? sum : null;
    });
    const grandTotalAll = borrowers.reduce((acc, r) => acc + (isFormalOnly ? r.formalLoansTotal : r.totalAllDeductions), 0);
    footValues.push(grandTotalAll);

    const footRow = ws.addRow(footValues);
    footRow.height = 28;
    footRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFDE68A' }
      };
      cell.font = {
        name: 'Segoe UI',
        size: 10.5,
        bold: true,
        color: { argb: 'FF18181B' }
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF18181B' } },
        bottom: { style: 'double', color: { argb: 'FF18181B' } },
        left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
      };
      if (typeof cell.value === 'number') {
        cell.numFmt = '"₱"#,##0.00;("₱"#,##0.00);"-"';
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: colNum <= 2 ? 'left' : 'center' };
      }
    });

    // 5. Compute Dynamic Auto-Width for all columns (Nothing will overflow / get clipped!)
    ws.columns.forEach((col) => {
      let maxLen = col.header ? String(col.header).length : 12;
      col.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.value !== null && cell.value !== undefined) {
          let valStr = String(cell.value);
          if (typeof cell.value === 'number') {
            valStr = '₱' + cell.value.toLocaleString('en-US', { minimumFractionDigits: 2 });
          }
          if (valStr.length > maxLen) maxLen = valStr.length;
        }
      });
      col.width = Math.max(maxLen + 5, 17);
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Sysgen_RPO_Active_Loans_Only_${isFormalOnly ? 'Institutional' : 'All'}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast(state.lang === 'tl'
      ? 'Na-download na ang Excel ng May Loan Lamang! (Original format + Grand Total sa ibaba + Locked Header & Names)'
      : 'Downloaded Sysgen_RPO_Active_Loans_Only.xlsx (Original format + Grand Total row + Locked Header & Names)!');
  } catch (err) {
    console.error('Error in ExcelJS export:', err);
    showToast('Export error: ' + err.message);
  }
});

// Full Multi-Sheet Download
btnExportAllXlsx.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();

  // Sheet 1: All Records
  const wsAll = XLSX.utils.json_to_sheet(state.records.map(r => r.raw));
  wsAll['!views'] = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];
  XLSX.utils.book_append_sheet(wb, wsAll, 'All Payroll Records');

  // Sheet 2: Only Active Loans
  const wsLoans = XLSX.utils.json_to_sheet(state.records.filter(r => r.hasLoan).map(r => r.raw));
  wsLoans['!views'] = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];
  XLSX.utils.book_append_sheet(wb, wsLoans, 'Only Loans');

  // Sheet 3: Without Loans
  const wsNoLoans = XLSX.utils.json_to_sheet(state.records.filter(r => !r.hasLoan).map(r => r.raw));
  wsNoLoans['!views'] = [{ state: 'frozen', ySplit: 1, xSplit: 2 }];
  XLSX.utils.book_append_sheet(wb, wsNoLoans, 'No Loans');

  // Compute column auto-widths for all sheets so headers don't clip
  const colWidths = state.headers.map(h => {
    let maxLen = String(h).length;
    state.records.forEach(r => {
      const v = r.raw[h];
      if (v !== undefined && v !== null) {
        const len = String(v).length;
        if (len > maxLen) maxLen = len;
      }
    });
    return { wch: Math.max(maxLen + 4, 15) };
  });
  wsAll['!cols'] = colWidths;
  wsLoans['!cols'] = colWidths;
  wsNoLoans['!cols'] = colWidths;

  XLSX.writeFile(wb, 'Sysgen_RPO_Complete_Loans_Report.xlsx');
  showToast('Downloaded Sysgen_RPO_Complete_Loans_Report.xlsx!');
});

// -----------------------------------------------------------------------------
// BILINGUAL LANGUAGE SWITCHER
// -----------------------------------------------------------------------------
function setLanguage(lang) {
  state.lang = lang;
  localStorage.setItem('sysgen_lang', lang);

  btnLangEn.classList.toggle('active', lang === 'en');
  btnLangTl.classList.toggle('active', lang === 'tl');

  const t = i18n[lang];

  // Update DOM Elements
  document.getElementById('txtSplashGreeting').innerHTML = t.splashGreeting;
  document.getElementById('txtSplashRole').textContent = t.splashRole;
  document.getElementById('txtWarningTitle').textContent = t.warningTitle;
  document.getElementById('txtWarningDesc').textContent = t.warningDesc;
  document.getElementById('txtBtnLogin').textContent = t.btnLogin;
  document.getElementById('txtCreditLabel').textContent = t.creditLabel;
  document.getElementById('txtBadgeEngine').textContent = t.badgeEngine;
  document.getElementById('txtHeaderSubtitle').textContent = t.headerSubtitle;
  document.getElementById('txtLangLabel').textContent = t.langLabel;
  document.getElementById('txtThemeLabel').textContent = t.themeLabel;

  document.getElementById('txtNavExcel').textContent = t.navExcel;
  document.getElementById('txtNavOnlyLoan').textContent = t.navOnlyLoan;
  document.getElementById('txtNavNoLoan').textContent = t.navNoLoan;
  document.getElementById('txtNavOverall').textContent = t.navOverall;
  document.getElementById('txtNavOverallSub').textContent = t.navOverallSub;
  document.getElementById('txtNavCompute').textContent = t.navCompute;
  document.getElementById('txtNavComputeSub').textContent = t.navComputeSub;

  document.getElementById('txtBtnUpload').textContent = t.btnUpload;
  document.getElementById('txtBtnDownloadOnlyLoans').textContent = t.btnDownloadOnlyLoans;
  document.getElementById('txtBtnDownloadAll').textContent = t.btnDownloadAll;
  document.getElementById('txtPageSizeLabel').textContent = t.pageSizeLabel;
  document.getElementById('txtBtnCopyExport').textContent = t.btnCopyExport;
  document.getElementById('txtExportJson').textContent = t.exportJson;
  document.getElementById('txtExportTsv').textContent = t.exportTsv;
  document.getElementById('txtExportCsv').textContent = t.exportCsv;
  document.getElementById('txtExportMd').textContent = t.exportMd;

  globalSearchInput.placeholder = t.searchPlaceholder;

  document.getElementById('txtSplitTopBadge').textContent = t.splitTopBadge;
  document.getElementById('txtSplitTopTitle').innerHTML = `${t.splitTopTitle} (<span id="splitNoLoanCount">${countNoLoan.textContent}</span>)`;
  document.getElementById('txtSplitBottomBadge').textContent = t.splitBottomBadge;
  document.getElementById('txtSplitBottomTitle').innerHTML = `${t.splitBottomTitle} (<span id="splitWithLoanCount">${countOnlyLoan.textContent}</span>)`;

  document.getElementById('txtKpiTotalDeductions').textContent = t.kpiTotalDeductions;
  document.getElementById('txtKpiHeadcount').textContent = t.kpiHeadcount;
  document.getElementById('txtKpiHighestCat').textContent = t.kpiHighestCat;
  document.getElementById('txtKpiAvgDeduction').textContent = t.kpiAvgDeduction;
  document.getElementById('txtKpiAvgDeductionSub').textContent = t.kpiAvgDeductionSub;
  document.getElementById('txtBreakdownTitle').textContent = t.breakdownTitle;
  document.getElementById('txtCalcTitle').textContent = t.calcTitle;
  document.getElementById('txtCalcBadge').textContent = t.calcBadge;
  document.getElementById('txtCalcLoanType').textContent = t.calcLoanType;
  document.getElementById('txtCalcPrincipal').textContent = t.calcPrincipal;
  document.getElementById('txtCalcTerm').textContent = t.calcTerm;
  document.getElementById('txtCalcRate').textContent = t.calcRate;
  document.getElementById('txtCalcMonthlyLabel').textContent = t.calcMonthlyLabel;
  document.getElementById('txtCalcSemiMonthlyLabel').textContent = t.calcSemiMonthlyLabel;
  document.getElementById('txtCalcInterestLabel').textContent = t.calcInterestLabel;
  document.getElementById('txtCalcPayableLabel').textContent = t.calcPayableLabel;
  document.getElementById('txtTopBorrowersTitle').textContent = t.topBorrowersTitle;
  document.getElementById('txtTopBorrowersBadge').textContent = t.topBorrowersBadge;

  if (btnSplashLangEn && btnSplashLangTl) {
    btnSplashLangEn.classList.toggle('active', lang === 'en');
    btnSplashLangTl.classList.toggle('active', lang === 'tl');
  }
  const txtSplashLangLabel = document.getElementById('txtSplashLangLabel');
  if (txtSplashLangLabel) txtSplashLangLabel.textContent = t.splashLangLabel;

  const txtLoanFilterLabel = document.getElementById('txtLoanFilterLabel');
  if (txtLoanFilterLabel) txtLoanFilterLabel.textContent = t.loanFilterLabel;

  // Update Scroll Helper Elements
  if (txtScrollNavTitle) txtScrollNavTitle.textContent = t.scrollNavTitle;
  if (txtBtnScrollLeft) txtBtnScrollLeft.textContent = t.btnScrollLeft;
  if (txtBtnScrollRight) txtBtnScrollRight.textContent = t.btnScrollRight;
  if (btnJumpToNames) btnJumpToNames.textContent = t.jumpNames;
  if (btnJumpToActiveList) btnJumpToActiveList.textContent = t.jumpLoanList;
  if (btnJumpToInstitutions) btnJumpToInstitutions.textContent = t.jumpInstitutions;
  if (btnJumpToTotal) btnJumpToTotal.textContent = t.jumpTotal;
  if (txtScrollTipLabel) txtScrollTipLabel.textContent = t.scrollTip;

  // Update Empty State & App Elements
  const txtEmptyTitle = document.getElementById('txtEmptyTitle');
  if (txtEmptyTitle) txtEmptyTitle.textContent = t.emptyTitle;
  const txtEmptyDesc = document.getElementById('txtEmptyDesc');
  if (txtEmptyDesc) txtEmptyDesc.textContent = t.emptyDesc;
  const txtBtnEmptyUpload = document.getElementById('txtBtnEmptyUpload');
  if (txtBtnEmptyUpload) txtBtnEmptyUpload.textContent = t.btnEmptyUpload;
  const txtBtnDemoData = document.getElementById('txtBtnDemoData');
  if (txtBtnDemoData) txtBtnDemoData.textContent = t.btnDemoData;
  const txtEmptyHint = document.getElementById('txtEmptyHint');
  if (txtEmptyHint) txtEmptyHint.textContent = t.emptyHint;
  const txtBtnInstallApp = document.getElementById('txtBtnInstallApp');
  if (txtBtnInstallApp) txtBtnInstallApp.textContent = t.btnInstallApp;

  updateGlobalCounters();
  renderCurrentView();
}

btnLangEn.addEventListener('click', () => setLanguage('en'));
btnLangTl.addEventListener('click', () => setLanguage('tl'));
if (btnSplashLangEn) btnSplashLangEn.addEventListener('click', () => setLanguage('en'));
if (btnSplashLangTl) btnSplashLangTl.addEventListener('click', () => setLanguage('tl'));

// -----------------------------------------------------------------------------
// SPLASH & FLIP ANIMATION
// -----------------------------------------------------------------------------
btnLogin.addEventListener('click', () => {
  const startRect = splashLogo.getBoundingClientRect();
  headerBrandLogo.classList.remove('hidden');
  const targetRect = headerBrandLogo.getBoundingClientRect();

  const clone = splashLogo.cloneNode(true);
  clone.id = 'flyingLogoClone';
  clone.style.position = 'fixed';
  clone.style.left = `${startRect.left}px`;
  clone.style.top = `${startRect.top}px`;
  clone.style.width = `${startRect.width}px`;
  clone.style.height = `${startRect.height}px`;
  clone.style.zIndex = '10000';
  clone.style.transition = 'all 0.85s cubic-bezier(0.16, 1, 0.3, 1)';
  clone.style.borderRadius = '16px';
  clone.style.pointerEvents = 'none';
  document.body.appendChild(clone);

  splashLogo.style.opacity = '0';
  headerBrandLogo.style.opacity = '0';

  splashOverlay.classList.add('dissolve');
  app.classList.add('revealed');

  requestAnimationFrame(() => {
    clone.style.left = `${targetRect.left}px`;
    clone.style.top = `${targetRect.top}px`;
    clone.style.width = `${targetRect.width}px`;
    clone.style.height = `${targetRect.height}px`;
    clone.style.borderRadius = '10px';
  });

  setTimeout(() => {
    headerBrandLogo.style.opacity = '1';
    clone.remove();
    splashOverlay.classList.add('hidden');
    showToast(state.lang === 'tl' ? 'Maligayang pagdating, Ivy Salamanca! Handa na ang Sysgen RPO Loans Suite.' : 'Welcome, Ivy Salamanca! The Sysgen RPO Loans Engine is ready.');
  }, 860);
});

btnLogout.addEventListener('click', () => {
  splashOverlay.classList.remove('hidden');
  requestAnimationFrame(() => {
    splashOverlay.classList.remove('dissolve');
    splashLogo.style.opacity = '1';
  });
  showToast(state.lang === 'tl' ? 'Naka-lock na muli ang system.' : 'System re-locked.');
});

// -----------------------------------------------------------------------------
// THEME SWITCHER
// -----------------------------------------------------------------------------
const themeButtons = [
  document.getElementById('btnThemeWood'),
  document.getElementById('btnThemeBlack'),
  document.getElementById('btnThemeWhite'),
  document.getElementById('btnThemeGold')
];

function applyTheme(themeClass) {
  state.activeTheme = themeClass;
  document.body.className = themeClass;
  themeButtons.forEach(b => {
    if (b && b.dataset.theme === themeClass) b.classList.add('active');
    else if (b) b.classList.remove('active');
  });
  localStorage.setItem('sysgen_theme', themeClass);
}

themeButtons.forEach(b => {
  if (b) {
    b.addEventListener('click', () => applyTheme(b.dataset.theme));
  }
});

const savedTheme = localStorage.getItem('sysgen_theme');
if (savedTheme) applyTheme(savedTheme);

const savedLang = localStorage.getItem('sysgen_lang');
if (savedLang) setLanguage(savedLang);

// -----------------------------------------------------------------------------
// ROW INSPECTOR MODAL
// -----------------------------------------------------------------------------
function openRowModal(record) {
  state.inspectingRecord = record;
  modalRowBadge.textContent = `#${record._index + 1}`;
  modalRowTitle.textContent = `${record.fullName} - ${state.lang === 'tl' ? 'Detalye ng Loan' : 'Loan Details'}`;
  modalRowContent.innerHTML = '';

  state.headers.forEach(hKey => {
    const val = record.raw[hKey];
    const displayVal = val !== undefined && val !== null ? val : '—';

    const card = document.createElement('div');
    card.className = 'field-card';
    card.innerHTML = `
      <div>
        <span class="field-label">${escapeHtml(hKey)}</span>
        <span class="field-val ${typeof val === 'number' && val !== 0 ? 'text-gold' : ''}">${escapeHtml(String(displayVal))}</span>
      </div>
      <button class="btn-field-copy" title="Copy value">Copy</button>
    `;

    card.querySelector('.btn-field-copy').addEventListener('click', () => {
      copyToClipboard(String(displayVal), `Copied [${hKey}]: "${displayVal}"`);
    });

    modalRowContent.appendChild(card);
  });

  rowModal.classList.remove('hidden');
}

function closeRowModal() {
  rowModal.classList.add('hidden');
  state.inspectingRecord = null;
}

btnModalClose.addEventListener('click', closeRowModal);
btnModalDone.addEventListener('click', closeRowModal);
rowModal.addEventListener('click', (e) => {
  if (e.target === rowModal) closeRowModal();
});

btnModalCopyAll.addEventListener('click', () => {
  if (!state.inspectingRecord) return;
  copyToClipboard(JSON.stringify(state.inspectingRecord.raw, null, 2), `Copied JSON row for ${state.inspectingRecord.fullName}!`);
});

btnModalCopyTsv.addEventListener('click', () => {
  if (!state.inspectingRecord) return;
  const line = state.headers.map(h => state.inspectingRecord.raw[h] !== undefined ? state.inspectingRecord.raw[h] : '').join('\t');
  copyToClipboard(line, `Copied row as TSV!`);
});

// -----------------------------------------------------------------------------
// SEARCH & BATCH SELECTION
// -----------------------------------------------------------------------------
globalSearchInput.addEventListener('input', (e) => {
  state.searchQuery = e.target.value;
  btnResetSearch.classList.toggle('hidden', !state.searchQuery);
  state.currentPage = 1;
  renderCurrentView();
});

btnResetSearch.addEventListener('click', () => {
  globalSearchInput.value = '';
  state.searchQuery = '';
  btnResetSearch.classList.add('hidden');
  state.currentPage = 1;
  renderCurrentView();
});

function renderCurrentView() {
  if (state.currentView === 'overall') renderOverallSplit();
  else if (state.currentView === 'compute') computeDashboardMetrics();
  else renderMainGrid();
}

function updateSelectionBox() {
  const count = state.selectedIndices.size;
  selectedRowsCount.textContent = count;
  selectionBatchBox.classList.toggle('hidden', count === 0);
}

btnCopySelectedBatch.addEventListener('click', () => {
  const selected = Array.from(state.selectedIndices)
    .sort((a, b) => a - b)
    .map(idx => state.records[idx].raw);

  copyToClipboard(JSON.stringify(selected, null, 2), `Copied ${selected.length} selected records as JSON!`);
});

// -----------------------------------------------------------------------------
// EXPORT & COPY DROPDOWN MENU
// -----------------------------------------------------------------------------
btnExportMenu.addEventListener('click', (e) => {
  e.stopPropagation();
  exportDropdownMenu.classList.toggle('show');
});

document.addEventListener('click', () => {
  exportDropdownMenu.classList.remove('show');
});

exportDropdownMenu.addEventListener('click', (e) => {
  const btn = e.target.closest('.dropdown-item');
  if (!btn) return;
  exportDropdownMenu.classList.remove('show');
  const action = btn.dataset.action;
  const filtered = getFilteredRecordsForCurrentView();
  const isOnlyLoan = state.currentView === 'onlyLoan';
  const isFormalOnly = state.loanFilterMode === 'formal';
  const t = i18n[state.lang];

  const totalColTitle = state.lang === 'tl' ? 'KABUUANG BAWAS SA LOAN' : 'TOTAL LOAN DEDUCTION';
  const exportHeaders = isOnlyLoan ? [...state.headers, totalColTitle] : [...state.headers];

  const getRowValues = (r) => {
    const vals = state.headers.map(h => {
      const v = r.raw[h];
      if (isOnlyLoan && typeof v === 'number' && v === 0) return '';
      return v !== undefined && v !== null ? v : '';
    });
    if (isOnlyLoan) {
      const tot = isFormalOnly ? r.formalLoansTotal : r.totalAllDeductions;
      vals.push(tot);
    }
    return vals;
  };

  const getGrandTotalValues = () => {
    const footVals = state.headers.map((h, i) => {
      if (i === 0) return state.lang === 'tl' ? 'KABUUANG TOTAL' : 'GRAND TOTAL';
      if (i === 1) return `(${filtered.length} ${t.borrowersCount})`;
      const sum = filtered.reduce((acc, r) => {
        const v = r.raw[h];
        return acc + (typeof v === 'number' && v > 0 ? v : 0);
      }, 0);
      return sum > 0 ? sum : '';
    });
    if (isOnlyLoan) {
      const grandTotalAll = filtered.reduce((acc, r) => acc + (isFormalOnly ? r.formalLoansTotal : r.totalAllDeductions), 0);
      footVals.push(grandTotalAll);
    }
    return footVals;
  };

  if (action === 'copy-current-json') {
    const jsonList = filtered.map(r => {
      const obj = { ...r.raw };
      if (isOnlyLoan) {
        obj[totalColTitle] = isFormalOnly ? r.formalLoansTotal : r.totalAllDeductions;
      }
      return obj;
    });
    if (isOnlyLoan) {
      const footObj = {
        [state.headers[0] || 'Last Name']: state.lang === 'tl' ? 'KABUUANG TOTAL' : 'GRAND TOTAL',
        [state.headers[1] || 'First Name']: `(${filtered.length} ${t.borrowersCount})`
      };
      state.headers.slice(2).forEach(h => {
        const sum = filtered.reduce((acc, r) => acc + (typeof r.raw[h] === 'number' && r.raw[h] > 0 ? r.raw[h] : 0), 0);
        if (sum > 0) footObj[h] = sum;
      });
      footObj[totalColTitle] = filtered.reduce((acc, r) => acc + (isFormalOnly ? r.formalLoansTotal : r.totalAllDeductions), 0);
      jsonList.push(footObj);
    }
    copyToClipboard(JSON.stringify(jsonList, null, 2), `Copied ${filtered.length} rows + Grand Total as JSON!`);
  } else if (action === 'copy-current-tsv') {
    const headerLine = exportHeaders.join('\t');
    const dataLines = filtered.map(r => getRowValues(r).join('\t'));
    const footLine = isOnlyLoan ? getGrandTotalValues().join('\t') : null;
    const allLines = footLine ? [headerLine, ...dataLines, footLine] : [headerLine, ...dataLines];
    copyToClipboard(allLines.join('\n'), `Copied ${filtered.length} rows + Grand Total as TSV (Ready to paste in Excel)!`);
  } else if (action === 'copy-current-csv') {
    const escCsv = (v) => `"${String(v !== undefined && v !== null ? v : '').replace(/"/g, '""')}"`;
    const headerLine = exportHeaders.map(escCsv).join(',');
    const dataLines = filtered.map(r => getRowValues(r).map(escCsv).join(','));
    const footLine = isOnlyLoan ? getGrandTotalValues().map(escCsv).join(',') : null;
    const allLines = footLine ? [headerLine, ...dataLines, footLine] : [headerLine, ...dataLines];
    copyToClipboard(allLines.join('\n'), `Copied ${filtered.length} rows + Grand Total as CSV!`);
  } else if (action === 'copy-current-md') {
    const headerLine = '| ' + exportHeaders.join(' | ') + ' |';
    const sep = '| ' + exportHeaders.map(() => '---').join(' | ') + ' |';
    const dataLines = filtered.map(r => '| ' + getRowValues(r).map(v => String(v).replace(/\|/g, '\\|')).join(' | ') + ' |');
    const footLine = isOnlyLoan ? '| ' + getGrandTotalValues().map(v => String(v).replace(/\|/g, '\\|')).join(' | ') + ' |' : null;
    const allLines = footLine ? [headerLine, sep, ...dataLines, footLine] : [headerLine, sep, ...dataLines];
    copyToClipboard(allLines.join('\n'), `Copied Markdown Table + Grand Total!`);
  }
});

// File Upload & Dynamic Parsing Handler
function handleUploadedFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheet = workbook.SheetNames[0];
      const jsonRows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: 0 });
      if (!jsonRows || jsonRows.length === 0) {
        showToast(state.lang === 'tl' ? 'Walang laman ang napiling Excel sheet!' : 'The selected Excel sheet has no rows!');
        return;
      }
      initDataset(jsonRows);
      showToast(state.lang === 'tl'
        ? `Matagumpay na nabasa ang ${file.name} (${jsonRows.length} talaan)!`
        : `Successfully parsed ${file.name} (${jsonRows.length} rows)!`);
    } catch (err) {
      alert('Error parsing Excel file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

fileInputGlobal.addEventListener('change', (e) => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  handleUploadedFile(file);
});

if (btnEmptyUpload) {
  btnEmptyUpload.addEventListener('click', () => {
    fileInputGlobal.click();
  });
}

if (btnLoadDemoData) {
  btnLoadDemoData.addEventListener('click', () => {
    initDataset(sysgenInitialLoanData);
    showToast(state.lang === 'tl' ? 'Naka-load ang Sysgen Sample Demo Data (269 talaan)!' : 'Loaded Sysgen Sample Demo Data (269 rows)!');
  });
}

// Drag & Drop on Empty State & Window
if (emptyDropzoneState) {
  ['dragenter', 'dragover'].forEach(eventName => {
    emptyDropzoneState.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      emptyDropzoneState.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    emptyDropzoneState.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      emptyDropzoneState.classList.remove('dragover');
    });
  });

  emptyDropzoneState.addEventListener('drop', (e) => {
    const files = e.dataTransfer && e.dataTransfer.files;
    if (files && files.length > 0) {
      handleUploadedFile(files[0]);
    }
  });
}

// Global Drag & Drop over workspace window
window.addEventListener('dragover', (e) => {
  e.preventDefault();
});

window.addEventListener('drop', (e) => {
  e.preventDefault();
  const files = e.dataTransfer && e.dataTransfer.files;
  if (files && files.length > 0) {
    const ext = files[0].name.toLowerCase();
    if (ext.endsWith('.xlsx') || ext.endsWith('.xls') || ext.endsWith('.csv')) {
      handleUploadedFile(files[0]);
    }
  }
});

// PWA Service Worker & Desktop App Install Support
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').catch(() => {});
  });
}

let deferredInstallPrompt = null;
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  if (btnInstallApp) btnInstallApp.classList.remove('hidden');
});

if (btnInstallApp) {
  btnInstallApp.addEventListener('click', async () => {
    if (deferredInstallPrompt) {
      deferredInstallPrompt.prompt();
      const choice = await deferredInstallPrompt.userChoice;
      if (choice && choice.outcome === 'accepted') {
        showToast(state.lang === 'tl' ? 'Matagumpay na na-install ang Sysgen Loans App!' : 'Sysgen Loans App successfully installed!');
        btnInstallApp.classList.add('hidden');
      }
      deferredInstallPrompt = null;
    } else {
      showToast(state.lang === 'tl' ? 'Maaaring i-install sa pamamagitan ng browser menu (Install App)' : 'Install available via browser address bar or menu (Install App)');
    }
  });
}

// Page size select
pageSizeSelect.addEventListener('change', (e) => {
  state.pageSize = e.target.value;
  state.currentPage = 1;
  renderMainGrid();
});

btnPrev.addEventListener('click', () => {
  if (state.currentPage > 1) {
    state.currentPage--;
    renderMainGrid();
  }
});

btnNext.addEventListener('click', () => {
  state.currentPage++;
  renderMainGrid();
});

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// -----------------------------------------------------------------------------
// TABLE QUICK-SCROLL NAVIGATION & INTERACTIVE MINI-TRACK
// -----------------------------------------------------------------------------
function setupTableScrollHelper() {
  if (!mainTableScrollContainer) return;

  if (btnScrollStepLeft) {
    btnScrollStepLeft.addEventListener('click', () => {
      mainTableScrollContainer.scrollBy({ left: -360, behavior: 'smooth' });
    });
  }

  if (btnScrollStepRight) {
    btnScrollStepRight.addEventListener('click', () => {
      mainTableScrollContainer.scrollBy({ left: 360, behavior: 'smooth' });
    });
  }

  if (btnJumpToNames) {
    btnJumpToNames.addEventListener('click', () => {
      mainTableScrollContainer.scrollTo({ left: 0, behavior: 'smooth' });
    });
  }

  if (btnJumpToActiveList) {
    btnJumpToActiveList.addEventListener('click', () => {
      mainTableScrollContainer.scrollTo({ left: 240, behavior: 'smooth' });
    });
  }

  if (btnJumpToInstitutions) {
    btnJumpToInstitutions.addEventListener('click', () => {
      mainTableScrollContainer.scrollTo({ left: 620, behavior: 'smooth' });
    });
  }

  if (btnJumpToTotal) {
    btnJumpToTotal.addEventListener('click', () => {
      mainTableScrollContainer.scrollTo({ left: mainTableScrollContainer.scrollWidth, behavior: 'smooth' });
    });
  }

  function updateScrollIndicator() {
    if (!scrollMiniThumb) return;
    const maxScroll = mainTableScrollContainer.scrollWidth - mainTableScrollContainer.clientWidth;
    if (maxScroll <= 0) {
      scrollMiniThumb.style.width = '100%';
      scrollMiniThumb.style.left = '0%';
      return;
    }
    const ratio = mainTableScrollContainer.scrollLeft / maxScroll;
    const thumbWidthPercent = Math.max(16, Math.min(45, (mainTableScrollContainer.clientWidth / mainTableScrollContainer.scrollWidth) * 100));
    scrollMiniThumb.style.width = `${thumbWidthPercent}%`;
    const availableTrackPercent = 100 - thumbWidthPercent;
    scrollMiniThumb.style.left = `${ratio * availableTrackPercent}%`;
  }

  mainTableScrollContainer.addEventListener('scroll', updateScrollIndicator);
  window.addEventListener('resize', updateScrollIndicator);

  if (scrollMiniIndicator) {
    scrollMiniIndicator.addEventListener('click', (e) => {
      const rect = scrollMiniIndicator.getBoundingClientRect();
      const clickRatio = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
      const maxScroll = mainTableScrollContainer.scrollWidth - mainTableScrollContainer.clientWidth;
      mainTableScrollContainer.scrollTo({ left: clickRatio * maxScroll, behavior: 'smooth' });
    });
  }

  // Shift + Mouse Wheel smooth horizontal panning
  mainTableScrollContainer.addEventListener('wheel', (e) => {
    if (e.shiftKey) {
      e.preventDefault();
      mainTableScrollContainer.scrollLeft += e.deltaY;
    }
  }, { passive: false });

  setTimeout(updateScrollIndicator, 150);
}

setupTableScrollHelper();

// Initialize dataset: Starts blank on launch as requested!
initDataset([]);
