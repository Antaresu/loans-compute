# Sysgen RPO - Loans Intelligence & Computation Suite

A state-of-the-art internal payroll loan computation, analytics, and Excel intelligence desktop & web application engineered for **Sysgen RPO**.

---

## 👥 Project Credits & Authorization
* **Authorized Personnel:** Ivy Salamanca *(Loans & Payroll Computation Division)*
* **Engineered & Developed by:** Christian Antero Salamanca

---

## ✨ Key Features

1. **Dynamic Excel (.xlsx) Parser & Ingestion Engine**
   * Upload any payroll spreadsheet cutoff — columns, employee names, and deduction headers are automatically normalized.
   * Auto-identifies and calculates all formal loans and charges (*SSS Salary Loan, SSS Calamity, HDMF / Pag-IBIG Loan, HDMF MP2, AUB Loan, Cash Advances, HMO, and custom deductions*).

2. **Accurate "Only Loans" View & Zero Suppression**
   * Automatically filters out accounts with zero loan balances so payroll officers only work with active borrowers.
   * Zero balances are cleanly suppressed with subtle accounting dashes.
   * Appends an exact **TOTAL LOAN DEDUCTION** column for every employee.

3. **Executive Accounting Excel (.xlsx) Export**
   * Exports in clean Black & White corporate formatting with a Gold accent line.
   * **Frozen Panes:** Row 1 header locked at top; Column A & B locked on horizontal pan.
   * **Accounting Grand Total:** Classic double bottom border underline on the total deduction row.
   * Dynamic auto-fit column widths so headers and numbers are never clipped.

4. **Multi-Format Fast Copy Tools**
   * One-click copying for visible records as TSV *(instant paste to Excel/Google Sheets)*, JSON, CSV, or Markdown.
   * Individual row inspection modal with cell-level copy buttons.

5. **Interactive Amortization Calculator**
   * Real-time principal, interest, term, monthly, and semi-monthly (15/30 cutoff) calculation simulator.

6. **Bilingual & Multi-Theme UI**
   * 100% full English and Tagalog (Filipino) language support.
   * Luxury Wood, Obsidian Black, Executive White, and Imperial Gold themes.

7. **Application Packaging**
   * **PWA:** Installable as a native desktop application directly from Chrome/Edge with one click.
   * **Electron:** Ready for standalone Windows desktop `.exe` execution.

---

## 🚀 Getting Started

### Prerequisites
* [Node.js](https://nodejs.org/) (version 18 or higher recommended)
* npm

### Installation
```bash
# Clone the repository
git clone https://github.com/Antaresu/loans-compute.git

# Enter project directory
cd loans-compute

# Install dependencies
npm install
```

### Running Locally
```bash
# Start Vite development server
npm run dev
```
Open your browser at `http://localhost:5173/`.

### Building for Production
```bash
npm run build
```

---

## 🔒 Confidentiality & Usage Notice
Strictly for Sysgen RPO internal loan computations and authorized personnel. All formulas, deductions, and payroll records are proprietary and confidential.
