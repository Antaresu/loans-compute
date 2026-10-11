import * as XLSX from 'xlsx';
import { getSampleLoanWorkbookData } from './sample-data.js';

const sample = getSampleLoanWorkbookData();
const wb = XLSX.utils.book_new();

// Sheet 1: Loans
const ws1 = XLSX.utils.aoa_to_sheet(sample.data);
XLSX.utils.book_append_sheet(wb, ws1, "Active Loans");

// Sheet 2: Rate Table
const ws2 = XLSX.utils.aoa_to_sheet([
  ["Tier Code", "Loan Type", "Interest Rate", "Max Term Months"],
  ["T1", "Salary Loan", 0.015, 36],
  ["T2", "Emergency Loan", 0.010, 12],
  ["T3", "Educational Loan", 0.012, 24],
  ["T4", "Appliance Loan", 0.018, 12],
  ["T5", "Calamity Loan", 0.012, 18]
]);
XLSX.utils.book_append_sheet(wb, ws2, "Interest Computation Rates");

const outputPath = './sample_loans.xlsx';
XLSX.writeFile(wb, outputPath);
console.log(`Successfully generated demo Excel file at: ${outputPath}`);
