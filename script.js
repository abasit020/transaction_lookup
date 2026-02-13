/*
  Transaction Lookup Tool
  -----------------------
  This script runs fully in the browser and:
  1) Reads first-sheet data from two uploaded .xlsx files (SheetJS)
  2) Lets users map lookup and amount columns
  3) Matches Sales rows to Account rows by lookup value
  4) Groups matched transactions by Account Number and computes totals
*/

const state = {
  salesRows: [],
  accountRows: [],
  salesHeaders: [],
  accountHeaders: []
};

const salesFileInput = document.getElementById('salesFile');
const accountFileInput = document.getElementById('accountFile');
const salesStatus = document.getElementById('salesStatus');
const accountStatus = document.getElementById('accountStatus');

const salesLookupColumn = document.getElementById('salesLookupColumn');
const accountLookupColumn = document.getElementById('accountLookupColumn');
const salesAmountColumn = document.getElementById('salesAmountColumn');

const processBtn = document.getElementById('processBtn');
const messageEl = document.getElementById('message');

const resultsSection = document.getElementById('resultsSection');
const resultsBody = document.getElementById('resultsBody');
const grandCount = document.getElementById('grandCount');
const grandAmount = document.getElementById('grandAmount');

salesFileInput.addEventListener('change', async (event) => {
  await handleFileLoad(event.target.files[0], 'sales');
});

accountFileInput.addEventListener('change', async (event) => {
  await handleFileLoad(event.target.files[0], 'account');
});

processBtn.addEventListener('click', processTransactions);

async function handleFileLoad(file, type) {
  clearMessage();

  if (!file) {
    return;
  }

  if (!window.XLSX) {
    setMessage('SheetJS failed to load. Please check your internet connection and reload.', true);
    return;
  }

  try {
    const rows = await readFirstSheet(file);

    if (!rows.length) {
      throw new Error('The first sheet has no data rows.');
    }

    const headers = Object.keys(rows[0]);
    if (!headers.length) {
      throw new Error('Could not detect headers. Ensure the first row contains column names.');
    }

    if (type === 'sales') {
      state.salesRows = rows;
      state.salesHeaders = headers;
      salesStatus.textContent = `Loaded ${file.name} (${rows.length} rows)`;
      populateSelect(salesLookupColumn, headers, 'Select Sales lookup column');
      populateSelect(salesAmountColumn, headers, 'Select Sales amount column');
    } else {
      state.accountRows = rows;
      state.accountHeaders = headers;
      accountStatus.textContent = `Loaded ${file.name} (${rows.length} rows)`;
      populateSelect(accountLookupColumn, headers, 'Select Account lookup column');
    }

    toggleProcessButton();
  } catch (error) {
    setMessage(`${type === 'sales' ? 'Sales' : 'Account'} file error: ${error.message}`, true);
  }
}

function readFirstSheet(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];

        // defval keeps empty cells present in each row object
        const jsonRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
        resolve(jsonRows);
      } catch (error) {
        reject(new Error('Unable to parse the Excel file.'));
      }
    };

    reader.onerror = () => reject(new Error('Failed to read the uploaded file.'));
    reader.readAsArrayBuffer(file);
  });
}

function populateSelect(selectEl, options, placeholder) {
  selectEl.innerHTML = '';

  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.textContent = placeholder;
  selectEl.appendChild(placeholderOption);

  options.forEach((header) => {
    const option = document.createElement('option');
    option.value = header;
    option.textContent = header;
    selectEl.appendChild(option);
  });

  selectEl.disabled = false;
}

function toggleProcessButton() {
  const ready =
    state.salesRows.length > 0 &&
    state.accountRows.length > 0 &&
    !salesLookupColumn.disabled &&
    !accountLookupColumn.disabled &&
    !salesAmountColumn.disabled;

  processBtn.disabled = !ready;
}

function processTransactions() {
  clearMessage();
  resultsBody.innerHTML = '';
  resultsSection.hidden = true;

  const salesLookup = salesLookupColumn.value;
  const accountLookup = accountLookupColumn.value;
  const amountColumn = salesAmountColumn.value;

  if (!salesLookup || !accountLookup || !amountColumn) {
    setMessage('Please select all required columns before processing.', true);
    return;
  }

  // Try to identify a dedicated account number column from the account file.
  // Fallback to selected account lookup column if not found.
  const detectedAccountNumberColumn =
    state.accountHeaders.find((header) => /account\s*number/i.test(header)) ||
    state.accountHeaders.find((header) => /acct\s*number/i.test(header)) ||
    state.accountHeaders.find((header) => /^account$/i.test(header)) ||
    accountLookup;

  // Build quick lookup map: account lookup value -> account number
  const lookupToAccountNumber = new Map();
  state.accountRows.forEach((row) => {
    const lookupValue = normalizeValue(row[accountLookup]);
    const accountNumber = normalizeValue(row[detectedAccountNumberColumn]);

    if (lookupValue) {
      lookupToAccountNumber.set(lookupValue, accountNumber || lookupValue);
    }
  });

  // Aggregate by account number
  const aggregate = new Map();
  let totalCount = 0;
  let totalAmount = 0;

  state.salesRows.forEach((row) => {
    const salesLookupValue = normalizeValue(row[salesLookup]);
    if (!salesLookupValue) {
      return;
    }

    const accountNumber = lookupToAccountNumber.get(salesLookupValue);
    if (!accountNumber) {
      return;
    }

    const amount = parseAmount(row[amountColumn]);
    if (!aggregate.has(accountNumber)) {
      aggregate.set(accountNumber, { count: 0, amount: 0 });
    }

    const current = aggregate.get(accountNumber);
    current.count += 1;
    current.amount += amount;

    totalCount += 1;
    totalAmount += amount;
  });

  if (aggregate.size === 0) {
    setMessage('No matching records found for the selected lookup columns.', true);
    return;
  }

  const sortedRows = [...aggregate.entries()].sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  sortedRows.forEach(([accountNumber, summary]) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(accountNumber)}</td>
      <td>${summary.count}</td>
      <td>${formatCurrency(summary.amount)}</td>
    `;
    resultsBody.appendChild(tr);
  });

  grandCount.textContent = totalCount.toLocaleString();
  grandAmount.textContent = formatCurrency(totalAmount);
  resultsSection.hidden = false;

  setMessage(
    `Processed ${totalCount.toLocaleString()} matched transactions across ${aggregate.size.toLocaleString()} accounts.`,
    false,
    true
  );
}

function normalizeValue(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function parseAmount(value) {
  if (typeof value === 'number') {
    return value;
  }

  // Remove commas, spaces, and currency symbols for robust parsing
  const normalized = String(value || '')
    .replace(/[$,\s]/g, '')
    .trim();

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatCurrency(value) {
  return value.toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function setMessage(text, isError = false, isSuccess = false) {
  messageEl.textContent = text;
  messageEl.classList.remove('error', 'success');

  if (isError) {
    messageEl.classList.add('error');
  } else if (isSuccess) {
    messageEl.classList.add('success');
  }
}

function clearMessage() {
  setMessage('');
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
