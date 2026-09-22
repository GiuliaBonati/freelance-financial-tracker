// ============================================================
// FINANCIAL TRACKER — MASTER AUTOMATION
// ============================================================

function onEdit(e) {
  
  updateGuideLastChecked_(e);
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const column = e.range.getColumn();

    // ==========================================================
  // ACCOUNTS — LAST UPDATED
  // When Balance (C) changes, update Last Updated (F)
  // ==========================================================

  if (
    sheetName === "Accounts" &&
    row >= 2 &&
    column === 3
  ) {
    sheet
      .getRange(row, 6)
      .setValue(new Date())
      .setNumberFormat("dd/MM/yyyy");

    return;
  }

  // ==========================================================
  // 1. INVOICES → INCOME
  // ==========================================================

  if (sheetName === "Invoices") {
    syncInvoicesToIncome();
    return;
  }

  // ==========================================================
  // 2. FIXED COSTS → EXPENSES
  // Column L = Log Expense checkbox
  // ==========================================================

  if (
    sheetName === "Fixed Costs" &&
    row >= 2 &&
    column === 12 &&
    e.value === "TRUE"
  ) {
    logFixedCostToExpenses_(e);
    return;
  }

  // ==========================================================
  // 3. DEPENDENT DROPDOWNS
  // ==========================================================

  const config = {
    "Expenses": {
      categoryColumn: 3,
      subcategoryColumn: 4
    },
    "Fixed Costs": {
      categoryColumn: 2,
      subcategoryColumn: 3
    }
  };

  if (!config[sheetName]) return;

  if (
    row < 2 ||
    column !== config[sheetName].categoryColumn
  ) {
    return;
  }

  updateDependentDropdown_(
    e,
    config[sheetName]
  );
}


// ============================================================
// DEPENDENT DROPDOWNS
// Category → Subcategory
// ============================================================

function updateDependentDropdown_(e, config) {

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const category = e.range.getValue();

  const subcategoryCell =
    sheet.getRange(
      row,
      config.subcategoryColumn
    );

  // Clear previous subcategory
  subcategoryCell.clearContent();
  subcategoryCell.clearDataValidations();

  if (!category) return;

  const settingsSheet =
    e.source.getSheetByName("Settings");

  if (!settingsSheet) return;

  const lastRow =
    settingsSheet.getLastRow();

  if (lastRow < 2) return;

  const mapping =
    settingsSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();

  const subcategories =
    mapping
      .filter(row =>
        row[0] === category &&
        row[1] !== ""
      )
      .map(row => row[1]);

  if (subcategories.length === 0) return;

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        subcategories,
        true
      )
      .setAllowInvalid(false)
      .build();

  subcategoryCell
    .setDataValidation(rule);
}


// ============================================================
// FIXED COSTS → EXPENSES
// ============================================================

function logFixedCostToExpenses_(e) {

  const ss = e.source;

  const fixedSheet =
    e.range.getSheet();

  const expensesSheet =
    ss.getSheetByName("Expenses");

  if (!expensesSheet) {
    throw new Error(
      'Sheet "Expenses" not found.'
    );
  }

  const row =
    e.range.getRow();

  // Fixed Costs structure after adding Log Expense:
  //
  // A Expense
  // B Category
  // C Subcategory
  // D Amount
  // E Frequency
  // F Monthly Equivalent
  // G Annual Cost
  // H Payment Method
  // I Account
  // J Next Payment
  // K Active
  // L Log Expense
  // M Notes

  const data =
    fixedSheet
      .getRange(
        row,
        1,
        1,
        13
      )
      .getValues()[0];

  const description   = data[0];
  const category      = data[1];
  const subcategory   = data[2];
  const amount        = data[3];
  const frequency     = data[4];
  const paymentMethod = data[7];
  const account       = data[8];
  const nextPayment   = data[9];
  const active        = data[10];

  const checkbox =
    fixedSheet.getRange(
      row,
      12
    );


  // ==========================================================
  // VALIDATION
  // ==========================================================

  const missingRequiredData =
    !description ||
    !category ||
    !subcategory ||
    amount === "" ||
    amount === null ||
    !frequency ||
    !paymentMethod ||
    !account ||
    !nextPayment;

  if (
    active !== "Yes" ||
    missingRequiredData
  ) {

    checkbox.setValue(false);

    ss.toast(
      "Complete all required fields and make sure Active is Yes.",
      "Expense not logged",
      5
    );

    return;
  }


  // ==========================================================
  // DUPLICATE PROTECTION
  //
  // Same:
  // Date + Description + Amount + Expense Type
  // ==========================================================

  const lastExpenseRow =
    expensesSheet.getLastRow();

  if (lastExpenseRow >= 2) {

    const expenseData =
      expensesSheet
        .getRange(
          2,
          1,
          lastExpenseRow - 1,
          8
        )
        .getValues();

    const paymentTime =
      normalizeDate_(nextPayment);

    const duplicate =
      expenseData.some(expense => {

        const expenseDate =
          expense[0];

        if (!expenseDate) return false;

        const expenseTime =
          normalizeDate_(expenseDate);

        const expenseDescription =
          expense[1];

        const expenseAmount =
          expense[4];

        const expenseType =
          expense[5];

        return (
          expenseTime === paymentTime &&
          expenseDescription === description &&
          Number(expenseAmount) === Number(amount) &&
          expenseType === "Fixed"
        );
      });

    if (duplicate) {

      checkbox.setValue(false);

      ss.toast(
        "This fixed expense is already recorded in Expenses.",
        "Duplicate prevented",
        5
      );

      return;
    }
  }


  // ==========================================================
  // FIND FIRST EMPTY ROW IN EXPENSES
  // ==========================================================

  const targetRow =
    findFirstEmptyRow_(
      expensesSheet,
      1,
      2
    );


 

// ==========================================================
// WRITE EXPENSE
// ==========================================================

// Remove any existing validations from the target cells.
// The row is generated automatically, so validation is
// not required for the imported values.
expensesSheet
  .getRange(targetRow, 1, 1, 8)
  .clearDataValidations();

expensesSheet
  .getRange(targetRow, 1, 1, 8)
  .setValues([[
    nextPayment,
    description,
    category,
    subcategory,
    amount,
    "Fixed",
    paymentMethod,
    account
  ]]);

// ==========================================================
// RESTORE DROPDOWNS ON GENERATED EXPENSE ROW
// ==========================================================

restoreExpenseValidations_(
  ss,
  expensesSheet,
  targetRow,
  category
);
  // ==========================================================
  // RESTORE EXPENSE HELPER FORMULAS
  // ==========================================================

  restoreExpenseFormulas_(
    expensesSheet
  );

  SpreadsheetApp.flush();


  // ==========================================================
  // CALCULATE NEXT PAYMENT
  // ==========================================================

  const newDate =
    calculateNextPayment_(
      nextPayment,
      frequency
    );

  if (!newDate) {

    // Remove the Expense we just created,
    // because Frequency was invalid.
    expensesSheet
      .getRange(
        targetRow,
        1,
        1,
        8
      )
      .clearContent();

    checkbox.setValue(false);

    ss.toast(
      "Frequency not recognized. Expense was not logged.",
      "Check Fixed Costs",
      5
    );

    return;
  }


  // ==========================================================
  // UPDATE FIXED COST
  // ==========================================================

  fixedSheet
    .getRange(
      row,
      10
    )
    .setValue(newDate);

  checkbox.setValue(false);

  SpreadsheetApp.flush();

  ss.toast(
    "Fixed expense logged successfully.",
    "Expenses updated",
    4
  );
}


// ============================================================
// CALCULATE NEXT PAYMENT
// ============================================================

function calculateNextPayment_(
  currentDate,
  frequency
) {

  const newDate =
    new Date(currentDate);

  if (
    Number.isNaN(
      newDate.getTime()
    )
  ) {
    return null;
  }

  switch (frequency) {

    case "Monthly":
      newDate.setMonth(
        newDate.getMonth() + 1
      );
      break;

    case "Quarterly":
      newDate.setMonth(
        newDate.getMonth() + 3
      );
      break;

    case "Yearly":
      newDate.setFullYear(
        newDate.getFullYear() + 1
      );
      break;

    default:
      return null;
  }

  return newDate;
}


// ============================================================
// NORMALIZE DATE
// Used for duplicate checking
// ============================================================

function normalizeDate_(date) {

  const d =
    new Date(date);

  d.setHours(
    0,
    0,
    0,
    0
  );

  return d.getTime();
}


// ============================================================
// FIND FIRST EMPTY ROW
// ============================================================

function findFirstEmptyRow_(
  sheet,
  column,
  startRow
) {

  const maxRows =
    sheet.getMaxRows();

  const values =
    sheet
      .getRange(
        startRow,
        column,
        maxRows - startRow + 1,
        1
      )
      .getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (values[i][0] === "") {
      return startRow + i;
    }
  }

  sheet.insertRowAfter(maxRows);

  return maxRows + 1;
}


// ============================================================
// EXPENSE HELPER FORMULAS
// I = Month
// J = Month Number
// K = Year
// ============================================================

function restoreExpenseFormulas_(
  expensesSheet
) {

  expensesSheet
    .getRange("I2")
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="";"";TEXT(A2:A;"mmmm")))'
    );

  expensesSheet
    .getRange("J2")
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="";"";MONTH(A2:A)))'
    );

  expensesSheet
    .getRange("K2")
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="";"";YEAR(A2:A)))'
    );
}


// ============================================================
// INVOICES → INCOME
// ============================================================

function syncInvoicesToIncome() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const invoiceSheet =
    ss.getSheetByName("Invoices");

  const incomeSheet =
    ss.getSheetByName("Income");

  if (
    !invoiceSheet ||
    !incomeSheet
  ) {
    return;
  }

  const invoiceLastRow =
    invoiceSheet.getLastRow();

  // Paid invoices currently present
  const paidInvoices =
    new Map();


  // ==========================================================
  // READ PAID INVOICES
  // ==========================================================

  if (invoiceLastRow >= 2) {

    const invoiceData =
      invoiceSheet
        .getRange(
          2,
          1,
          invoiceLastRow - 1,
          9
        )
        .getValues();

    invoiceData.forEach(row => {

      const invoiceId =
        row[0]; // A

      const client =
        row[2]; // C

      const project =
        row[3]; // D

      const service =
        row[4]; // E

      const channel =
        row[5]; // F

      const grossAmount =
        row[6]; // G

      const paymentDate =
        row[8]; // I

      if (
        invoiceId &&
        paymentDate
      ) {

        paidInvoices.set(
          String(invoiceId),
          [
            paymentDate,
            client,
            project,
            service,
            channel,
            grossAmount,
            invoiceId,
            "Paid"
          ]
        );
      }
    });
  }


  // ==========================================================
  // FIND LAST USED INCOME ROW
  // ==========================================================

  const incomeColumnA =
    incomeSheet
      .getRange("A:A")
      .getValues();

  let incomeLastRow = 1;

  incomeColumnA.forEach(
    (row, index) => {

      if (row[0] !== "") {
        incomeLastRow =
          index + 1;
      }
    }
  );


  // ==========================================================
  // EXISTING INVOICE ROWS
  // ==========================================================

  const existingInvoiceRows =
    new Map();

  const rowsToClear =
    [];


  if (incomeLastRow >= 2) {

    const incomeData =
      incomeSheet
        .getRange(
          2,
          1,
          incomeLastRow - 1,
          8
        )
        .getValues();

    incomeData.forEach(
      (row, index) => {

        const sheetRow =
          index + 2;

        const invoiceId =
          row[6]; // G

        // Manual/non-invoice income
        if (!invoiceId) return;

        const id =
          String(invoiceId);


        // Duplicate Invoice ID
        if (
          existingInvoiceRows.has(id)
        ) {

          rowsToClear.push(
            sheetRow
          );

          return;
        }

        existingInvoiceRows.set(
          id,
          sheetRow
        );


        // Invoice no longer paid / removed
        if (
          !paidInvoices.has(id)
        ) {

          rowsToClear.push(
            sheetRow
          );

          return;
        }


        // Update existing synced row
        incomeSheet
          .getRange(
            sheetRow,
            1,
            1,
            8
          )
          .setValues([[
            ...paidInvoices.get(id)
          ]]);

        paidInvoices.delete(id);
      }
    );
  }


  // ==========================================================
  // CLEAR OBSOLETE / DUPLICATE INVOICE ROWS
  // ==========================================================

  rowsToClear.forEach(row => {

    incomeSheet
      .getRange(
        row,
        1,
        1,
        8
      )
      .clearContent();
  });


  // ==========================================================
  // CREATE NEW PAID INVOICES
  // ==========================================================

 paidInvoices.forEach(values => {

  const targetRow =
    findFirstCompletelyEmptyIncomeRow_(
      incomeSheet
    );

  incomeSheet
    .getRange(
      targetRow,
      1,
      1,
      8
    )
    .setValues([
      values
    ]);
});

  // ==========================================================
  // RESTORE INCOME HELPER FORMULAS
  // ==========================================================

  restoreIncomeFormulas_(
    incomeSheet
  );

  SpreadsheetApp.flush();
}


// ============================================================
// INCOME HELPER FORMULAS
// I = Month
// J = Month Number
// K = Year
// ============================================================
function findFirstCompletelyEmptyIncomeRow_(
  incomeSheet
) {

  const startRow = 2;
  const maxRows = incomeSheet.getMaxRows();

  const values =
    incomeSheet
      .getRange(
        startRow,
        1,
        maxRows - startRow + 1,
        8
      )
      .getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const rowIsEmpty =
      values[i].every(
        value => value === ""
      );

    if (rowIsEmpty) {
      return startRow + i;
    }
  }

  incomeSheet.insertRowAfter(maxRows);

  return maxRows + 1;
}

function restoreIncomeFormulas_(
  incomeSheet
) {

  incomeSheet
    .getRange("I2")
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="";"";TEXT(A2:A;"mmmm")))'
    );

  incomeSheet
    .getRange("J2")
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="";"";MONTH(A2:A)))'
    );

  incomeSheet
    .getRange("K2")
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="";"";YEAR(A2:A)))'
    );
}

function restoreExpenseValidations_(
  ss,
  expensesSheet,
  targetRow,
  category
) {

  const settingsSheet =
    ss.getSheetByName("Settings");

  if (!settingsSheet) return;

  const lastRow =
    settingsSheet.getLastRow();

  // ========================================
  // READ SETTINGS
  // ========================================

  const settingsData =
    settingsSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();

  // Unique Expense Categories — Settings A
  const categories = [
    ...new Set(
      settingsData
        .map(row => row[0])
        .filter(Boolean)
    )
  ];

  // Subcategories belonging to current Category
  // Settings A:B
  const subcategories =
    settingsData
      .filter(row =>
        row[0] === category &&
        row[1]
      )
      .map(row => row[1]);

  // Accounts — Settings E
  const accounts = [
    ...new Set(
      settingsData
        .map(row => row[4])
        .filter(Boolean)
    )
  ];

  // Payment Methods — Settings F
  const paymentMethods = [
    ...new Set(
      settingsData
        .map(row => row[5])
        .filter(Boolean)
    )
  ];


  // ========================================
  // CATEGORY — C
  // ========================================

  if (categories.length) {

    const rule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          categories,
          true
        )
        .setAllowInvalid(false)
        .build();

    expensesSheet
      .getRange(targetRow, 3)
      .setDataValidation(rule);
  }


  // ========================================
  // SUBCATEGORY — D
  // ========================================

  if (subcategories.length) {

    const rule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          subcategories,
          true
        )
        .setAllowInvalid(false)
        .build();

    expensesSheet
      .getRange(targetRow, 4)
      .setDataValidation(rule);
  }


  // ========================================
  // EXPENSE TYPE — F
  // ========================================

  const expenseTypeRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        ["Variable", "Fixed"],
        true
      )
      .setAllowInvalid(false)
      .build();

  expensesSheet
    .getRange(targetRow, 6)
    .setDataValidation(
      expenseTypeRule
    );


  // ========================================
  // PAYMENT METHOD — G
  // ========================================

  if (paymentMethods.length) {

    const rule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          paymentMethods,
          true
        )
        .setAllowInvalid(false)
        .build();

    expensesSheet
      .getRange(targetRow, 7)
      .setDataValidation(rule);
  }


  // ========================================
  // ACCOUNT — H
  // ========================================

  if (accounts.length) {

    const rule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          accounts,
          true
        )
        .setAllowInvalid(false)
        .build();

    expensesSheet
      .getRange(targetRow, 8)
      .setDataValidation(rule);
  }
}

function updateGuideLastChecked_(e) {
  const sheet = e.range.getSheet();

  // Run only in the Guide sheet
  if (sheet.getName() !== 'Guide') return;

  const row = e.range.getRow();
  const column = e.range.getColumn();

  // DONE = column H
  if (column !== 8) return;

  // Only react when a checkbox is checked
  if (e.value !== 'TRUE') return;

  // Make sure this is actually a checklist row:
  // column B contains the item number
  const itemNumber = sheet.getRange(row, 2).getValue();
  if (itemNumber === '' || isNaN(Number(itemNumber))) return;

  // LAST UPDATED = column J
  const lastUpdatedCell = sheet.getRange(row, 10);

  lastUpdatedCell.setValue(new Date());
  lastUpdatedCell.setNumberFormat('dd/MM/yyyy');
}

function resetWeeklyFinanceCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Guide');

  if (!sheet) {
    throw new Error('Guide sheet not found.');
  }

  // DONE column = H
  // Finds the checklist rows dynamically using the numbers in column B
  const lastRow = sheet.getLastRow();

  for (let row = 1; row <= lastRow; row++) {
    const itemNumber = sheet.getRange(row, 2).getValue();

    if (itemNumber !== '' && !isNaN(Number(itemNumber))) {
      sheet.getRange(row, 8).setValue(false);
    }
  }
}
