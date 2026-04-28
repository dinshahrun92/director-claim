const SPREADSHEET_ID = "1s20UbtnHp4WTLT1fVV7n0THfIo-YEs6TXqXOxBxMA38";
const MAIN_FOLDER_ID  = "1xOvGtcG9ypSppJPyZW_eWGEY7mRyIMlS";

// Placeholder text written by the old createInitialDraft — used only for
// backward-compatible filtering of legacy rows.
const LEGACY_PLACEHOLDER = "NEW DRAFT CREATED";

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function sanitize(v) {
  return v == null ? "" : v.toString().trim().replace(/[<>"'`]/g, "");
}

// ── setup ──────────────────────────────────────────────────────────────────

function doGet() {
  setupSheets();
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Claim & Payment Portal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const schemas = {
    App_Users:  ["UserID","Password","FullName"],
    App_Claims: ["RefNo","UserID","Recipient","InvoiceDate","Type","InvoiceNo",
                 "Description","Amount","FileUrl","Status","PaymentStatus",
                 "PaymentDate","Timestamp","CheckedBy","ApprovedBy","PaymentVia"]
  };

  Object.entries(schemas).forEach(([name, headers]) => {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);

    // Always verify and auto-correct headers (handles new sheet and any drift)
    const needsHeader = sheet.getLastRow() === 0;
    const headerMismatch = !needsHeader &&
      sheet.getRange(1, 1, 1, headers.length).getValues()[0]
           .some((v, i) => v.toString() !== headers[i]);

    if (needsHeader || headerMismatch) {
      sheet.getRange(1, 1, 1, headers.length)
           .setValues([headers])
           .setFontWeight("bold")
           .setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
  });
}

// ── auth ───────────────────────────────────────────────────────────────────

function login(id, pass) {
  if (!id || !pass) return { success: false, message: "ID and password required." };

  const data = getSheet("App_Users").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString() &&
        data[i][1].toString() === pass.toString()) {
      return { success: true, userId: id.toString(), name: data[i][2] };
    }
  }
  return { success: false, message: "Invalid credentials." };
}

function registerUser(id, pass, name) {
  id   = sanitize(id);
  pass = sanitize(pass);
  name = sanitize(name);

  if (!id || !pass || !name) return { success: false, message: "All fields are required." };
  if (id.length < 3)         return { success: false, message: "User ID must be at least 3 characters." };
  if (pass.length < 4)       return { success: false, message: "Password must be at least 4 characters." };

  const sheet = getSheet("App_Users");
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id) return { success: false, message: "User ID already exists." };
  }

  sheet.appendRow([id, pass, name]);
  return { success: true, message: "Account created. Please log in." };
}

// ── claims ─────────────────────────────────────────────────────────────────

function createInitialDraft(userId, recipient) {
  try {
    userId    = sanitize(userId);
    recipient = sanitize(recipient);
    if (!userId) return { success: false, message: "Invalid user." };

    const refNo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd")
                + "-" + Math.floor(1000 + Math.random() * 9000);

    // RefNo is generated here and returned to the client; no sheet row is written.
    // The draft only persists once the user clicks Save/Submit (saveBatchClaims).
    // An abandoned session simply leaves no trace in the sheet.
    return { success: true, refNo };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveBatchClaims(items, userId, recipient, isDraft, existingRef) {
  try {
    userId    = sanitize(userId);
    recipient = sanitize(recipient);
    existingRef = sanitize(existingRef);

    if (!userId || !existingRef) return { success: false, message: "Invalid parameters." };
    if (!Array.isArray(items) || items.length === 0)
      return { success: false, message: "No items to save." };

    const sheet  = getSheet("App_Claims");
    const data   = sheet.getDataRange().getValues();
    const status = isDraft ? "Draft" : "Submitted";
    const ts     = new Date();

    // Collect rows belonging to this ref (bottom-up to keep row indices valid)
    const toDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === existingRef && data[i][1].toString() === userId) {
        toDelete.push(i + 1); // 1-based sheet row
      }
    }
    for (let i = toDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(toDelete[i]);
    }

    // Append new rows
    items.forEach(item => {
      const amount = parseFloat(item.amount) || 0;

      // Store InvoiceDate as a formatted string to avoid timezone ambiguity
      let invoiceDate = "";
      if (item.date) {
        try {
          invoiceDate = Utilities.formatDate(new Date(item.date), Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch (_) {
          invoiceDate = item.date;
        }
      }

      sheet.appendRow([
        existingRef,
        userId,
        sanitize(recipient).toUpperCase(),
        invoiceDate,
        sanitize(item.type).toUpperCase(),
        sanitize(item.invNo).toUpperCase(),
        sanitize(item.description).toUpperCase(),
        amount,
        "NO ATTACHMENT",
        status,
        "Unpaid",
        "",
        ts,
        "", // CheckedBy
        "", // ApprovedBy
        sanitize(item.paymentVia)  // PaymentVia
      ]);
    });

    SpreadsheetApp.flush(); // ensure rows are committed before any subsequent read
    return { success: true, message: (isDraft ? "Draft saved: " : "Claim submitted: ") + existingRef };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getUserClaims(userId) {
  userId = sanitize(userId);
  if (!userId) return [];

  const tz   = Session.getScriptTimeZone();
  const data = getSheet("App_Claims").getDataRange().getValues();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString() !== userId) continue;

    const ref = data[i][0].toString();
    if (!grouped[ref]) {
      let dateStr = "";
      if (data[i][12]) {
        try {
          dateStr = Utilities.formatDate(new Date(data[i][12]), tz, "yyyy-MM-dd");
        } catch (_) {}
      }
      let paymentDateStr = "";
      if (data[i][11]) {
        try {
          paymentDateStr = Utilities.formatDate(new Date(data[i][11]), tz, "yyyy-MM-dd");
        } catch (_) {}
      }
      grouped[ref] = {
        refNo:        ref,
        recipient:    data[i][2] || "",
        total:        0,
        amountPaid:   0,
        _paidCount:   0,
        _rowCount:    0,
        status:       (data[i][9] || "").toString().trim(),
        payStatus:    "Unpaid",   // recomputed below
        date:         dateStr,
        paymentDate:  paymentDateStr
      };
    }
    const amount = parseFloat(data[i][7]) || 0;
    grouped[ref].total      += amount;
    grouped[ref]._rowCount  += 1;
    if ((data[i][10] || "").toString().trim() === "Paid") {
      grouped[ref]._paidCount += 1;
      grouped[ref].amountPaid += amount;
    }
  }

  return Object.values(grouped).map(g => {
    g.payStatus = g._paidCount === 0           ? "Unpaid"
                : g._paidCount === g._rowCount ? "Paid"
                :                                "Partial";
    delete g._paidCount;
    delete g._rowCount;
    return g;
  }).sort((a, b) => {
    const da = a.date ? new Date(a.date) : new Date(0);
    const db = b.date ? new Date(b.date) : new Date(0);
    return db - da;
  });
}

function getDraftDetails(refNo, userId) {
  refNo  = sanitize(refNo);
  userId = sanitize(userId);
  if (!refNo || !userId) return [];

  const data = getSheet("App_Claims").getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== refNo)   continue;
    if (data[i][1].toString() !== userId)  continue; // ownership check
    if (data[i][6].toString().toUpperCase() === LEGACY_PLACEHOLDER) continue; // skip old placeholders

    const rawDate = data[i][3];
    let dateStr = "";
    if (rawDate) {
      try {
        dateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
      } catch (_) {}
    }

    rows.push({
      date:        dateStr,
      type:        data[i][4],
      invNo:       data[i][5],
      description: data[i][6],
      amount:      data[i][7],
      paymentVia:  (data[i][15] || "").toString(),
      rowNum:      i + 1,              // 1-based sheet row, used for per-item payment
      payStatus:   (data[i][10] || "").toString().trim()
    });
  }
  return rows;
}

function processPayments(refNoList, userId) {
  userId = sanitize(userId);
  if (!userId) return { success: false, message: "Unauthorised." };
  if (!Array.isArray(refNoList) || refNoList.length === 0)
    return { success: false, message: "No references provided." };

  const sheet = getSheet("App_Claims");
  const data  = sheet.getDataRange().getValues();
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    const rowRef    = data[i][0].toString();
    const rowStatus = data[i][9].toString();

    if (!refNoList.includes(rowRef)) continue;
    if (rowStatus.trim() !== "Submitted")   continue; // only submitted claims

    sheet.getRange(i + 1, 11).setValue("Paid");
    sheet.getRange(i + 1, 12).setValue(today);
  }
  return { success: true };
}

// Returns all claims grouped by reference number, without filtering by owner.
// Used by the Payment Tracking tab so a director can see everyone's submitted claims.
function getAllClaims() {
  const tz   = Session.getScriptTimeZone();
  const data = getSheet("App_Claims").getDataRange().getValues();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    const ref = data[i][0].toString();
    if (!ref) continue;

    if (!grouped[ref]) {
      let dateStr = "";
      if (data[i][12]) {
        try { dateStr = Utilities.formatDate(new Date(data[i][12]), tz, "yyyy-MM-dd"); } catch (_) {}
      }
      let paymentDateStr = "";
      if (data[i][11]) {
        try { paymentDateStr = Utilities.formatDate(new Date(data[i][11]), tz, "yyyy-MM-dd"); } catch (_) {}
      }
      grouped[ref] = {
        refNo:       ref,
        recipient:   data[i][2] || "",
        submittedBy: data[i][1] || "",
        total:       0,
        amountPaid:  0,
        _paidCount:  0,
        _rowCount:   0,
        status:      (data[i][9] || "").toString().trim(),
        payStatus:   "Unpaid",
        date:        dateStr,
        paymentDate: paymentDateStr
      };
    }
    const amount = parseFloat(data[i][7]) || 0;
    grouped[ref].total     += amount;
    grouped[ref]._rowCount += 1;
    if ((data[i][10] || "").toString().trim() === "Paid") {
      grouped[ref]._paidCount += 1;
      grouped[ref].amountPaid += amount;
    }
  }

  return Object.values(grouped).map(g => {
    g.payStatus = g._paidCount === 0           ? "Unpaid"
                : g._paidCount === g._rowCount ? "Paid"
                :                                "Partial";
    delete g._paidCount;
    delete g._rowCount;
    return g;
  }).sort((a, b) => {
    const da = a.date ? new Date(a.date) : new Date(0);
    const db = b.date ? new Date(b.date) : new Date(0);
    return db - da;
  });
}

// Returns item rows for any claim regardless of ownership.
// Used by the payment panel so a director can view and pay any submitted claim's items.
function getClaimDetails(refNo) {
  refNo = sanitize(refNo);
  if (!refNo) return [];

  const data = getSheet("App_Claims").getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== refNo) continue;
    if (data[i][6].toString().toUpperCase() === LEGACY_PLACEHOLDER) continue;

    const rawDate = data[i][3];
    let dateStr = "";
    if (rawDate) {
      try { dateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch (_) {}
    }

    rows.push({
      date:        dateStr,
      type:        data[i][4],
      invNo:       data[i][5],
      description: data[i][6],
      amount:      data[i][7],
      paymentVia:  (data[i][15] || "").toString(),
      rowNum:      i + 1,
      payStatus:   (data[i][10] || "").toString().trim()
    });
  }
  return rows;
}

// Changes all Draft rows for a reference number to Submitted status.
function submitDraft(refNo, userId) {
  refNo  = sanitize(refNo);
  userId = sanitize(userId);
  if (!refNo || !userId) return { success: false, message: "Invalid parameters." };

  const sheet = getSheet("App_Claims");
  const data  = sheet.getDataRange().getValues();
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== refNo)   continue;
    if (data[i][1].toString() !== userId)  continue; // ownership check
    if (data[i][9].toString().trim() !== "Draft") continue;
    sheet.getRange(i + 1, 10).setValue("Submitted");
    updated++;
  }

  if (!updated) return { success: false, message: "No draft rows found." };
  SpreadsheetApp.flush();
  return { success: true, message: "Claim submitted: " + refNo };
}

// Marks individual sheet rows (by 1-based row number) as Paid.
// Used by the per-item selection in the claim detail side panel.
function processRowPayments(rowNums, userId) {
  userId = sanitize(userId);
  if (!userId) return { success: false, message: "Unauthorised." };
  if (!Array.isArray(rowNums) || !rowNums.length)
    return { success: false, message: "No rows provided." };

  const sheet = getSheet("App_Claims");
  const data  = sheet.getDataRange().getValues();
  const today = new Date();

  rowNums.forEach(rowNum => {
    const idx = Number(rowNum);
    if (!Number.isInteger(idx) || idx < 2 || idx > data.length) return; // skip header + OOB
    const i = idx - 1; // 0-based index into data array
    if (data[i][9].toString().trim() !== "Submitted")  return; // only submitted claims
    if ((data[i][10] || "").toString().trim() === "Paid") return; // skip already-paid rows
    sheet.getRange(idx, 11).setValue("Paid");
    sheet.getRange(idx, 12).setValue(today);
  });

  return { success: true };
}

// Returns all data needed to render a printable claim report.
// Includes line items, header info, and Prepared/Checked/Approved by details.
function getClaimReport(refNo) {
  refNo = sanitize(refNo);
  if (!refNo) return null;

  const tz        = Session.getScriptTimeZone();
  const claimData = getSheet("App_Claims").getDataRange().getValues();
  const userData  = getSheet("App_Users").getDataRange().getValues();

  // Build userId → fullName lookup
  const userMap = {};
  for (let i = 1; i < userData.length; i++) {
    userMap[userData[i][0].toString()] = userData[i][2].toString();
  }

  const items = [];
  let userId = "", recipient = "", submittedDate = "", checkedBy = "", approvedBy = "";
  let total = 0;

  for (let i = 1; i < claimData.length; i++) {
    if (claimData[i][0].toString() !== refNo) continue;
    if (claimData[i][6].toString().toUpperCase() === LEGACY_PLACEHOLDER) continue;

    // Capture header fields from the first matching row
    if (!userId) {
      userId     = claimData[i][1].toString();
      recipient  = claimData[i][2].toString();
      checkedBy  = (claimData[i][13] || "").toString().trim();
      approvedBy = (claimData[i][14] || "").toString().trim();
      try {
        submittedDate = Utilities.formatDate(new Date(claimData[i][12]), tz, "dd MMM yyyy");
      } catch (_) {}
    }

    const rawDate = claimData[i][3];
    let dateStr = "";
    if (rawDate) {
      try { dateStr = Utilities.formatDate(new Date(rawDate), tz, "dd MMM yyyy"); } catch (_) {}
    }

    const amount = parseFloat(claimData[i][7]) || 0;
    total += amount;

    items.push({
      date:        dateStr,
      type:        claimData[i][4].toString(),
      invNo:       claimData[i][5].toString(),
      description: claimData[i][6].toString(),
      amount:      amount,
      paymentVia:  (claimData[i][15] || "").toString(),
      rowNum:      i + 1,
      payStatus:   (claimData[i][10] || "").toString().trim()
    });
  }

  if (!userId) return null;

  return {
    refNo,
    recipient,
    submittedDate,
    total,
    items,
    preparedBy:   userMap[userId] || userId,
    preparedById: userId,
    checkedBy,
    approvedBy
  };
}
