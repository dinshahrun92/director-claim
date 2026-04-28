const SPREADSHEET_ID = "1s20UbtnHp4WTLT1fVV7n0THfIo-YEs6TXqXOxBxMA38";
const MAIN_FOLDER_ID  = "1xOvGtcG9ypSppJPyZW_eWGEY7mRyIMlS";

// ── helpers ────────────────────────────────────────────────────────────────

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
                 "PaymentDate","Timestamp"]
  };

  Object.entries(schemas).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
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

    getSheet("App_Claims").appendRow([
      refNo, userId, recipient, "", "", "", "New Draft Created",
      0, "", "Draft", "Unpaid", "", new Date()
    ]);
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
      sheet.appendRow([
        existingRef,
        userId,
        sanitize(recipient),
        item.date    || "",
        sanitize(item.type),
        sanitize(item.invNo),
        sanitize(item.description),
        amount,
        "No Attachment",
        status,
        "Unpaid",
        "",
        ts
      ]);
    });

    return { success: true, message: (isDraft ? "Draft saved: " : "Claim submitted: ") + existingRef };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getUserClaims(userId) {
  userId = sanitize(userId);
  if (!userId) return [];

  const data    = getSheet("App_Claims").getDataRange().getValues();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString() !== userId) continue;

    const ref = data[i][0].toString();
    if (!grouped[ref]) {
      grouped[ref] = {
        refNo:     ref,
        recipient: data[i][2] || "",
        total:     0,
        status:    data[i][9]  || "",
        payStatus: data[i][10] || "",
        date:      data[i][12]
      };
    }
    grouped[ref].total += parseFloat(data[i][7]) || 0;
  }

  return Object.values(grouped).sort((a, b) => new Date(b.date) - new Date(a.date));
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
    if (data[i][6] === "New Draft Created") continue; // skip placeholder

    const rawDate = data[i][3];
    let dateStr = "";
    if (rawDate) {
      try { dateStr = Utilities.formatDate(new Date(rawDate), "GMT+8", "yyyy-MM-dd"); } catch (_) {}
    }

    rows.push({
      date:        dateStr,
      type:        data[i][4],
      invNo:       data[i][5],
      description: data[i][6],
      amount:      data[i][7]
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
    const rowOwner  = data[i][1].toString();
    const rowStatus = data[i][9].toString();

    if (!refNoList.includes(rowRef)) continue;
    if (rowOwner  !== userId)        continue; // only own claims
    if (rowStatus !== "Submitted")   continue; // only submitted claims

    sheet.getRange(i + 1, 11).setValue("Paid");
    sheet.getRange(i + 1, 12).setValue(today);
  }
  return { success: true };
}
