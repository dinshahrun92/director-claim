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

function getUserClaims(userId) { // eslint-disable-line no-unused-vars
  // Returns all claims from all users (userId param retained for backward compatibility)
  const claimsData = getSheet("App_Claims").getDataRange().getValues();
  const usersData  = getSheet("App_Users").getDataRange().getValues();

  // Build userId → fullName map
  const userMap = {};
  for (let i = 1; i < usersData.length; i++) {
    userMap[usersData[i][0].toString()] = usersData[i][2] || usersData[i][0].toString();
  }

  const grouped = {};
  for (let i = 1; i < claimsData.length; i++) {
    const ref = claimsData[i][0].toString();
    if (!ref) continue;

    const isPlaceholder = claimsData[i][6] === "New Draft Created";
    const ownerId = claimsData[i][1].toString();

    // Always create the group entry so even fresh drafts appear in the listing
    if (!grouped[ref]) {
      grouped[ref] = {
        refNo:      ref,
        ownerId:    ownerId,
        ownerName:  userMap[ownerId] || ownerId,
        recipient:  claimsData[i][2] || "",
        total:      0,
        status:     claimsData[i][9]  || "",
        paidItems:  0,
        totalItems: 0,
        date:       claimsData[i][12]
      };
    }

    // Placeholder rows (initial draft marker) don't count toward totals
    if (isPlaceholder) continue;

    grouped[ref].total      += parseFloat(claimsData[i][7]) || 0;
    grouped[ref].totalItems += 1;
    if ((claimsData[i][10] || "").toString().toLowerCase() === "paid") {
      grouped[ref].paidItems += 1;
    }
  }

  // Compute payStatus per group
  const result = Object.values(grouped).map(g => {
    let payStatus;
    if (g.paidItems === 0)               payStatus = "Unpaid";
    else if (g.paidItems >= g.totalItems) payStatus = "Paid";
    else                                 payStatus = "Partial";
    const { paidItems, totalItems, ...rest } = g;
    return { ...rest, payStatus };
  });

  return result.sort((a, b) => new Date(b.date) - new Date(a.date));
}

function getDraftDetails(refNo, userId) {
  refNo = sanitize(refNo);
  if (!refNo) return [];

  const data = getSheet("App_Claims").getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== refNo)    continue;
    if (data[i][6] === "New Draft Created") continue; // skip placeholder

    const rawDate = data[i][3];
    let dateStr = "";
    if (rawDate) {
      try { dateStr = Utilities.formatDate(new Date(rawDate), "GMT+8", "yyyy-MM-dd"); } catch (_) {}
    }

    rows.push({
      rowNum:        i + 1,
      date:          dateStr,
      type:          data[i][4],
      invNo:         data[i][5],
      description:   data[i][6],
      amount:        data[i][7],
      claimStatus:   data[i][9]  || "",
      paymentStatus: data[i][10] || ""
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

// ── submit selected draft rows ─────────────────────────────────────────────

function submitSelectedRows(rowNums, userId) {
  userId = sanitize(userId);
  if (!userId) return { success: false, message: "Unauthorised." };
  if (!Array.isArray(rowNums) || !rowNums.length)
    return { success: false, message: "No rows selected." };

  const sheet = getSheet("App_Claims");
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const sheetRow = i + 1;
    if (!rowNums.includes(sheetRow))           continue;
    if (data[i][1].toString() !== userId)      continue; // ownership check
    if (data[i][9].toString() !== "Draft")     continue; // only draft rows

    sheet.getRange(sheetRow, 10).setValue("Submitted");
  }
  return { success: true };
}

// ── claim report data ──────────────────────────────────────────────────────

function getClaimReport(refNo) {
  refNo = sanitize(refNo);
  if (!refNo) return null;

  const claimsData = getSheet("App_Claims").getDataRange().getValues();
  const usersData  = getSheet("App_Users").getDataRange().getValues();

  let recipient = "", ownerId = "", submittedDate = "", total = 0;
  const items = [];

  for (let i = 1; i < claimsData.length; i++) {
    if (claimsData[i][0].toString() !== refNo)    continue;
    if (claimsData[i][6] === "New Draft Created")  continue;

    if (!ownerId) {
      ownerId   = claimsData[i][1] || "";
      recipient = claimsData[i][2] || "";
      const rawTs = claimsData[i][12];
      if (rawTs) {
        try { submittedDate = Utilities.formatDate(new Date(rawTs), "GMT+8", "dd MMM yyyy"); } catch (_) {}
      }
    }

    const rawDate = claimsData[i][3];
    let dateStr = "";
    if (rawDate) {
      try { dateStr = Utilities.formatDate(new Date(rawDate), "GMT+8", "dd MMM yyyy"); } catch (_) {}
    }

    const amount = parseFloat(claimsData[i][7]) || 0;
    total += amount;
    items.push({
      rowNum:      i + 1,
      date:        dateStr,
      type:        claimsData[i][4],
      invNo:       claimsData[i][5],
      description: claimsData[i][6],
      paymentVia:  "",
      amount:      amount
    });
  }

  if (!items.length) return null;

  let ownerName = ownerId;
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0].toString() === ownerId) { ownerName = usersData[i][2]; break; }
  }

  return {
    refNo,
    recipient,
    submittedDate,
    total,
    items,
    preparedBy:   ownerName,
    preparedById: ownerId,
    checkedBy:    "",
    approvedBy:   recipient
  };
}
