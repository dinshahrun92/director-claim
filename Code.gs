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
    App_Users:       ["UserID","Password","FullName"],
    App_Claims:      ["RefNo","UserID","Recipient","InvoiceDate","Type","InvoiceNo",
                      "Description","Amount","FileUrl","Status","PaymentStatus",
                      "PaymentDate","Timestamp","PaymentVia"],
    App_Settings:    ["Category","Value"],
    App_Attachments: ["RefNo","UserID","FileName","FileUrl","Timestamp","FileId"]
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

    const tz   = Session.getScriptTimeZone();
    const now  = new Date();
    const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN',
                    'JUL','AUG','SEP','OCT','NOV','DEC'];
    const prefix = MONTHS[now.getMonth()] + Utilities.formatDate(now, tz, "yy"); // e.g. JAN26

    // Determine next sequential number for this month-year prefix
    const sheet     = getSheet("App_Claims");
    const allData   = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [];
    let maxSeq = 0;
    for (let i = 1; i < allData.length; i++) {
      const ref = allData[i][0].toString();
      if (ref.startsWith(prefix + "-")) {
        const seq = parseInt(ref.slice(prefix.length + 1), 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    }
    const refNo = prefix + "-" + (maxSeq + 1).toString().padStart(3, "0");

    sheet.appendRow([
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
        ts,
        sanitize(item.paymentVia || "")
      ]);
    });

    return { success: true, message: (isDraft ? "Draft saved: " : "Claim submitted: ") + existingRef };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getUserClaims(userId) { // eslint-disable-line no-unused-vars
  // Returns all claims from all users (userId param retained for backward compatibility)
  const claimsSheet = getSheet("App_Claims");
  const usersSheet  = getSheet("App_Users");
  if (!claimsSheet || !usersSheet) return [];

  const claimsData = claimsSheet.getLastRow() > 1 ? claimsSheet.getDataRange().getValues() : [];
  const usersData  = usersSheet.getLastRow()  > 1 ? usersSheet.getDataRange().getValues()  : [];

  const tz = Session.getScriptTimeZone();

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
    const ownerId  = claimsData[i][1].toString();
    const rowStatus = (claimsData[i][9] || "").toString();

    if (!grouped[ref]) {
      // Convert the Timestamp (Date object) to a plain ISO string to avoid GAS
      // serialisation failures when returning to the client via google.script.run.
      let dateStr = "";
      const rawTs = claimsData[i][12];
      if (rawTs) {
        try { dateStr = Utilities.formatDate(new Date(rawTs), tz, "yyyy-MM-dd'T'HH:mm:ss"); } catch (_) {}
      }

      grouped[ref] = {
        refNo:      ref,
        ownerId:    ownerId,
        ownerName:  userMap[ownerId] || ownerId,
        recipient:  (claimsData[i][2] || "").toString(),
        total:      0,
        status:     rowStatus,
        paidItems:  0,
        totalItems: 0,
        date:       dateStr
      };
    } else {
      // If any row of this ref is still Draft the whole claim is shown as Draft;
      // otherwise keep the first-seen status (Submitted / Approved etc.).
      if (rowStatus === "Draft") {
        grouped[ref].status = "Draft";
      }
    }

    // Placeholder rows (initial draft marker) don't count toward totals
    if (isPlaceholder) continue;

    grouped[ref].total      += parseFloat(claimsData[i][7]) || 0;
    grouped[ref].totalItems += 1;
    if ((claimsData[i][10] || "").toString().toLowerCase() === "paid") {
      grouped[ref].paidItems += 1;
    }
    // Accumulate description text for client-side full-text search
    const desc = (claimsData[i][6] || "").toString().trim();
    if (desc) grouped[ref].descText = (grouped[ref].descText || "") + " " + desc.toLowerCase();
  }

  // Compute payStatus per group and return plain serialisable objects
  const result = Object.values(grouped).map(g => {
    let payStatus;
    if (g.paidItems === 0)               payStatus = "Unpaid";
    else if (g.paidItems >= g.totalItems) payStatus = "Paid";
    else                                 payStatus = "Partial";
    return {
      refNo:     g.refNo,
      ownerId:   g.ownerId,
      ownerName: g.ownerName,
      recipient: g.recipient,
      total:     g.total,
      status:    g.status,
      date:      g.date,
      payStatus: payStatus,
      descText:  (g.descText || "").trim()
    };
  });

  return result.sort((a, b) => {
    if (!a.date && !b.date) return 0;
    if (!a.date) return 1;
    if (!b.date) return -1;
    return b.date < a.date ? -1 : b.date > a.date ? 1 : 0;
  });
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
      paymentStatus: data[i][10] || "",
      paymentVia:    data[i][13] || ""
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
      rowNum:        i + 1,
      date:          dateStr,
      type:          claimsData[i][4],
      invNo:         claimsData[i][5],
      description:   claimsData[i][6],
      paymentVia:    claimsData[i][13] || "",
      paymentStatus: claimsData[i][10] || "",
      amount:        amount
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

const MAX_SEARCH_RESULTS = 80; // cap on server-side search results

// ── delete draft ──────────────────────────────────────────────────────────

function deleteDraft(refNo, userId) {
  try {
    refNo  = sanitize(refNo);
    userId = sanitize(userId);
    if (!refNo || !userId) return { success: false, message: "Invalid parameters." };

    const sheet = getSheet("App_Claims");
    const data  = sheet.getDataRange().getValues();
    const toDelete = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() !== refNo) continue;
      if (data[i][1].toString() !== userId)
        return { success: false, message: "Unauthorized: you don't own this draft." };
      const rowStatus = (data[i][9] || "").toString();
      if (rowStatus === "Submitted" || rowStatus === "Approved")
        return { success: false, message: "Cannot delete: claim has already been submitted." };
      toDelete.push(i + 1);
    }

    if (!toDelete.length) return { success: false, message: "Draft not found." };

    // Delete rows bottom-up to preserve row indices
    for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);

    // Clean up attachments
    const attSheet = getSheet("App_Attachments");
    if (attSheet && attSheet.getLastRow() > 1) {
      const attData = attSheet.getDataRange().getValues();
      for (let i = attData.length - 1; i >= 1; i--) {
        if (attData[i][0].toString() === refNo) {
          attSheet.deleteRow(i + 1);
          try { DriveApp.getFileById(attData[i][5].toString()).setTrashed(true); } catch (_) {}
        }
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ── mark rows as paid ─────────────────────────────────────────────────────

function markRowsAsPaid(rowNums, userId) {
  userId = sanitize(userId);
  if (!userId) return { success: false, message: "Unauthorised." };
  if (!Array.isArray(rowNums) || !rowNums.length)
    return { success: false, message: "No rows selected." };

  const sheet = getSheet("App_Claims");
  const data  = sheet.getDataRange().getValues();
  const today = new Date();
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const sheetRow = i + 1;
    if (!rowNums.includes(sheetRow)) continue;
    if (data[i][1].toString() !== userId) continue; // ownership
    if (data[i][9].toString() !== "Submitted") continue; // only submitted rows
    sheet.getRange(sheetRow, 11).setValue("Paid");
    sheet.getRange(sheetRow, 12).setValue(today);
    updated++;
  }

  return updated > 0
    ? { success: true, updated }
    : { success: false, message: "No eligible submitted rows found. Only submitted items owned by you can be marked paid." };
}

// ── item-level search ──────────────────────────────────────────────────────

function searchClaimItems(query) {
  query = sanitize(query).toLowerCase().trim();
  if (!query || query.length < 2) return [];

  const claimsSheet = getSheet("App_Claims");
  const usersSheet  = getSheet("App_Users");
  if (!claimsSheet) return [];

  const claimsData = claimsSheet.getLastRow() > 1 ? claimsSheet.getDataRange().getValues() : [];
  const usersData  = usersSheet && usersSheet.getLastRow() > 1 ? usersSheet.getDataRange().getValues() : [];
  const tz = Session.getScriptTimeZone();

  const userMap = {};
  for (let i = 1; i < usersData.length; i++) {
    userMap[usersData[i][0].toString()] = usersData[i][2] || usersData[i][0].toString();
  }

  const results = [];
  for (let i = 1; i < claimsData.length; i++) {
    const row = claimsData[i];
    if ((row[6] || "").toString() === "New Draft Created") continue;

    const description = (row[6] || "").toString();
    const type        = (row[4] || "").toString();
    const invNo       = (row[5] || "").toString();
    const refNo       = (row[0] || "").toString();
    const recipient   = (row[2] || "").toString();
    const ownerId     = (row[1] || "").toString();

    if (description.toLowerCase().includes(query) ||
        type.toLowerCase().includes(query)         ||
        invNo.toLowerCase().includes(query)        ||
        refNo.toLowerCase().includes(query)        ||
        recipient.toLowerCase().includes(query)    ||
        (userMap[ownerId] || "").toLowerCase().includes(query)) {

      const rawDate = row[3];
      let dateStr = "";
      if (rawDate) {
        try { dateStr = Utilities.formatDate(new Date(rawDate), tz, "dd MMM yyyy"); } catch (_) {}
      }

      results.push({
        refNo:       refNo,
        ownerId:     ownerId,
        ownerName:   userMap[ownerId] || ownerId,
        recipient:   recipient,
        date:        dateStr,
        type:        type,
        invNo:       invNo,
        description: description,
        amount:      parseFloat(row[7]) || 0,
        claimStatus: (row[9]  || "").toString(),
        payStatus:   (row[10] || "").toString()
      });
    }
  }

  // Sort: drafts first, then by date desc; limit to 80 results
  results.sort((a, b) => {
    if (a.claimStatus === "Draft" && b.claimStatus !== "Draft") return -1;
    if (a.claimStatus !== "Draft" && b.claimStatus === "Draft") return 1;
    if (a.date < b.date) return 1;
    if (a.date > b.date) return -1;
    return 0;
  });

  return results.slice(0, MAX_SEARCH_RESULTS);
}

// ── settings ───────────────────────────────────────────────────────────────

const SETTING_DEFAULTS = {
  claimTypes: ["Petrol","TnG","Refreshment","Medical","Others"],
  payVias:    ["Cash","MBB 1","MBB 2","CIMB","SC","UOB"]
};

function getSettings() {
  const sheet = getSheet("App_Settings");
  if (!sheet || sheet.getLastRow() < 2) return SETTING_DEFAULTS;

  const data = sheet.getDataRange().getValues();
  const claimTypes = [], payVias = [];
  for (let i = 1; i < data.length; i++) {
    const cat = data[i][0].toString().trim();
    const val = data[i][1].toString().trim();
    if (!val) continue;
    if (cat === "ClaimType") claimTypes.push(val);
    else if (cat === "PayVia")    payVias.push(val);
  }
  return {
    claimTypes: claimTypes.length ? claimTypes : SETTING_DEFAULTS.claimTypes,
    payVias:    payVias.length    ? payVias    : SETTING_DEFAULTS.payVias
  };
}

function saveSettings(claimTypes, payVias) {
  try {
    const sheet = getSheet("App_Settings");
    if (!sheet) return { success: false, message: "Settings sheet not found." };

    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);

    const rows = [];
    (claimTypes || []).forEach(t => rows.push(["ClaimType", sanitize(t)]));
    (payVias    || []).forEach(v => rows.push(["PayVia",    sanitize(v)]));

    if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ── attachments ────────────────────────────────────────────────────────────

function uploadAttachment(refNo, userId, fileName, base64Data, mimeType) {
  try {
    refNo    = sanitize(refNo);
    userId   = sanitize(userId);
    fileName = sanitize(fileName);
    if (!refNo || !userId || !base64Data) return { success: false, message: "Invalid parameters." };

    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const iter = mainFolder.getFoldersByName(refNo);
    const refFolder = iter.hasNext() ? iter.next() : mainFolder.createFolder(refNo);

    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType || "application/octet-stream", fileName);
    const file    = refFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = file.getUrl();
    const fileId  = file.getId();

    getSheet("App_Attachments").appendRow([refNo, userId, fileName, fileUrl, new Date(), fileId]);
    return { success: true, fileName, fileUrl };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getAttachments(refNo) {
  refNo = sanitize(refNo);
  if (!refNo) return [];
  const sheet = getSheet("App_Attachments");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === refNo) {
      results.push({ fileName: data[i][2], fileUrl: data[i][3] });
    }
  }
  return results;
}

function deleteAttachment(refNo, userId, fileUrl) {
  try {
    refNo  = sanitize(refNo);
    userId = sanitize(userId);
    if (!refNo || !userId || !fileUrl) return { success: false, message: "Invalid parameters." };

    const sheet = getSheet("App_Attachments");
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: "Attachment not found." };

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString() === refNo &&
          data[i][1].toString() === userId &&
          data[i][3].toString() === fileUrl) {
        sheet.deleteRow(i + 1);
        try { DriveApp.getFileById(data[i][5].toString()).setTrashed(true); } catch (_) {}
        return { success: true };
      }
    }
    return { success: false, message: "Attachment not found." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
