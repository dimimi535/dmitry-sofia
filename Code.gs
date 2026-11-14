/**
 * RSVP endpoint for Dmitry & Sofia wedding site.
 *
 * SETUP:
 * 1. Create a new Google spreadsheet.
 * 2. Copy its ID from the URL between /d/ and /edit.
 * 3. Replace SPREADSHEET_ID below.
 * 4. In the spreadsheet, optionally create a sheet named "RSVP".
 * 5. Deploy this script as a Web app (Execute as: Me; Who has access: Anyone).
 * 6. Copy the Web app URL ending in /exec into index.html as RSVP_ENDPOINT.
 *
 * Keep the spreadsheet itself private. Do not publish or share its contents with guests.
 */
const SPREADSHEET_ID = "1ENExKWVTvxDiHpQB7VCwDVlQkW_ly7PdHz4q_-7U6a8";
const SHEET_NAME = "RSVP";

function doPost(e) {
  try {
    if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf("ВСТАВЬТЕ") === 0) {
      throw new Error("Spreadsheet ID is not configured.");
    }

    const p = (e && e.parameter) ? e.parameter : {};
    const guestName = clean(p.guestName, 120);
    const attendance = clean(p.attendance, 40);
    const guestCount = clean(p.guestCount, 2);
    const additionalGuests = clean(p.additionalGuests, 500);
    const diet = clean(p.diet, 100);
    const alcohol = clean(p.alcohol, 100);
    const dietDetails = clean(p.dietDetails, 500);
    const mainDish = clean(p.mainDish, 20);
    const newForm = p.formVersion === "2";
    const comment = clean(newForm ? p.guestComment : p.comment, 1000);
    const event = clean(p.event, 120);
    const submittedAt = new Date();

    if (!guestName) throw new Error("Имя обязательно.");
    if (attendance !== "Да, буду" && attendance !== "К сожалению, не смогу") {
      throw new Error("Некорректный ответ о присутствии.");
    }
    if (attendance === "Да, буду" && (!/^[1-6]$/.test(guestCount) || !diet || !alcohol)) {
      throw new Error("Проверьте количество гостей, питание и предпочтения по напиткам.");
    }
    if (attendance === "Да, буду" && Number(guestCount) > 1 && !additionalGuests) {
      throw new Error("Укажите имена дополнительных гостей.");
    }

    if (newForm && attendance === "Да, буду" && !["Рыба", "Мясо"].includes(mainDish)) {
      throw new Error("Выберите основное блюдо.");
    }
    if (newForm && attendance === "Да, буду" && diet === "Есть ограничения или аллергии" && !dietDetails) {
      throw new Error("Укажите ограничения в питании.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      saveResponse([
        submittedAt,
        safeCell(guestName),
        attendance,
        attendance === "Да, буду" ? Number(guestCount) : 0,
        attendance === "Да, буду" ? safeCell(additionalGuests) : "",
        attendance === "Да, буду" ? safeCell(diet) : "",
        attendance === "Да, буду" ? safeCell(dietDetails) : "",
        attendance === "Да, буду" ? safeCell(alcohol) : "",
        safeCell(comment),
        safeCell(event)
      ], attendance === "Да, буду" ? safeCell(mainDish) : "");
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveResponse(row, mainDish) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const dishHeader = "Основное блюдо";
  const headers = [
    "Дата и время ответа", "Имя и фамилия", "Присутствие", "Количество гостей",
    "Имена дополнительных гостей", "Питание", "Подробности питания",
    "Предпочтения по напиткам", "Комментарий", "Мероприятие"
  ];
  let dishIndex = headers.length;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([...headers, dishHeader]);
    sheet.setFrozenRows(1);
  } else {
    const lastColumn = Math.max(headers.length, sheet.getLastColumn());
    const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    dishIndex = firstRow.indexOf(dishHeader, headers.length);
    if (dishIndex < 0) {
      dishIndex = lastColumn;
      if (dishIndex >= sheet.getMaxColumns()) sheet.insertColumnAfter(sheet.getMaxColumns());
      sheet.getRange(1, dishIndex + 1).setValue(dishHeader);
    }
  }
  while (row.length < dishIndex) row.push("");
  row[dishIndex] = mainDish;
  sheet.appendRow(row);
}

function clean(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

// Prevent spreadsheet formulas from being evaluated from user-supplied text.
function safeCell(value) {
  const s = String(value || "");
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet() {
  return ContentService
    .createTextOutput("RSVP endpoint is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}
