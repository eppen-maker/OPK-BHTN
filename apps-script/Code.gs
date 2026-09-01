// Google Apps Script Web App bundet til "KUNDELISTE BHTN"-arket.
// Erstatter Sheets API + tjenestekonto-nøkkel for å skrive nye kunder inn i "Kunder live".
// Oppsett: se apps-script/README.md.

const SHEET_ID = "1DyP84qxN27nBo6v3F7f36i8Ljs0EJZWcTCPYlVdlC4U";
const SHEET_TAB = "Kunder live";
const SHARED_SECRET = "4eeb6ea868f3fb04746ebd79e9ea619d32f7db425d56a5c6";
// Rad 1-2 er tittel/lenke-celler i "Kunder live" — de faktiske kolonneoverskriftene ligger på rad 3.
const HEADER_ROW = 3;

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== SHARED_SECRET) {
      return jsonOut({ error: "unauthorized" });
    }
    const action = body.action || "add";
    if (action === "add") return jsonOut(addCustomerRow(body.data || {}));
    if (action === "lookup") return jsonOut(lookupCompany(body.companyName || ""));
    if (action === "deleteByOrgnr") return jsonOut(deleteRowsByOrgnr(body.orgnrs || []));
    if (action === "formatColumn") return jsonOut(formatColumn(body.column, body.options || {}));
    return jsonOut({ error: "unknown action: " + action });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// Admin: finn nøyaktig hva som står i arket for et gitt firmanavn (for feilsøking av
// "gikk ikke over"-tilfeller uten å måtte gjette).
function lookupCompany(companyName) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  if (!sheet) throw new Error('Fant ikke fanen "' + SHEET_TAB + '"');
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const norm = (s) => (s || "").toString().toLowerCase().replace(/[^a-zæøå0-9]/g, "");
  const bedriftCol = headerRow.findIndex((h) => norm(h) === norm("Bedrift"));
  if (bedriftCol === -1) throw new Error('Fant ikke kolonnen "Bedrift"');
  if (lastRow <= HEADER_ROW) return { found: false };
  const rows = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
  const target = (companyName || "").trim().toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    if ((rows[i][bedriftCol] || "").toString().trim().toLowerCase() === target) {
      return { found: true, row: i + HEADER_ROW + 1, values: rows[i] };
    }
  }
  return { found: false };
}

// Admin: slett rader der Org.nummer-kolonnen matcher en av de gitte verdiene (brukt til å
// rydde bort testrader). Sletter nedenfra og opp så radnumrene ikke forskyves underveis.
function deleteRowsByOrgnr(orgnrs) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  if (!sheet) throw new Error('Fant ikke fanen "' + SHEET_TAB + '"');
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const norm = (s) => (s || "").toString().toLowerCase().replace(/[^a-zæøå0-9]/g, "");
  const orgCol = headerRow.findIndex((h) => norm(h) === norm("Org.nummer"));
  if (orgCol === -1) throw new Error('Fant ikke kolonnen "Org.nummer"');
  if (lastRow <= HEADER_ROW) return { deleted: [] };

  const targets = new Set((orgnrs || []).map((o) => (o || "").toString().replace(/\s/g, "")));
  const rows = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
  const deleted = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const orgVal = (rows[i][orgCol] || "").toString().replace(/\s/g, "");
    if (targets.has(orgVal)) {
      const rowNum = i + HEADER_ROW + 1;
      sheet.deleteRow(rowNum);
      deleted.push(rowNum);
    }
  }
  return { deleted };
}

// Admin: still en kolonne (bokstav som "I", eller kolonneoverskrift som "Sum") — valgfri
// tekstfarge, tallformat og horisontal justering. Gjelder fra header-raden og nedover.
function formatColumn(columnRef, options) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  if (!sheet) throw new Error('Fant ikke fanen "' + SHEET_TAB + '"');
  const lastCol = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), HEADER_ROW + 1);

  let colIndex;
  if (/^[A-Za-z]+$/.test((columnRef || "").trim())) {
    colIndex = columnRef.trim().toUpperCase().split("").reduce((acc, c) => acc * 26 + (c.charCodeAt(0) - 64), 0);
  } else {
    const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const norm = (s) => (s || "").toString().toLowerCase().replace(/[^a-zæøå0-9]/g, "");
    colIndex = headerRow.findIndex((h) => norm(h) === norm(columnRef)) + 1;
  }
  if (!colIndex || colIndex < 1) throw new Error("Fant ikke kolonnen: " + columnRef);

  const range = sheet.getRange(HEADER_ROW + 1, colIndex, lastRow - HEADER_ROW, 1);
  if (options.fontColor) range.setFontColor(options.fontColor);
  if (options.numberFormat) range.setNumberFormat(options.numberFormat);
  if (options.horizontalAlignment) range.setHorizontalAlignment(options.horizontalAlignment);
  return { formatted: true, column: colIndex };
}

// Skjemaet leverer datoer som YYYY-MM-DD (HTML <input type="date">); arket bruker DD.MM.YYYY.
function formatDateNorwegian(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  if (!m) return iso || "";
  return `${m[3]}.${m[2]}.${m[1]}`;
}

function addCustomerRow(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  if (!sheet) throw new Error('Fant ikke fanen "' + SHEET_TAB + '"');

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const norm = (s) => (s || "").toString().toLowerCase().replace(/[^a-zæøå0-9]/g, "");
  const findCol = (name) => headerRow.findIndex((h) => norm(h) === norm(name));
  const protect = (v) => (typeof v === "string" && /^[+=@-]/.test(v) ? "'" + v : v);

  const bedriftCol = findCol("Bedrift");
  if (bedriftCol === -1) throw new Error('Fant ikke kolonnen "Bedrift"');

  const bedriftValues =
    lastRow > HEADER_ROW
      ? sheet
          .getRange(HEADER_ROW + 1, bedriftCol + 1, lastRow - HEADER_ROW, 1)
          .getValues()
          .map((r) => (r[0] || "").toString().trim().toLowerCase())
      : [];

  const targetName = (data.companyName || "").trim().toLowerCase();
  const existingIdx = bedriftValues.indexOf(targetName);
  if (existingIdx !== -1) {
    return { duplicate: true, row: existingIdx + HEADER_ROW + 1 };
  }

  // Mange ark har formler forhåndsutfylt langt nedover rutenettet (malrader), så getLastRow()
  // kan gi et altfor høyt tall. Finn derfor siste reelle rad basert på Bedrift-kolonnen.
  let lastDataRow = HEADER_ROW;
  bedriftValues.forEach((v, i) => {
    if (v) lastDataRow = i + HEADER_ROW + 1;
  });
  let targetRow = lastDataRow + 1;

  // Sikkerhetssjekk: hopp forbi enhver rad som allerede har innhold i EN AV kolonnene
  // (f.eks. en oppsummeringsrad lenger ned), slik at vi aldri overskriver noe eksisterende.
  while (targetRow <= lastRow) {
    const existingRowValues = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
    if (existingRowValues.some((v) => (v || "").toString().trim())) {
      targetRow++;
    } else {
      break;
    }
  }

  const byHeader = {};
  byHeader[norm("Org.nummer")] = data.orgnr || "";
  byHeader[norm("Oppstartsdato")] = formatDateNorwegian(data.contractDate);
  byHeader[norm("Bedrift")] = data.companyName || "";
  byHeader[norm("Kontaktperson")] = protect(data.clientContact || "");
  byHeader[norm("E-post")] = data.clientEmail || "";
  byHeader[norm("Mobilnummer")] = protect(data.clientPhone || "");
  byHeader[norm("Fakturering")] = formatDateNorwegian(data.invoiceDate);
  byHeader[norm("Fakt. frekvens")] = data.invoiceFrequency || "";
  byHeader[norm("Sum")] = data.fee || "";
  byHeader[norm("Antall ansatte")] = data.customerCount || "";
  byHeader[norm("Annen info")] = protect(data.otherInfo || "");

  const row = headerRow.map((h) => {
    const key = norm(h);
    return key in byHeader ? byHeader[key] : "";
  });

  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { duplicate: false, row: targetRow };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
