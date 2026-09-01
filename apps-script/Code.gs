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
    // "add" er den faste, Slack-knappen bruker denne. "ops" er en generisk liste med
    // admin-operasjoner (slette, formatere, lese/skrive celler) — se runOps() for hvilke
    // op-typer som støttes. Nye op-typer legges til i runOps() uten at selve doPost endres.
    if (body.action === "add" || !body.ops) {
      return jsonOut(addCustomerRow(body.data || {}));
    }
    return jsonOut({ results: runOps(body.ops || []) });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function getSheet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  if (!sheet) throw new Error('Fant ikke fanen "' + SHEET_TAB + '"');
  return sheet;
}

function norm(s) {
  return (s || "").toString().toLowerCase().replace(/[^a-zæøå0-9]/g, "");
}

// Godtar enten en kolonnebokstav ("I") eller en kolonneoverskrift ("Sum").
function resolveCol(sheet, ref) {
  if (/^[A-Za-z]+$/.test((ref || "").trim())) {
    return ref
      .trim()
      .toUpperCase()
      .split("")
      .reduce((acc, c) => acc * 26 + (c.charCodeAt(0) - 64), 0);
  }
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const idx = headerRow.findIndex((h) => norm(h) === norm(ref));
  if (idx === -1) throw new Error("Fant ikke kolonnen: " + ref);
  return idx + 1;
}

// Kjører en liste med admin-operasjoner mot arket, i rekkefølge. Dette er den generiske
// inngangen for alt utenom selve "legg til kunde"-flyten — nye ønsker (slette, formatere,
// lese/skrive celler) trenger ikke en ny kodeendring, bare en ny op her i listen.
//
// Støttede op-typer:
//   { op: "deleteRows", rows: [7, 9] }                         — sletter gitte radnumre
//   { op: "deleteByOrgnr", orgnrs: ["999999995"] }              — sletter rader med matchende Org.nummer
//   { op: "formatColumn", column: "I", fontColor, numberFormat, horizontalAlignment }
//   { op: "setCell", row: 12, column: "E", value: "ny@epost.no" }
//   { op: "readRow", row: 12 }                                   — returnerer alle celleverdier i raden
//   { op: "lookup", companyName: "Firma AS" }                    — finn rad basert på firmanavn
function runOps(ops) {
  const sheet = getSheet();
  return (ops || []).map((op) => {
    try {
      switch (op.op) {
        case "deleteRows": {
          const lastRow = sheet.getLastRow();
          const rows = Array.from(new Set((op.rows || []).map(Number)))
            .filter((n) => Number.isInteger(n) && n > HEADER_ROW && n <= lastRow)
            .sort((a, b) => b - a);
          rows.forEach((r) => sheet.deleteRow(r));
          return { op: "deleteRows", deleted: rows };
        }
        case "deleteByOrgnr": {
          const orgCol = resolveCol(sheet, "Org.nummer");
          const lastRow = sheet.getLastRow();
          if (lastRow <= HEADER_ROW) return { op: "deleteByOrgnr", deleted: [] };
          const targets = new Set((op.orgnrs || []).map((o) => (o || "").toString().replace(/\s/g, "")));
          const values = sheet.getRange(HEADER_ROW + 1, orgCol, lastRow - HEADER_ROW, 1).getValues();
          const deleted = [];
          for (let i = values.length - 1; i >= 0; i--) {
            const v = (values[i][0] || "").toString().replace(/\s/g, "");
            if (targets.has(v)) {
              const rowNum = i + HEADER_ROW + 1;
              sheet.deleteRow(rowNum);
              deleted.push(rowNum);
            }
          }
          return { op: "deleteByOrgnr", deleted };
        }
        case "formatColumn": {
          const col = resolveCol(sheet, op.column);
          const lastRow = Math.max(sheet.getLastRow(), HEADER_ROW + 1);
          const range = sheet.getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 1);
          if (op.fontColor) range.setFontColor(op.fontColor);
          if (op.numberFormat) range.setNumberFormat(op.numberFormat);
          if (op.horizontalAlignment) range.setHorizontalAlignment(op.horizontalAlignment);
          return { op: "formatColumn", column: col };
        }
        case "setCell": {
          const col = resolveCol(sheet, op.column);
          sheet.getRange(op.row, col).setValue(op.value);
          return { op: "setCell", row: op.row, column: col };
        }
        case "readRow": {
          const lastCol = sheet.getLastColumn();
          const values = sheet.getRange(op.row, 1, 1, lastCol).getValues()[0];
          return { op: "readRow", row: op.row, values };
        }
        case "lookup": {
          return { op: "lookup", ...lookupCompany(op.companyName || "") };
        }
        default:
          return { op: op.op, error: "ukjent op-type" };
      }
    } catch (err) {
      return { op: op.op, error: String(err) };
    }
  });
}

function lookupCompany(companyName) {
  const sheet = getSheet();
  const bedriftCol = resolveCol(sheet, "Bedrift");
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { found: false };
  const lastCol = sheet.getLastColumn();
  const rows = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
  const target = (companyName || "").trim().toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    if ((rows[i][bedriftCol - 1] || "").toString().trim().toLowerCase() === target) {
      return { found: true, row: i + HEADER_ROW + 1, values: rows[i] };
    }
  }
  return { found: false };
}

// Skjemaet leverer datoer som YYYY-MM-DD (HTML <input type="date">); arket bruker DD.MM.YYYY.
function formatDateNorwegian(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  if (!m) return iso || "";
  return `${m[3]}.${m[2]}.${m[1]}`;
}

function addCustomerRow(data) {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
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

  // Standardformatering på nye rader, satt direkte her fordi kolonneformatering ikke
  // arves automatisk av rader lagt til senere via API — hver ny rad må formateres selv.
  sheet.getRange(targetRow, 1, 1, row.length).setFontSize(8);

  const sumCol = findCol("Sum");
  if (sumCol !== -1) {
    const sumRange = sheet.getRange(targetRow, sumCol + 1);
    sumRange.setFontColor("#ea9999");
    sumRange.setNumberFormat('#,##0.00" kr"');
  }
  const oppstartCol = findCol("Oppstartsdato");
  if (oppstartCol !== -1) sheet.getRange(targetRow, oppstartCol + 1).setHorizontalAlignment("center");
  const mobilCol = findCol("Mobilnummer");
  if (mobilCol !== -1) sheet.getRange(targetRow, mobilCol + 1).setHorizontalAlignment("center");

  return { duplicate: false, row: targetRow };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
