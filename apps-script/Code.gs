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
    return jsonOut(addCustomerRow(body.data || {}));
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
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
