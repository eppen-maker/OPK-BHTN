// Google Apps Script Web App for Drive-mappestrukturen under "1. Kundeinfo" i BHTN.
// Helt separat fra CRM-scriptet (Code.gs / "Kunder live") — rører ikke det arket i det hele
// tatt. Bruker DriveApp direkte, som gir en fullstendig, pålitelig mappeliste (i motsetning
// til søkebaserte Drive-integrasjoner som kan være begrenset til en indeks av nylig sette
// filer og derfor ikke vise hele mappeinnholdet).

const KUNDER_FOLDER_ID = "1hv5qdpGDtvTANa318RhNzh8u1F04vnUF"; // "1. Kundeinfo/2. Kunder"
const MALKUNDE_FOLDER_ID = "1IoozOFmHrtk0rmXXazQ0vmnturU8CFNB"; // ".../1. Malfiler/MALKUNDE - Kopieres ved nye kunder"
const SHARED_SECRET = "bhtn-drive-8f2c4a91d6e3b7051f9c2ea4d8b6317c";
// Kontrakt-malen ligger i MALKUNDE-mappen — brukes kun som en trygg, skrivebeskyttet
// autoriseringssjekk (åpnes read-only), aldri for å redigere selve malen.
const SAMPLE_CONTRACT_ID = "1RFq0e1X61u-6IxuMYX-t8wOf5-z-iu6YtOOmT3cmWsI";

// Kjøres KUN manuelt, én gang, direkte i editoren (velg denne funksjonen i dropdown-menyen
// øverst og trykk Kjør/Run). fillContractPlaceholders() bruker DocumentApp, som krever et eget
// samtykke-scope (documents) Google bare spør om når en funksjon kjøres direkte i editoren,
// ikke automatisk ved en ny deploy. Åpner malen kun for lesing — endrer ingenting.
function authorizeDocumentAccess() {
  DocumentApp.openById(SAMPLE_CONTRACT_ID).getBody().getText();
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== SHARED_SECRET) {
      return jsonOut({ error: "unauthorized" });
    }
    return jsonOut({ results: runOps(body.ops || []) });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// POST-svar fra Apps Script kommer tilbake pakket i en Google-omdirigering som curl ikke
// klarer å lese (kjente quirk — selve kjøringen fungerer, men svaret er utilgjengelig).
// GET-svar har ikke dette problemet, så tilby de skrivefrie op-ene her også, via query-param
// ?op=... i URL-en, slik at de kan verifiseres direkte.
function doGet(e) {
  try {
    const params = e.parameter || {};
    if (params.secret !== SHARED_SECRET) {
      return jsonOut({ error: "unauthorized" });
    }
    if (params.op === "nextCustomerNumber") {
      return jsonOut(computeNextCustomerNumber());
    }
    if (params.op === "listFolder") {
      const folder = DriveApp.getFolderById(params.folderId || KUNDER_FOLDER_ID);
      return jsonOut({ entries: listChildren(folder) });
    }
    return jsonOut({ error: "ukjent eller ikke-lesbar op-type for GET" });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// Støttede op-typer:
//   { op: "listFolder", folderId }                          — lister mapper+filer direkte under folderId
//   { op: "nextCustomerNumber" }                             — finner neste ledige kundenummer i "2. Kunder"
//   { op: "createCustomer", companyName, orgnr, contractDate, clientContact, clientEmail,
//     clientPhone, fee }                                     — oppretter "{nr}. {navn}"-mappe, kopierer
//                                                                inn alle filene fra MALKUNDE-mappen, og
//                                                                fyller ut kontrakt-KOPIEN (ikke malfilen)
//                                                                med kundedataene som er sendt inn
function runOps(ops) {
  return (ops || []).map((op) => {
    try {
      switch (op.op) {
        case "listFolder": {
          const folder = DriveApp.getFolderById(op.folderId);
          return { op: "listFolder", folderId: op.folderId, entries: listChildren(folder) };
        }
        case "nextCustomerNumber": {
          const { next, lastName } = computeNextCustomerNumber();
          return { op: "nextCustomerNumber", next, lastName };
        }
        case "createCustomer": {
          return { op: "createCustomer", ...createCustomer(op) };
        }
        default:
          return { op: op.op, error: "ukjent op-type" };
      }
    } catch (err) {
      return { op: op.op, error: String(err) };
    }
  });
}

// Returnerer direkte barn (mapper og filer) av en mappe, sortert på navn.
function listChildren(folder) {
  const entries = [];
  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    entries.push({ id: f.getId(), title: f.getName(), type: "folder" });
  }
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    entries.push({ id: f.getId(), title: f.getName(), type: "file", mimeType: f.getMimeType() });
  }
  entries.sort((a, b) => a.title.localeCompare(b.title, "no"));
  return entries;
}

// Kundemapper er navngitt "{nummer}. {firmanavn}" — finn høyeste eksisterende nummer blant
// mappene direkte under "2. Kunder", og returner neste ledige.
function computeNextCustomerNumber() {
  const folder = DriveApp.getFolderById(KUNDER_FOLDER_ID);
  const folders = folder.getFolders();
  let maxNum = 0;
  let lastName = null;
  while (folders.hasNext()) {
    const f = folders.next();
    const name = f.getName();
    const m = /^(\d+)\s*\./.exec(name);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxNum) {
        maxNum = n;
        lastName = name;
      }
    }
  }
  return { next: maxNum + 1, lastName };
}

// Oppretter "{nr}. {firmanavn}"-mappen under "2. Kunder", kopierer (ikke flytter) alle filene
// fra MALKUNDE-mappen inn i den nye kundemappen, og fyller ut kontrakt-KOPIEN (aldri malfilen
// i MALKUNDE selv) med kundedataene som er sendt inn.
function createCustomer(customer) {
  const name = (customer.companyName || "").trim();
  if (!name) throw new Error("companyName mangler");

  const kunderFolder = DriveApp.getFolderById(KUNDER_FOLDER_ID);

  // Nekt å opprette en ny kundemappe hvis firmanavnet allerede finnes blant eksisterende
  // mapper (uavhengig av nummer foran) — samme type duplikatsjekk som CRM-scriptet gjør mot
  // Bedrift-kolonnen i "Kunder live". Unngår at samme kunde ved en feil får to mapper (skjedde
  // med PARTSMEISTER AS: 69. ble opprettet manuelt samme dag som et testkall lagde 70.).
  const targetName = name.toLowerCase();
  const existingFolders = kunderFolder.getFolders();
  while (existingFolders.hasNext()) {
    const f = existingFolders.next();
    const existingName = f
      .getName()
      .replace(/^\d+\s*\.\s*/, "")
      .trim()
      .toLowerCase();
    if (existingName === targetName) {
      return { duplicate: true, folderId: f.getId(), folderName: f.getName(), folderUrl: f.getUrl() };
    }
  }

  const { next } = computeNextCustomerNumber();
  const folderName = `${next}. ${name}`;
  const newFolder = kunderFolder.createFolder(folderName);

  const malkundeFolder = DriveApp.getFolderById(MALKUNDE_FOLDER_ID);
  const copiedFiles = [];
  const files = malkundeFolder.getFiles();
  while (files.hasNext()) {
    const original = files.next();
    const copy = original.makeCopy(original.getName(), newFolder);
    copiedFiles.push({ id: copy.getId(), title: copy.getName() });
  }

  // Fyll ut kontrakt-KOPIEN som nettopp ble opprettet i kundemappen (aldri malfilen i
  // MALKUNDE) med kundedataene. Ikke-kritisk: feiler dette, er mappen og filene uansett
  // opprettet korrekt, så feilen fanges og legges ved i svaret i stedet for å kastes videre.
  const contractFile = copiedFiles.find((f) => f.title.indexOf("Kontrakt") !== -1);
  let contractFilled = false;
  let contractFillError = null;
  if (contractFile) {
    try {
      fillContractPlaceholders(contractFile.id, customer);
      contractFilled = true;
    } catch (err) {
      contractFillError = String(err);
    }
  }

  return {
    duplicate: false,
    folderId: newFolder.getId(),
    folderName,
    folderUrl: newFolder.getUrl(),
    copiedFiles,
    contractFilled,
    contractFillError,
  };
}

// Erstatter de gule plassholderfeltene ("SELSKAP AS", "xxx", og det feilaktig hardkodede
// "Vimms AS:" ved signaturfeltet) i kontrakt-kopien med kundens faktiske data.
function fillContractPlaceholders(fileId, customer) {
  const doc = DocumentApp.openById(fileId);
  const body = doc.getBody();

  const name = esc(customer.companyName || "");
  const orgnr = esc(formatOrgnr(customer.orgnr || ""));
  const address = esc(customer.clientAddress || "");
  const contact = esc(customer.clientContact || "");
  const email = esc(customer.clientEmail || "");
  const phone = esc(customer.clientPhone || "");
  const fee = esc(formatThousands(customer.fee || ""));
  const date = esc(customer.contractDate || "");

  // Apps Script sin replaceText() tolker IKKE $1/$2 i erstatningsteksten som
  // regex-backreferanser (skrives ut bokstavelig) — bruk derfor aldri capture-grupper i
  // erstatningen, sett heller inn ett fast mellomrom uansett hvor mange mellomrom originalen
  // hadde. Dette var også trolig årsaken til at Adresse-feltet ikke matchet i det hele tatt.
  body.replaceText("Oppdragsgiver:\\s*SELSKAP AS", "Oppdragsgiver: " + name);
  body.replaceText("Adresse:\\s*xxx", "Adresse: " + address);
  body.replaceText("Organisasjonsnummer:\\s*xxx", "Organisasjonsnummer: " + orgnr);
  body.replaceText("Kontaktperson:\\s*xxx", "Kontaktperson: " + contact);
  body.replaceText("Telefon:\\s*xxx", "Telefon: " + phone);
  body.replaceText("E-post:\\s*xxx", "E-post: " + email);
  body.replaceText("bistå xxx", "bistå " + name);
  body.replaceText("Digital Kartlegging i xxx", "Digital Kartlegging i " + name);
  body.replaceText("Kr\\s+xxx\\s+per år", "Kr " + fee + " per år");
  body.replaceText("Mandal, xxx", "Mandal, " + date);
  body.replaceText("Vimms AS:", name + ":");
  // Det som gjenstår av fristående "xxx" på dette punktet er signaturfeltet for kundens
  // daglig leder-navn — alle andre "xxx"-forekomster er allerede erstattet over.
  body.replaceText("\\bxxx\\b", contact);

  doc.saveAndClose();
}

// Formaterer organisasjonsnummer som "NNN NNN NNN".
function formatOrgnr(orgnr) {
  const digits = (orgnr || "").toString().replace(/\D/g, "");
  if (digits.length !== 9) return orgnr || "";
  return digits.replace(/(\d{3})(\d{3})(\d{3})/, "$1 $2 $3");
}

// Formaterer et tall med mellomrom som tusenskille, f.eks. 2990 -> "2 990".
function formatThousands(n) {
  const num = parseInt((n || "").toString().replace(/\D/g, ""), 10);
  if (isNaN(num)) return (n || "").toString();
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}

// Unngår at "$" i en kundeverdi (usannsynlig, men mulig i f.eks. et notat) blir tolket som
// en regex-backreferanse av replaceText().
function esc(s) {
  return (s || "").toString().replace(/\$/g, "$$$$");
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
