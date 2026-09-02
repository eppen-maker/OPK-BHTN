// Google Apps Script Web App for Drive-mappestrukturen under "1. Kundeinfo" i BHTN.
// Helt separat fra CRM-scriptet (Code.gs / "Kunder live") — rører ikke det arket i det hele
// tatt. Bruker DriveApp direkte, som gir en fullstendig, pålitelig mappeliste (i motsetning
// til søkebaserte Drive-integrasjoner som kan være begrenset til en indeks av nylig sette
// filer og derfor ikke vise hele mappeinnholdet).

const KUNDER_FOLDER_ID = "1hv5qdpGDtvTANa318RhNzh8u1F04vnUF"; // "1. Kundeinfo/2. Kunder"
const MALKUNDE_FOLDER_ID = "1IoozOFmHrtk0rmXXazQ0vmnturU8CFNB"; // ".../1. Malfiler/MALKUNDE - Kopieres ved nye kunder"
const SHARED_SECRET = "bhtn-drive-8f2c4a91d6e3b7051f9c2ea4d8b6317c";

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
//   { op: "createCustomer", companyName }                    — oppretter "{nr}. {navn}"-mappe og kopierer
//                                                                inn alle filene fra MALKUNDE-mappen
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
          return { op: "createCustomer", ...createCustomer(op.companyName || "") };
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

// Oppretter "{nr}. {firmanavn}"-mappen under "2. Kunder", og kopierer (ikke flytter) alle
// filene fra MALKUNDE-mappen inn i den nye kundemappen.
function createCustomer(companyName) {
  const name = (companyName || "").trim();
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

  return {
    duplicate: false,
    folderId: newFolder.getId(),
    folderName,
    folderUrl: newFolder.getUrl(),
    copiedFiles,
  };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
