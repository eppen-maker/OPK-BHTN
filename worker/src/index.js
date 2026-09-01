import { connect } from "cloudflare:sockets";

// Base64-kodet PNG av din logo (f.eks. samme fil som brukes på selve nettsiden:
// frontend/assets/logo-color.png), brukt som innebygd bilde (Content-ID) i e-post-signaturen.
const LOGO_BASE64 = ""; // TODO: base64 av din egen logo (PNG), eller la stå tom for å droppe logo i signaturen

const SHEET_TAB = "Kunder live"; // navnet på hovedfanen i "KUNDELISTE BHTN"-arket

function base64url(input) {
  let base64;
  if (typeof input === "string") {
    base64 = btoa(unescape(encodeURIComponent(input)));
  } else {
    base64 = btoa(String.fromCharCode(...new Uint8Array(input)));
  }
  return base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function importGooglePrivateKey(pem) {
  const pemContents = pem
    .replace(/\\n/g, "\n")
    .replace("-----BEGIN PRIVATE KEY-----", "")
    .replace("-----END PRIVATE KEY-----", "")
    .replace(/\s/g, "");
  const binaryDer = Uint8Array.from(atob(pemContents), (c) => c.charCodeAt(0));
  return crypto.subtle.importKey(
    "pkcs8",
    binaryDer.buffer,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"]
  );
}

async function getGoogleSheetsToken(env) {
  const header = { alg: "RS256", typ: "JWT" };
  const now = Math.floor(Date.now() / 1000);
  const claims = {
    iss: env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  };
  const unsigned = `${base64url(JSON.stringify(header))}.${base64url(JSON.stringify(claims))}`;
  const key = await importGooglePrivateKey(env.GOOGLE_SERVICE_ACCOUNT_KEY);
  const signature = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(unsigned));
  const jwt = `${unsigned}.${base64url(signature)}`;

  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt,
    }),
  });
  if (!res.ok) {
    throw new Error("Google-token feilet: " + (await res.text()).slice(0, 300));
  }
  const json = await res.json();
  return json.access_token;
}

const NO_MONTHS = ["jan", "feb", "mar", "apr", "mai", "jun", "jul", "aug", "sep", "okt", "nov", "des"];

function getISOWeek(date) {
  const d = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function formatWeekLabel(today) {
  const dayNum = today.getUTCDay() || 7; // mandag=1 .. søndag=7
  const monday = new Date(today.getTime() - (dayNum - 1) * 86400000);
  const sunday = new Date(monday.getTime() + 6 * 86400000);
  const week = getISOWeek(monday);
  const sameMonth = monday.getUTCMonth() === sunday.getUTCMonth();
  const startPart = sameMonth ? `${monday.getUTCDate()}.` : `${monday.getUTCDate()}. ${NO_MONTHS[monday.getUTCMonth()]}`;
  const endPart = `${sunday.getUTCDate()}. ${NO_MONTHS[sunday.getUTCMonth()]} ${sunday.getUTCFullYear()}`;
  return { week, range: `${startPart}–${endPart}` };
}

async function runWeeklyBackup(env) {
  const token = await getGoogleSheetsToken(env);
  const sheetId = env.GOOGLE_SHEET_ID;

  // Kolonne A = Dato, B = Bedrift i "KUNDELISTE BHTN"-arket. Arket har ingen org.nummer-kolonne,
  // så bedriftsnavnet (B) brukes som nøkkel for å oppdage nye kunder uke til uke — ikke
  // vanntett hvis to kunder får identisk navn, men det er det nærmeste vi har til en unik id her.
  const rows = await getSheetRange(token, sheetId, `'${SHEET_TAB}'!A2:B`);
  const current = rows.filter((r) => (r[1] || "").trim()).map((r) => ({ key: (r[1] || "").trim(), date: (r[0] || "").trim() }));

  const prevRaw = await env.BACKUP_STATE.get("last_customers");
  const prev = prevRaw ? JSON.parse(prevRaw) : [];
  const prevKeys = new Set(prev.map((c) => c.key));

  const added = current.filter((c) => !prevKeys.has(c.key));

  const { week, range } = formatWeekLabel(new Date());

  let changesLine;
  if (prevRaw === null) {
    changesLine = "Første automatiske sikkerhetskopi — ingen sammenligning tilgjengelig ennå.";
  } else if (added.length === 0) {
    changesLine = "Ingen nye kunder denne uken.";
  } else {
    const names = added.map((c) => c.key).join(", ");
    changesLine = `${added.length} ny${added.length === 1 ? "" : "e"} kunde${added.length === 1 ? "" : "r"} denne uken: ${names}.`;
  }

  const description = `Automatisk ukentlig sikkerhetskopi. Uke ${week} (${range}). ${changesLine}`;
  const fileName = `Backup uke ${week} (${range}).xlsx`; // TODO: sett gjerne et fast prefiks, f.eks. "Kundeark - Backup..."

  // Tjenestekontoen har ingen egen lagringsplass i Drive (vanlig for kontoer utenfor Google
  // Workspace), så en vanlig Drive-kopi feiler med "storage quota exceeded". Vi eksporterer i
  // stedet arket som .xlsx (ingen lagring involvert — bare en engangs-nedlasting av bytes) og
  // sender det som vedlegg på e-post via SMTP.
  const exportRes = await fetch(
    `https://www.googleapis.com/drive/v3/files/${sheetId}/export?mimeType=${encodeURIComponent(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!exportRes.ok) throw new Error("Kunne ikke eksportere ark: " + (await exportRes.text()).slice(0, 300));
  const buffer = await exportRes.arrayBuffer();
  const contentBase64 = arrayBufferToBase64(buffer);

  await sendSmtpMail(env, {
    to: env.BACKUP_RECIPIENT, // TODO: sett BACKUP_RECIPIENT som secret (mottaker av ukentlig backup)
    subject: `Ukentlig sikkerhetskopi av CRM — uke ${week}`,
    text: description,
    attachments: [
      {
        filename: fileName,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        contentBase64,
      },
    ],
  });

  await env.BACKUP_STATE.put("last_customers", JSON.stringify(current));

  return { fileName, description, addedCount: added.length, sizeBytes: buffer.byteLength };
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function getSheetRange(token, sheetId, range) {
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error("Kunne ikke lese ark: " + (await res.text()).slice(0, 300));
  const json = await res.json();
  return json.values || [];
}

// Skriver til "Kunder live"-fanen via et Google Apps Script Web App (bundet til selve arket)
// i stedet for Sheets API med en tjenestekonto-nøkkel. Samme radoppsett/dupliseringssjekk som
// før, men uten JWT/privatnøkkel-håndtering på worker-siden — se apps-script/Code.gs.
async function syncViaAppsScript(data, env) {
  const res = await fetch(env.GOOGLE_APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ secret: env.GOOGLE_APPS_SCRIPT_SECRET, data }),
  });
  if (!res.ok) throw new Error("Apps Script-kall feilet: " + (await res.text()).slice(0, 300));
  const json = await res.json();
  if (json.error) throw new Error(json.error);
  return json;
}

// TODO: bytt ut med domenene siden din faktisk kjører på.
const ALLOWED_ORIGIN_SUFFIXES = [".pages.dev", ".workers.dev"];
const ALLOWED_EXACT_ORIGINS = [];

function isAllowedOrigin(origin) {
  if (!origin) return false;
  if (origin.startsWith("http://localhost:") || origin.startsWith("http://127.0.0.1:")) {
    return true;
  }
  if (ALLOWED_EXACT_ORIGINS.includes(origin)) return true;
  try {
    const url = new URL(origin);
    if (url.protocol !== "https:") return false;
    return ALLOWED_ORIGIN_SUFFIXES.some((suffix) => url.hostname.endsWith(suffix));
  } catch {
    return false;
  }
}

function corsHeaders(origin) {
  if (!isAllowedOrigin(origin)) return {};
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    Vary: "Origin",
  };
}

function jsonResponse(body, status, origin) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders(origin) },
  });
}

function formatNOK(value) {
  const n = Number(value);
  if (!value || Number.isNaN(n)) return null;
  return "kr " + new Intl.NumberFormat("nb-NO").format(n);
}

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function buildSlackPayload(data) {
  const pills = [
    data.invoiceFrequency && `Fakturering: ${data.invoiceFrequency}`,
    data.customerCount && `${data.customerCount} ansatt${data.customerCount === "1" ? "" : "e"}`,
  ].filter(Boolean);

  const fields = [
    ["Org.nr", data.orgnr],
    ["Brreg/Proff", data.brregUrl ? `<${data.brregUrl}|${data.brregUrl}>` : null],
    ["Avtaledato", data.contractDate],
    ["Fakturering", data.invoiceDate],
    ["Fakt. frekvens", data.invoiceFrequency],
    ["Sum", formatNOK(data.fee)],
    ["Antall ansatte", data.customerCount],
    ["Kontakt hos kunde", [data.clientContact, data.clientEmail, data.clientPhone].filter(Boolean).join(" · ") || null],
    ["Intern kontaktperson", data.opkContact],
    ["Annen info", data.otherInfo],
  ]
    .filter(([, value]) => value)
    .map(([label, value]) => ({ type: "mrkdwn", text: `*${label}:*\n${value}` }));

  return {
    text: `🩺 Ny kunde: ${data.companyName || "Ukjent selskap"}`,
    blocks: [
      { type: "header", text: { type: "plain_text", text: `🩺 Ny kunde: ${data.companyName || "Ukjent selskap"}`, emoji: true } },
      ...(pills.length ? [{ type: "context", elements: pills.map((p) => ({ type: "mrkdwn", text: p })) }] : []),
      { type: "divider" },
      ...chunk(fields, 10).map((fieldGroup) => ({ type: "section", fields: fieldGroup })),
      {
        type: "actions",
        elements: [
          { type: "button", text: { type: "plain_text", text: "📋 Kopier", emoji: true }, action_id: "copy" },
          {
            type: "button",
            text: { type: "plain_text", text: "✅ Overfør til CRM", emoji: true },
            action_id: "added_to_crm",
            value: JSON.stringify(data),
          },
        ],
      },
    ],
  };
}

async function handleLookup(request, env, origin) {
  if (request.method !== "GET") {
    return jsonResponse({ error: "Method not allowed" }, 405, origin);
  }

  const url = new URL(request.url);
  const orgnr = (url.searchParams.get("orgnr") || "").replace(/\D/g, "");
  if (orgnr.length !== 9) {
    return jsonResponse({ error: "Organisasjonsnummer må ha 9 siffer" }, 400, origin);
  }

  const enhetRes = await fetch(`https://data.brreg.no/enhetsregisteret/api/enheter/${orgnr}`, {
    headers: { Accept: "application/json" },
  });
  if (!enhetRes.ok) {
    return jsonResponse(
      { error: enhetRes.status === 404 ? "Fant ikke selskapet i Brreg — sjekk organisasjonsnummeret" : "Klarte ikke å hente fra Brreg" },
      enhetRes.status === 404 ? 404 : 502,
      origin
    );
  }
  const enhet = await enhetRes.json();

  let dagligLeder = null;
  try {
    const roleRes = await fetch(`https://data.brreg.no/enhetsregisteret/api/enheter/${orgnr}/roller`, {
      headers: { Accept: "application/json" },
    });
    if (roleRes.ok) {
      const roller = await roleRes.json();
      const dagligLederGroup = (roller.rollegrupper || []).find((g) => g.type?.kode === "DAGL");
      const dagligLederRole = dagligLederGroup?.roller?.find((r) => !r.avregistrert) || dagligLederGroup?.roller?.[0];
      if (dagligLederRole?.person?.navn) {
        dagligLeder = `${dagligLederRole.person.navn.fornavn || ""} ${dagligLederRole.person.navn.etternavn || ""}`.trim();
      }
    }
  } catch {
    // roller-oppslag er best-effort — ignorer feil her
  }

  const employees =
    enhet.harRegistrertAntallAnsatte && typeof enhet.antallAnsatte === "number" ? enhet.antallAnsatte : null;
  const companyEmail = enhet.epostadresse || null;
  const companyPhone = enhet.telefon || enhet.mobil || null;
  const orgForm = enhet.organisasjonsform?.kode || null;

  return jsonResponse(
    {
      ok: true,
      name: enhet.navn,
      brregUrl: `https://www.proff.no/selskap/-/-/-/${orgnr}`,
      employees,
      dagligLeder,
      companyEmail,
      companyPhone,
      orgForm,
    },
    200,
    origin
  );
}

async function handleSearch(request, env, origin) {
  if (request.method !== "GET") {
    return jsonResponse({ error: "Method not allowed" }, 405, origin);
  }

  const url = new URL(request.url);
  const navn = (url.searchParams.get("navn") || "").trim();
  if (navn.length < 2) {
    return jsonResponse({ error: "Skriv minst 2 tegn" }, 400, origin);
  }

  const res = await fetch(`https://data.brreg.no/enhetsregisteret/api/enheter?navn=${encodeURIComponent(navn)}&size=8`, {
    headers: { Accept: "application/json" },
  });
  if (!res.ok) {
    return jsonResponse({ error: "Klarte ikke å søke i Brreg" }, 502, origin);
  }
  const json = await res.json();
  const results = (json._embedded?.enheter || []).map((e) => ({
    name: e.navn,
    orgnr: e.organisasjonsnummer,
    municipality: e.forretningsadresse?.kommune || e.postadresse?.kommune || null,
    orgForm: e.organisasjonsform?.beskrivelse || null,
  }));

  return jsonResponse({ ok: true, results }, 200, origin);
}

// Sender via rå SMTP (STARTTLS) med Cloudflare Workers sitt TCP Sockets-API — funker mot
// enhver vanlig e-postleverandør (Domeneshop, One.com, GoDaddy, ...), ikke bare Microsoft 365 /
// Google Workspace. Sett SMTP_HOST/SMTP_PORT/SMTP_USER/SMTP_PASSWORD som secrets.
async function readSmtpResponse(reader) {
  let buffer = "";
  const decoder = new TextDecoder();
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split("\r\n").filter(Boolean);
    const last = lines[lines.length - 1];
    if (last && /^\d{3} /.test(last)) break;
  }
  return buffer;
}

async function writeSmtpCommand(writer, command) {
  await writer.write(new TextEncoder().encode(command + "\r\n"));
}

function base64EncodeString(str) {
  return btoa(unescape(encodeURIComponent(str)));
}

// Bygger selve MIME-meldingen (headere + body, med \r\n gjennomgående). Brukes både av SMTP
// (sending) og IMAP APPEND (lagre som utkast) — de to transportene trenger identiske bytes,
// bare ulik "innpakning" (SMTP DATA-kommando vs. IMAP-literal).
function buildEmailMime({ from, to, cc = [], subject, text, html, inlineImages = [], attachments = [] }) {
  const normalize = (s) => s.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
  const normalizedText = normalize(text);

  // Bygges nedenfra og opp: text(+html) -> (+innebygde bilder) -> (+vedlegg), slik at hvert
  // nivå bare er til stede når det faktisk trengs (velkomstmail bruker html+bilde, ukentlig
  // backup-mail bruker bare text+vedlegg).
  let bodyMime;
  if (html) {
    const altBoundary = "opk-alt-" + Date.now();
    bodyMime =
      `Content-Type: multipart/alternative; boundary="${altBoundary}"\r\n\r\n` +
      `--${altBoundary}\r\n` +
      'Content-Type: text/plain; charset="UTF-8"\r\n\r\n' +
      `${normalizedText}\r\n\r\n` +
      `--${altBoundary}\r\n` +
      'Content-Type: text/html; charset="UTF-8"\r\n\r\n' +
      `${normalize(html)}\r\n\r\n` +
      `--${altBoundary}--\r\n`;
  } else {
    bodyMime = 'Content-Type: text/plain; charset="UTF-8"\r\n\r\n' + `${normalizedText}\r\n`;
  }

  let relatedMime = bodyMime;
  if (inlineImages.length) {
    const relBoundary = "opk-rel-" + Date.now();
    let rel = `Content-Type: multipart/related; boundary="${relBoundary}"\r\n\r\n`;
    rel += `--${relBoundary}\r\n${bodyMime}\r\n`;
    for (const img of inlineImages) {
      rel += `--${relBoundary}\r\n`;
      rel += `Content-Type: ${img.contentType}\r\n`;
      rel += "Content-Transfer-Encoding: base64\r\n";
      rel += `Content-ID: <${img.cid}>\r\n`;
      rel += `Content-Disposition: inline; filename="${img.filename}"\r\n\r\n`;
      rel += `${img.contentBase64.replace(/(.{76})/g, "$1\r\n")}\r\n\r\n`;
    }
    rel += `--${relBoundary}--\r\n`;
    relatedMime = rel;
  }

  let contentMime = relatedMime;
  if (attachments.length) {
    const mixBoundary = "opk-mix-" + Date.now();
    let mix = `Content-Type: multipart/mixed; boundary="${mixBoundary}"\r\n\r\n`;
    mix += `--${mixBoundary}\r\n${relatedMime}\r\n`;
    for (const att of attachments) {
      mix += `--${mixBoundary}\r\n`;
      mix += `Content-Type: ${att.contentType}; name="${att.filename}"\r\n`;
      mix += "Content-Transfer-Encoding: base64\r\n";
      mix += `Content-Disposition: attachment; filename="${att.filename}"\r\n\r\n`;
      mix += `${att.contentBase64.replace(/(.{76})/g, "$1\r\n")}\r\n\r\n`;
    }
    mix += `--${mixBoundary}--\r\n`;
    contentMime = mix;
  }

  let mime = "";
  mime += `From: ${from}\r\n`; // TODO: evt. sett et visningsnavn, f.eks. `"Ditt Firma" <${from}>`
  mime += `To: ${to}\r\n`;
  if (cc.length) mime += `Cc: ${cc.join(", ")}\r\n`;
  mime += `Subject: ${subject}\r\n`;
  mime += "MIME-Version: 1.0\r\n";
  mime += contentMime;
  return mime;
}

async function sendSmtpMail(env, { to, cc = [], subject, text, html, inlineImages = [], attachments = [] }) {
  const host = env.SMTP_HOST;
  const port = Number(env.SMTP_PORT) || 587;
  const from = env.SMTP_USER;

  let socket = connect({ hostname: host, port }, { secureTransport: "starttls" });
  let writer = socket.writable.getWriter();
  let reader = socket.readable.getReader();

  await readSmtpResponse(reader); // 220 greeting

  await writeSmtpCommand(writer, `EHLO ${(env.SMTP_USER || "localhost").split("@")[1] || "localhost"}`);
  await readSmtpResponse(reader);

  await writeSmtpCommand(writer, "STARTTLS");
  const startTlsResp = await readSmtpResponse(reader);
  if (!startTlsResp.startsWith("220")) throw new Error("STARTTLS ble avvist: " + startTlsResp.slice(0, 200));

  writer.releaseLock();
  reader.releaseLock();
  socket = socket.startTls();
  writer = socket.writable.getWriter();
  reader = socket.readable.getReader();

  await writeSmtpCommand(writer, `EHLO ${(env.SMTP_USER || "localhost").split("@")[1] || "localhost"}`);
  await readSmtpResponse(reader);

  await writeSmtpCommand(writer, "AUTH LOGIN");
  await readSmtpResponse(reader);
  await writeSmtpCommand(writer, base64EncodeString(env.SMTP_USER));
  await readSmtpResponse(reader);
  await writeSmtpCommand(writer, base64EncodeString(env.SMTP_PASSWORD));
  const authResp = await readSmtpResponse(reader);
  if (!authResp.startsWith("235")) throw new Error("SMTP-autentisering feilet: " + authResp.slice(0, 200));

  await writeSmtpCommand(writer, `MAIL FROM:<${from}>`);
  await readSmtpResponse(reader);

  const allRecipients = [to, ...cc].filter(Boolean);
  for (const rcpt of allRecipients) {
    await writeSmtpCommand(writer, `RCPT TO:<${rcpt}>`);
    const rcptResp = await readSmtpResponse(reader);
    if (!rcptResp.startsWith("250")) throw new Error(`Mottaker ${rcpt} ble avvist: ` + rcptResp.slice(0, 200));
  }

  await writeSmtpCommand(writer, "DATA");
  await readSmtpResponse(reader);

  const mime = buildEmailMime({ from, to, cc, subject, text, html, inlineImages, attachments });
  const escaped = mime.replace(/\r\n\./g, "\r\n.."); // dot-stuffing
  await writer.write(new TextEncoder().encode(escaped + "\r\n.\r\n"));
  const dataResp = await readSmtpResponse(reader);
  if (!dataResp.startsWith("250")) throw new Error("SMTP-sending feilet: " + dataResp.slice(0, 200));

  await writeSmtpCommand(writer, "QUIT");
  try {
    await readSmtpResponse(reader);
  } catch {
    /* ignorer — meldingen er allerede sendt */
  }

  writer.releaseLock();
  reader.releaseLock();
  await socket.close();
}

function hasImapLineStartingWith(buffer, prefix) {
  return buffer.split("\r\n").some((line) => line.startsWith(prefix));
}

async function readImapUntil(reader, predicate) {
  let buffer = "";
  const decoder = new TextDecoder();
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    if (predicate(buffer)) break;
  }
  return buffer;
}

// Lagrer meldingen direkte i "Drafts"-mappen via IMAP APPEND, i stedet for å sende den —
// brukes for velkomstmailen, som skal kunne gjennomgås før noen trykker send selv (i vanlig
// e-postklient). TODO: "Drafts" er mappenavnet — noen leverandører bruker "INBOX.Drafts" e.l.,
// sjekk med en test-kjøring om APPEND blir avvist.
async function saveEmailDraftViaImap(env, { to, cc = [], subject, text, html, inlineImages = [] }) {
  const host = env.IMAP_HOST;
  const port = Number(env.IMAP_PORT) || 993;
  const from = env.SMTP_USER;

  const socket = connect({ hostname: host, port }, { secureTransport: "on" });
  const writer = socket.writable.getWriter();
  const reader = socket.readable.getReader();

  await readImapUntil(reader, (buf) => buf.includes("\r\n")); // * OK-hilsen

  const escapeImapQuoted = (s) => s.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  await writer.write(
    new TextEncoder().encode(`A1 LOGIN "${escapeImapQuoted(env.SMTP_USER)}" "${escapeImapQuoted(env.SMTP_PASSWORD)}"\r\n`)
  );
  const loginResp = await readImapUntil(reader, (buf) => hasImapLineStartingWith(buf, "A1 "));
  if (!hasImapLineStartingWith(loginResp, "A1 OK")) throw new Error("IMAP-innlogging feilet: " + loginResp.slice(0, 300));

  const mime = buildEmailMime({ from, to, cc, subject, text, html, inlineImages });
  const bytes = new TextEncoder().encode(mime);

  await writer.write(new TextEncoder().encode(`A2 APPEND "Drafts" (\\Draft) {${bytes.length}}\r\n`));
  const contResp = await readImapUntil(
    reader,
    (buf) => hasImapLineStartingWith(buf, "+") || hasImapLineStartingWith(buf, "A2 ")
  );
  if (!hasImapLineStartingWith(contResp, "+")) {
    throw new Error("IMAP APPEND (Drafts) ble avvist: " + contResp.slice(0, 300));
  }

  await writer.write(bytes);
  await writer.write(new TextEncoder().encode("\r\n"));
  const appendResp = await readImapUntil(reader, (buf) => hasImapLineStartingWith(buf, "A2 "));
  if (!hasImapLineStartingWith(appendResp, "A2 OK")) throw new Error("IMAP APPEND feilet: " + appendResp.slice(0, 300));

  await writer.write(new TextEncoder().encode("A3 LOGOUT\r\n"));
  try {
    await readImapUntil(reader, (buf) => hasImapLineStartingWith(buf, "A3 "));
  } catch {
    /* ignorer — utkastet er allerede lagret */
  }

  writer.releaseLock();
  reader.releaseLock();
  await socket.close();
}

function escapeHtml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Speiler linjeskiftene i ren tekst nøyaktig (hver \n -> <br>) i stedet for å tolke dem som
// "avsnitt" med fast avstand — da følger antall tomme linjer brukeren skriver inn direkte.
function textToHtmlParagraphs(text) {
  return `<div>${escapeHtml(text).replace(/\n/g, "<br>\n")}</div>`;
}

const COMPANY_LOGO_CID = "companylogo";

// TODO: adresse/telefon under er fortsatt plassholdere — bytt ut med BHTNs faktiske
// kontaktinfo. Denne må holdes i sync med signaturblokken nederst i buildWelcomeEmail
// (frontend/index.html) — se signatureMarker-kommentaren i sendWelcomeEmail under for hvorfor.
function buildSignatureHtml() {
  return `<table cellpadding="0" cellspacing="0" border="0" style="margin-top:12px;">
  <tr>
    <td style="padding-right:14px;vertical-align:top;">
      <img src="cid:${COMPANY_LOGO_CID}" alt="Logo" width="70" style="display:block;">
    </td>
    <td style="vertical-align:top;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;line-height:1.5;">
      <strong>BHTN</strong><br>
      Gateadresse 1, 0000 Sted<br>
      Tlf: 00 00 00 00<br>
      <a href="mailto:post@bhtn.no" style="color:#0563C1;">post@bhtn.no</a><br>
      <a href="https://bhtn.no/" style="color:#0563C1;">https://bhtn.no/</a>
    </td>
  </tr>
</table>`;
}

// "body" fra skjemaet slutter alltid med "...Med vennlig hilsen [noen]\n\nBHTN\n..."
// (se buildWelcomeEmail i frontend). Alt FØR signatureMarker (inkludert selve "Med vennlig
// hilsen"-linjen og brukerens egne linjeskift) rendres 1:1 som vanlig tekst, slik at avstanden
// brukeren har satt opp beholdes nøyaktig. Bare selve adresseblokken fra og med markøren bygges
// om til en tabell med innebygd logo.
// TODO: signatureMarker MÅ matche starten på signaturblokken i buildWelcomeEmail (frontend) —
// og teksten i buildSignatureHtml over — ellers kuttes greeting-teksten feil sted.
async function sendWelcomeEmail(env, { to, cc, subject, body }) {
  const signatureMarker = "BHTN";
  const markerIdx = body.indexOf(signatureMarker);
  const greeting = markerIdx === -1 ? body : body.slice(0, markerIdx).replace(/\n+$/, "");

  const html = `<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#111111;">${textToHtmlParagraphs(
    greeting
  )}${buildSignatureHtml()}</div>`;

  await saveEmailDraftViaImap(env, {
    to,
    cc: Array.isArray(cc) ? cc : [],
    subject,
    text: body,
    html,
    inlineImages: [{ cid: COMPANY_LOGO_CID, filename: "logo.png", contentType: "image/png", contentBase64: LOGO_BASE64 }],
  });
}

async function handleWelcomeEmail(request, env, origin) {
  if (request.method !== "POST") {
    return jsonResponse({ error: "Method not allowed" }, 405, origin);
  }

  let payload;
  try {
    payload = await request.json();
  } catch {
    return jsonResponse({ error: "Ugyldig JSON" }, 400, origin);
  }

  if (!payload.to || !payload.subject || !payload.body) {
    return jsonResponse({ error: "to, subject og body er påkrevd" }, 400, origin);
  }

  if (!env.IMAP_HOST || !env.SMTP_USER || !env.SMTP_PASSWORD) {
    return jsonResponse({ error: "E-postutsending er ikke konfigurert ennå" }, 500, origin);
  }

  try {
    await sendWelcomeEmail(env, payload);
    return jsonResponse({ ok: true }, 200, origin);
  } catch (e) {
    return jsonResponse({ error: e.message || "Kunne ikke sende e-post" }, 502, origin);
  }
}

async function handleSlackSubmit(request, env, origin) {
  if (request.method !== "POST") {
    return jsonResponse({ error: "Method not allowed" }, 405, origin);
  }

  let data;
  try {
    data = await request.json();
  } catch {
    return jsonResponse({ error: "Ugyldig JSON" }, 400, origin);
  }

  if (!data.companyName) {
    return jsonResponse({ error: "companyName er påkrevd" }, 400, origin);
  }

  if (!env.SLACK_WEBHOOK_URL) {
    return jsonResponse({ error: "Server er ikke konfigurert" }, 500, origin);
  }

  const slackResponse = await fetch(env.SLACK_WEBHOOK_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(buildSlackPayload(data)),
  });

  if (!slackResponse.ok) {
    const slackErrorText = await slackResponse.text().catch(() => "");
    return jsonResponse(
      { error: `Kunne ikke sende til Slack (${slackResponse.status}): ${slackErrorText.slice(0, 200) || "ukjent årsak"}` },
      502,
      origin
    );
  }

  return jsonResponse({ ok: true }, 200, origin);
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return result === 0;
}

async function verifySlackSignature(request, env) {
  const timestamp = request.headers.get("X-Slack-Request-Timestamp");
  const signature = request.headers.get("X-Slack-Signature");
  const body = await request.text();

  if (!timestamp || !signature) return { valid: false, body };

  const fiveMinutesAgo = Math.floor(Date.now() / 1000) - 60 * 5;
  if (Number(timestamp) < fiveMinutesAgo) return { valid: false, body };

  const base = `v0:${timestamp}:${body}`;
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(env.SLACK_SIGNING_SECRET),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sigBuffer = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(base));
  const hex = Array.from(new Uint8Array(sigBuffer))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
  const computed = `v0=${hex}`;

  return { valid: timingSafeEqual(computed, signature), body };
}

async function processAddedToCrm(payload, data, env) {
  const sheetsConfigured = env.GOOGLE_APPS_SCRIPT_URL && env.GOOGLE_APPS_SCRIPT_SECRET;
  let resultText;
  if (!sheetsConfigured) {
    resultText = "⚠️ Google Sheets er ikke konfigurert ennå.";
  } else {
    try {
      const result = await syncViaAppsScript(data, env);
      const who = payload.user?.username || payload.user?.name || "noen";
      if (result.duplicate) {
        resultText = `⚠️ "${data.companyName || "?"}" finnes allerede i "${SHEET_TAB}", rad ${result.row}. Sjekk om dette er riktig — ingen ny rad ble lagt til.`;
      } else {
        resultText = `✅ Lagt til i CRM av ${who}.`;
      }
    } catch (e) {
      resultText = `⚠️ Klarte ikke å legge til i CRM: ${(e.message || "ukjent feil").slice(0, 200)}`;
    }
  }

  if (payload.response_url) {
    try {
      const originalBlocks = payload.message?.blocks || [];
      const updatedBlocks = originalBlocks
        .filter((b) => b.type !== "actions")
        .concat({ type: "context", elements: [{ type: "mrkdwn", text: resultText }] });

      await fetch(payload.response_url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ replace_original: true, text: payload.message?.text || "", blocks: updatedBlocks }),
      });
    } catch {
      // Meldingen ble ikke oppdatert visuelt, men selve CRM-tillegget skjedde likevel — ikke la dette feile bakgrunnsjobben
    }
  }
}

async function handleSlackInteractivity(request, env, ctx) {
  if (request.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }
  if (!env.SLACK_SIGNING_SECRET) {
    return new Response("Not configured", { status: 500 });
  }

  const { valid, body } = await verifySlackSignature(request, env);
  if (!valid) {
    return new Response("Invalid signature", { status: 401 });
  }

  const params = new URLSearchParams(body);
  const payloadRaw = params.get("payload");
  if (!payloadRaw) return new Response("Missing payload", { status: 400 });

  let payload;
  try {
    payload = JSON.parse(payloadRaw);
  } catch {
    return new Response("Bad payload", { status: 400 });
  }

  if (payload.type !== "block_actions") return new Response("", { status: 200 });

  const action = payload.actions?.[0];
  if (!action || action.action_id !== "added_to_crm") {
    return new Response("", { status: 200 }); // andre knapper (Kopier, Til <system>) er kun visuelle
  }

  let data = {};
  try {
    data = JSON.parse(action.value);
  } catch {
    // ignorer — data blir tom, syncAllTabs vil feile pent
  }

  // Slack krever svar innen 3 sek. Skriving til 10 faner tar lenger enn det,
  // så vi svarer med en gang og gjør selve jobben i bakgrunnen, og oppdaterer
  // meldingen via response_url når den er ferdig.
  ctx.waitUntil(processAddedToCrm(payload, data, env));

  return new Response("", { status: 200 });
}

export default {
  async fetch(request, env, ctx) {
    const origin = request.headers.get("Origin");
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders(origin) });
    }

    if (url.pathname === "/welcome-email") {
      return handleWelcomeEmail(request, env, origin);
    }

    if (url.pathname === "/lookup") {
      return handleLookup(request, env, origin);
    }

    if (url.pathname === "/search") {
      return handleSearch(request, env, origin);
    }

    if (url.pathname === "/slack-interactivity") {
      return handleSlackInteractivity(request, env, ctx);
    }



    return handleSlackSubmit(request, env, origin);
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(
      runWeeklyBackup(env).catch((e) => {
        console.error("Ukentlig backup feilet:", e.message);
      })
    );
  },
};
