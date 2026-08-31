# Onboarding sirkulær helse — oppsett og deploy

Kundeonboarding-skjema: en statisk React-side (ingen build-steg) som poster en
formatert melding til en Slack-kanal via en Cloudflare Worker. Fra Slack kan
noen trykke en knapp for å synkronisere kunden inn i et Google Sheets-ark.
Verktøyet kan i tillegg sende en velkomstmail (lagres som utkast, ikke
sendes automatisk) og ta en ukentlig sikkerhetskopi av arket.

Dette er en genericisert kopi av et tidligere prosjekt — se `TODO`-kommentarer
i kildekoden for alt som må tilpasses før dette kan settes i produksjon.

## Struktur

```
frontend/
  index.html              Selve skjemaet, React lastet via CDN
  functions/
    _middleware.js         Passordluke (Cloudflare Pages Function)
  assets/logo-color.png    TODO: bytt ut med din egen logo
worker/
  src/index.js             Slack, Google Sheets, Brreg-oppslag, e-post, backup
  wrangler.toml             Worker-konfigurasjon (cron, KV-namespace)
  .dev.vars.example         Mal for lokale secrets — kopier til .dev.vars og fyll inn
```

## Forutsetninger

- Cloudflare-konto
- Node.js installert (brukes til `npx wrangler`)
- En Slack-app (ikke bare en enkel Incoming Webhook — se steg 2, siden
  "Overfør til CRM"-knappen krever ekte interaktivitet)
- Et Google Sheets-dokument + en Google Cloud-tjenestekonto med tilgang til det
- En e-postkonto med SMTP + IMAP (vanlig hos de fleste domeneleverandører)

## 1. Logg inn med Wrangler

```bash
npx wrangler login
```

## 2. Slack-app

En enkel Incoming Webhook holder for å POSTE meldinger, men **knappen som
overfører kunden til CRM krever en fullverdig Slack-app** med Interactivity
aktivert:

1. Opprett en app på https://api.slack.com/apps
2. Aktiver Incoming Webhooks, hent webhook-URL-en for riktig kanal
3. Aktiver Interactivity & Shortcuts, sett Request URL til
   `https://<din-worker>.workers.dev/slack-interactivity` (URL-en finner du
   etter steg 3 under)
4. Noter Signing Secret fra "Basic Information"

## 3. Deploy Worker

```bash
cd worker
cp .dev.vars.example .dev.vars   # fyll inn ekte verdier lokalt for testing
npx wrangler secret put SLACK_WEBHOOK_URL
npx wrangler secret put SLACK_SIGNING_SECRET
npx wrangler secret put GOOGLE_SERVICE_ACCOUNT_EMAIL
npx wrangler secret put GOOGLE_SERVICE_ACCOUNT_KEY
npx wrangler secret put GOOGLE_SHEET_ID
npx wrangler secret put SMTP_HOST
npx wrangler secret put SMTP_PORT
npx wrangler secret put SMTP_USER
npx wrangler secret put SMTP_PASSWORD
npx wrangler secret put IMAP_HOST
npx wrangler secret put IMAP_PORT
npx wrangler secret put BACKUP_RECIPIENT
npx wrangler kv namespace create BACKUP_STATE   # lim inn ID-en i wrangler.toml
npx wrangler deploy
```

Worker-navnet er allerede satt til `sirkular-helse-onboarding-worker` i
`wrangler.toml`, så URL-en blir `https://sirkular-helse-onboarding-worker.<konto>.workers.dev`.
Noter denne — den trengs i steg 5.

## 4. Google Sheets — VIKTIG, må kodes for ditt ark

`worker/src/index.js` inneholder en **eksempel**-versjon av `syncToSheet` og
`SECONDARY_TABS`, ikke en ferdig løsning. Google Sheets-synken er
skreddersøm som må gjøres for akkurat ditt ark:

1. Del arket med tjenestekontoens e-post (Editor-tilgang)
2. Les overskriftsraden i hovedfanen for å se eksakte kolonnenavn
3. Match skjemafeltene mot kolonnene i `byHeader`-oppsettet i `syncToSheet`
4. Hvis du har flere faner som skal få kundens org.nummer/nøkkel skrevet inn
   automatisk: legg dem til i `SECONDARY_TABS`, med samme fremgangsmåte som
   er beskrevet i kommentarene der (les overskrift + en eksempelrad med
   `valueRenderOption=FORMULA` for å finne eksakte formler)

## 5. Koble skjemaet til Worker-en

Åpne `frontend/index.html`, finn:

```js
const SLACK_ENDPOINT = "https://din-worker.din-konto.workers.dev";
```

og erstatt med URL-en fra steg 3.

## 6. Deploy skjemaet til Cloudflare Pages

```bash
cd frontend
npx wrangler pages project create sirkular-helse-onboarding
npx wrangler pages deploy . --project-name=sirkular-helse-onboarding
npx wrangler pages secret put SITE_PASSWORD --project-name=sirkular-helse-onboarding
```

Uten `SITE_PASSWORD` viser siden en "ikke konfigurert"-feil i stedet for
skjemaet. Passordet lagres i en `HttpOnly`-cookie i 30 dager.

## 7. E-post (velkomstmail + ukentlig backup)

Verktøyet sender via rå SMTP (STARTTLS) — funker med de fleste
domeneleverandører (Domeneshop, One.com, GoDaddy, ...), ikke bare Microsoft
365/Google Workspace. Velkomstmailen lagres som **utkast** (IMAP APPEND til
"Drafts") i stedet for å sendes automatisk, slik at noen kan gjennomgå den
først. Sjekk om leverandøren din bruker et annet mappenavn enn "Drafts".

Firmanavnet ("BHTN") er allerede satt i `buildSignatureHtml`
(worker) og signaturblokken i `buildWelcomeEmail` (frontend). Adresse/
telefon/e-post/nettside/logo er fortsatt plassholdere — bytt disse ut med
ekte info, og hold de to filene i sync (se `signatureMarker`-kommentaren i
`sendWelcomeEmail`).

Den ukentlige backupen eksporterer hele Google-arket som .xlsx og sender det
til `BACKUP_RECIPIENT`. Tidspunktet styres av `crons` i `wrangler.toml`
(cron-syntaks, UTC-tid).

## Lokal testing

```bash
cd worker
npx wrangler dev --port 8799
```

```bash
cd frontend
npx wrangler pages dev . --port 8788
```

## CORS

`ALLOWED_ORIGIN_SUFFIXES`/`ALLOWED_EXACT_ORIGINS` i `worker/src/index.js`
styrer hvilke domener som får lov til å kalle workeren. Legg til ditt
faktiske domene der.
