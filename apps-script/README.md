# Oppsett av CRM-overføring (uten privatnøkkel)

1. Åpne Google-arket ("KUNDELISTE BHTN") → **Extensions/Utvidelser → Apps Script**
2. Slett eventuell eksempelkode i editoren, lim inn hele innholdet fra `Code.gs` i denne mappen
3. Trykk **Deploy → New deployment**
4. Trykk tannhjulet ved siden av "Select type" → velg **Web app**
5. Under "Execute as": **Me**. Under "Who has access": **Anyone**
6. Trykk **Deploy**, godkjenn tilgangene den ber om
7. Kopier URL-en som vises (slutter på `/exec`)

Send URL-en til Claude (eller lim den selv inn som `GOOGLE_APPS_SCRIPT_URL`-secret på
Cloudflare Worker-en, sammen med `GOOGLE_APPS_SCRIPT_SECRET` = verdien av `SHARED_SECRET`
i `Code.gs`).

Ingen tjenestekonto, ingen privatnøkkel, ingen JSON-fil nødvendig.
