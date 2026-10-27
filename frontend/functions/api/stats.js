// Kjøres bak samme passordluke som resten av siden (frontend/functions/_middleware.js
// gjelder for alle ruter, inkludert denne) — ingen ekstra innlogging nødvendig her.
// Selve statistikk-hemmeligheten (STATS_SECRET) sendes bare server-til-server til Worker-en,
// aldri til nettleseren, slik at /stats-endepunktet ikke kan kalles direkte utenfra.
const STATS_WORKER_URL = "https://sirkular-helse-onboarding-worker.sirkular-helse-bhtn.workers.dev";

export async function onRequestGet(context) {
  const { env } = context;

  if (!env.STATS_SECRET) {
    return new Response(JSON.stringify({ error: "Mangler STATS_SECRET i Pages-miljøet" }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }

  const res = await fetch(`${STATS_WORKER_URL}/stats`, {
    headers: { "X-Stats-Secret": env.STATS_SECRET },
  });

  const body = await res.text();
  return new Response(body, {
    status: res.status,
    headers: { "Content-Type": "application/json" },
  });
}
