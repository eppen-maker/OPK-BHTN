const COOKIE_NAME = "onboarding_auth";
const MAX_AGE = 60 * 60 * 24 * 30; // 30 dager

function loginPage(showError) {
  return `<!DOCTYPE html>
<html lang="no">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Onboarding – Logg inn</title>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;800&display=swap" rel="stylesheet" />
<style>
  body { margin:0; min-height:100vh; display:flex; align-items:center; justify-content:center; background:#faf7f4; font-family:"Plus Jakarta Sans",sans-serif; }
  .box { background:#fff; border-radius:20px; box-shadow:0 4px 12px rgba(41,24,68,.08); padding:40px; width:100%; max-width:340px; }
  h1 { font-size:20px; font-weight:800; color:#291844; margin:0 0 8px; }
  p { color:#5a4a78; font-size:13px; margin:0 0 24px; }
  input { width:100%; box-sizing:border-box; padding:12px 16px; border-radius:12px; border:1.5px solid #f3eee8; font-size:14px; margin-bottom:16px; }
  button { width:100%; padding:12px; border:none; border-radius:9999px; background:#291844; color:#fff; font-weight:700; font-size:14px; cursor:pointer; }
  .error { color:#c4304a; font-size:13px; margin-bottom:16px; }
</style>
</head>
<body>
  <form class="box" method="POST" action="/_auth">
    <h1>Onboarding</h1>
    <p>Skriv inn passordet for å fortsette.</p>
    ${showError ? '<div class="error">Feil passord, prøv igjen.</div>' : ""}
    <input type="password" name="password" placeholder="Passord" autofocus required />
    <button type="submit">Logg inn</button>
  </form>
</body>
</html>`;
}

export async function onRequest(context) {
  const { request, env, next } = context;
  const url = new URL(request.url);

  if (url.pathname === "/_auth" && request.method === "POST") {
    const form = await request.formData();
    const password = form.get("password");

    if (!env.SITE_PASSWORD) {
      return new Response("Siden er ikke konfigurert (mangler SITE_PASSWORD).", { status: 500 });
    }

    if (password === env.SITE_PASSWORD) {
      const headers = new Headers({ Location: "/" });
      headers.append(
        "Set-Cookie",
        `${COOKIE_NAME}=${encodeURIComponent(env.SITE_PASSWORD)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${MAX_AGE}`
      );
      return new Response(null, { status: 302, headers });
    }

    return new Response(loginPage(true), { status: 401, headers: { "Content-Type": "text/html; charset=utf-8" } });
  }

  const cookieHeader = request.headers.get("Cookie") || "";
  const cookies = Object.fromEntries(
    cookieHeader.split(";").filter(Boolean).map((c) => {
      const [k, ...v] = c.trim().split("=");
      return [k, decodeURIComponent(v.join("="))];
    })
  );

  if (!env.SITE_PASSWORD || cookies[COOKIE_NAME] !== env.SITE_PASSWORD) {
    return new Response(loginPage(false), { status: 401, headers: { "Content-Type": "text/html; charset=utf-8" } });
  }

  return next();
}
