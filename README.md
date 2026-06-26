# Tenant Guard

An Entra security & compliance console. A single-file, backend-less web app (like CloudCapsule)
that — to start — lists **dormant members** and **dormant guests** who haven't signed in beyond
your chosen window, counting **both interactive and non-interactive** sign-ins, and lets an admin
**delete** them. Built to grow into a broader tenant-hygiene toolkit over time. Runs entirely
client-side (MSAL.js + Microsoft Graph), so it hosts on **GitHub Pages** with no server and no secrets.

---

## How sign-in activity is judged

For every user, the app reads `signInActivity` and takes the **most recent** of:

- `lastSignInDateTime` — interactive
- `lastNonInteractiveSignInDateTime` — non-interactive
- `lastSuccessfulSignInDateTime` — last success

"Days inactive" = today minus that most-recent timestamp. Users with no sign-in ever are flagged
as **never** (optionally, using account age as a proxy so brand-new accounts aren't swept in).

> Requires an **Entra ID P1 or P2** license in the tenant. Without it, Graph returns
> `Authentication_RequestFromNonPremiumTenantOrB2CTenant` and the app shows "license: none".

---

## One-time: create the app registration (you, the publisher)

Do this **once** in your own tenant. Every other organization just signs in and consents.

1. **Entra admin center → App registrations → New registration.**
   - Name: `Tenant Guard`
   - Supported account types: **Accounts in any organizational directory (multitenant)**
   - Redirect URI → platform **Single-page application (SPA)** → your GitHub Pages URL,
     e.g. `https://YOURNAME.github.io/tenant-guard/` (exact, with trailing slash).
2. **API permissions → Add → Microsoft Graph → Delegated**, add:
   - `User.Read`
   - `User.Read.All`
   - `AuditLog.Read.All`
   - `Directory.Read.All`
   - `User.ReadWrite.All`
3. Copy the **Application (client) ID**.
4. In `index.html`, set `DEFAULT_CLIENT_ID` to that GUID. (Or leave it and let each org paste their
   own client ID under **Settings** — see "BYO app registration" below.)

That's it. There's **no client secret** — a SPA uses PKCE, which is why it can live on static hosting.

---

## Host on GitHub Pages

1. Create a repo, e.g. `tenant-guard`. Add `index.html` (and this `README.md`) to the `main` branch.
2. **Repo → Settings → Pages → Build and deployment → Deploy from a branch.**
   Branch: `main`, folder: `/ (root)` → **Save**.
3. Wait ~1 minute. Your site is at `https://YOURNAME.github.io/tenant-guard/`.
4. Make sure that **exact** URL is registered as the SPA redirect URI (step 1 above). The app
   computes its own redirect URI from the current page and shows it under **Settings** if you need
   to copy it.

Custom domain works too — just register that URL as the redirect URI instead.

---

## How any organization onboards (the "creates the app on its own" part)

You do **not** create an app registration per customer. The multi-tenant registration above is the
only one. When another org's admin uses your site:

1. They open `https://YOURNAME.github.io/tenant-guard/` and click **Sign in with Microsoft**.
2. A **Global Administrator** is shown the consent screen for the permissions above.
3. Granting **admin consent** automatically provisions the **enterprise application
   (service principal)** in *their* tenant — Entra creates it for them. No manual app setup.

If consent wasn't granted up front, the in-app **Settings → Grant admin consent (one-time)** button
re-runs the admin consent flow.

### BYO app registration (optional, for orgs that won't trust a shared app)
An organization that prefers its app registration to live in **its own** tenant can register its own
multi-tenant SPA app (same steps, same permissions, their own redirect URI) and paste that client ID
into **Settings → Application (client) ID → Save & reload**. The page stores it locally in the
browser; nothing else changes.

> Note: an app cannot bootstrap-create its *own* registration from a cold start (you'd need an app to
> call Graph in the first place). Admin consent provisioning the enterprise app is the correct,
> supported "self-service" mechanism — that's what this uses.

---

## Who can run it

- **Read / list dormant accounts:** any signed-in admin with the consented read scopes
  (Reports Reader is enough for the data).
- **Delete accounts:** the signed-in user needs a directory role that can delete users —
  **User Administrator** or **Global Administrator**. `User.ReadWrite.All` alone isn't enough without
  the role.

---

## Delete behaviour & safety

- Deletes go through Graph **`$batch`** (20 per call) with **429 / Retry-After** handling.
- Deletion is a **soft-delete** — accounts sit in **Deleted users** and are recoverable for **30 days**.
- A typed `DELETE` confirmation is required; your **own account is always excluded**.
- **Export CSV** first if you want an audit record of what was removed.

---

## Quick local test

```bash
# from the repo folder
python -m http.server 8080
# open http://localhost:8080  — register http://localhost:8080/ as an extra SPA redirect URI for testing
```

---

## Files

| File | Purpose |
|---|---|
| `index.html` | The entire app — UI, MSAL auth, Graph calls, delete flow. |
| `README.md` | This guide. |
