# DEMA Company Brain — Landing

Single-file landing page (waitlist) served by GitHub Pages on https://ai-dema-solutions.com

- `index.html` — the whole page (HTML + CSS + JS, no dependencies)
- `CNAME` — custom domain for GitHub Pages

## Backend (optional)

The form works without a backend: the counter is stored in the visitor's `localStorage`.
To collect the leads, set `WEBHOOK_URL` in `index.html` to a publicly reachable n8n webhook.
The endpoint receives `{ nome, azienda, email, settore, lang }` and must return `{ "taken": N }`.
