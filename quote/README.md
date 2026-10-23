# Quote Desk — `/quote`

The translation quoting and ordering page, served as a static file by GitHub Pages at
<https://ai-dema-solutions.com/quote/>.

Everything the page needs to price a job is inside `index.html`: the rate card (Listino
tecnico 2026), the working-calendar delivery rules, the volume breaks and the VAT rule.
`analysis.js` counts the words in uploaded files **in the browser** — nothing is uploaded
anywhere. The three `tier-*.webp` are the workflow drawings under each solution.

## Where the order goes

The page has no server. When a customer presses **Submit order** it posts the order to a
Google Apps Script web app — the same arrangement as the waitlist form on the home page —
which writes a row in a sheet and sends two emails: the confirmation with its PDF to the
customer, and the whole record to the desk.

That script is `backend/quote-apps-script.gs`, and its header explains how to deploy it.
Once it is deployed, paste its `/exec` URL into **`ORDERS_WEBHOOK`** near the top of the
script block in `index.html`:

```js
const ORDERS_WEBHOOK = 'https://script.google.com/macros/s/..../exec';
```

**Until that is filled in** — and whenever the call fails — the page hands the finished
order to the visitor's own mail client, pre-filled and addressed to `ORDERS_INBOX`. So the
page is useful from the moment it is live: nobody is ever left with nowhere to send a job.

## Two things that are deliberate

- **The prices do not move while you type.** The brief can be edited freely; nothing is
  re-priced until **Show prices** / **Update prices** is pressed. This was asked for.
- **There is no order book here.** Orders are recorded, never listed back: one customer has
  no business seeing another's jobs. The listing belongs in the DEMA admin interface.

## Euro is the list currency

Dollar prices convert at a fixed reference rate stated on the page, because a static file
cannot call a currency feed. The source project `translation-quote-app` reads the live ECB
rate; this page cannot, so **check that rate from time to time** — it is `USD_PER_EUR` in
`index.html`, with the date it was taken beside it.
