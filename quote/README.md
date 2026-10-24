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
- **There is no order book in the customer's page.** Orders are recorded, never listed back:
  one customer has no business seeing another's jobs. The listing lives in the back office.

## The back office — `/quote/admin/`

`admin/index.html` is the DEMA side: every order in the sheet, newest first, with the
counts, a search, a filter by solution, a row that opens onto the billing block and the
brief, and a CSV of whatever is on screen. It is the same board as the desk, read instead
of filled, and it **only reads** — nothing on that page can change an order.

It asks the Apps Script for the book by posting `{action: 'orders', token}` to the very
same `/exec`. Two things make that work, and both are easy to undo by accident:

- The token lives in **Project Settings › Script properties**, under `ADMIN_TOKEN`, never
  in this repository. **With no property set the endpoint stays shut**, so a half-finished
  deployment cannot hand the order book to the open internet.
- The token travels in the **body**, not the query string: a URL ends up in browser history
  and in referrers. It cannot travel in a custom header either — that would force a CORS
  preflight, and an Apps Script web app answers no `OPTIONS`.

The page is `noindex`, but its address is not a secret: the token is the lock. To change
it, set a new value in Script properties — every browser then simply asks once more.

**After editing `backend/quote-apps-script.gs` it has to be redeployed**, or the live
`/exec` goes on running the old code and the back office reads nothing: *Deploy › Manage
deployments › pencil › New version*. Never "New deployment" — that mints a different URL
and the quote page stops reaching it.

## Euro is the list currency

Dollar prices convert at a fixed reference rate stated on the page, because a static file
cannot call a currency feed. The source project `translation-quote-app` reads the live ECB
rate; this page cannot, so **check that rate from time to time** — it is `USD_PER_EUR` in
`index.html`, with the date it was taken beside it.
