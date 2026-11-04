# Hand-painted mug pre-orders

A one-page order form for holiday mug pre-orders. Buyers pick designs, add names, and see their total live. Orders go to a Google Sheet, and buyers get a Cash App link for their payment.

## Things to fill in (top of `index.html`, in `CONFIG`)

- `cashtag`: your Cash App $cashtag
- `endpoint`: your Google Apps Script web app URL (see below)
- Prices, deposit, deadline and design text are all there too.

## Photos

Put square photos in the `images/` folder with these names:

- `images/dinosaur.jpg`
- `images/mesaverde.jpg`
- `images/robert.jpg`
- `images/rose.jpg`

Until a photo is added, its card says "Photo coming soon."

## Connect the Google Sheet (about 5 minutes)

1. Make a new Google Sheet, e.g. "Mug Orders."
2. In the Sheet: **Extensions → Apps Script**. Delete what's there, paste in `apps-script/Code.gs`, and save.
3. **Deploy → New deployment →** gear icon **→ Web app**. Set *Execute as: Me* and *Who has access: Anyone*. Click Deploy and allow access.
4. Copy the Web app URL (it ends in `/exec`) into `endpoint` in `index.html`.

Each order becomes a row in the **Orders** tab, with its order number, mugs, total, amount due now and amount due at pickup. There are empty "Paid?" and "Picked up?" columns for you.

## Turn on the website

In this repo on GitHub: **Settings → Pages → Branch: main, folder: / (root) → Save.** After a minute the page is live at `https://willpach.github.io/<repo-name>/`.

The form shuts itself off after the deadline in `CONFIG`.
