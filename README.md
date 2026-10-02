# Fredagsdrinken

The office's Friday drink site: this week's drink and ingredients, voting from 1 to 10, a public top list, a committee admin page and a TV mode with a QR code.

It runs on **Cloudflare Pages** (free plan): static pages in `public/`, small server functions in `functions/`, and a **D1** database for the schedule and votes. Deploying works like GitHub Pages: push to GitHub and Cloudflare publishes it.

## Pages

| URL | Who | What |
|---|---|---|
| `/` | Everyone | Drink of the week, voting, top list |
| `/admin.html` | Drinkkommittén | Schedule, Excel import, who voted what, TV mode |
| `/tv.html` | Meeting TV | Drink, QR code "Rösta på veckans drink här!", live top list |

## How voting works

1. Scan the QR code or open the site.
2. Type the office password once. The phone remembers it.
3. Tap a score from 1 to 10, pick your name and send.

Voting again for the same week replaces your earlier vote. Voting for a week opens on that Friday at 15:40 Stockholm time, enforced by the server. Earlier weeks stay open.

The public site only shows weeks up to the current one, so upcoming drinks stay a surprise. It shows averages and vote counts, never who voted what. The committee sees names on the admin page and can remove a vote.

## Deploy (about 15 minutes, once)

You need a free Cloudflare account and a GitHub account.

1. **Put the code on GitHub.** Create a new repository and push this folder to it.
2. **Create the database.** In the Cloudflare dashboard go to *Storage & Databases → D1 SQL Database → Create*, and name it `fredagsdrinken`.
3. **Create the tables.** Open the new database, go to the *Console* tab, paste everything from `schema.sql` and run it.
4. **Connect the database to the code.** Copy the database ID from the database overview page. In `wrangler.toml`, replace `REPLACE_WITH_YOUR_DATABASE_ID` with it, then commit and push.
5. **Create the site.** Go to *Workers & Pages → Create → Pages → Connect to Git* and pick the repository. Leave the build command empty. The output directory is read from `wrangler.toml` (`public`).
6. **Set the two passwords.** In the Pages project go to *Settings → Variables and Secrets* and add these as **Secret** values for Production:
   - `OFFICE_PASSWORD`, the password everyone in the office types to vote
   - `ADMIN_PASSWORD`, the committee password for `/admin.html`
7. **Redeploy** so the passwords take effect (*Deployments → ⋯ → Retry deployment*).

The site is now live at `https://<project-name>.pages.dev`. You can add your own domain under *Custom domains*.

To change a password later, edit the secret and redeploy. Phones with the old office password will ask for the new one automatically.

## Common changes

- **Who can vote.** Edit `NAMES` in `lib/shared.js`, commit and push.
- **When voting opens.** Edit `OPEN_H` and `OPEN_M` in `lib/shared.js`.
- **The rotating images.** Replace `public/img/drink-0.jpg` to `drink-9.jpg`. Week number modulo 10 picks the image.
- **The top image.** Replace `public/img/hero.jpg`.

## Run it locally (optional)

```sh
npm install
cp .dev.vars.example .dev.vars   # then set your own test passwords
npm run db:local                 # create the tables in a local database
npm run dev                      # open http://localhost:8788
```

## Files

```
public/            static pages, styles, images, QR library
functions/api/     server endpoints
  state.js         GET  public data (schedule up to this week, vote totals)
  check.js         POST check the office password
  vote.js          POST save a vote
  admin/           committee endpoints, protected by ADMIN_PASSWORD
lib/shared.js      names list, voting time, date helpers
schema.sql         database tables
wrangler.toml      Cloudflare config
```

## Good to know

- The office password is shared, so anyone who has it can pick any name. The committee page shows every vote and lets you remove odd ones.
- Each name gets one vote per week. Voting again overwrites it.
- On the free plan, Cloudflare D1 and Pages allow far more traffic than an office of 20 will ever use.
