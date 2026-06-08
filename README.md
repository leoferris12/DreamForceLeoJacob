# Flightway — Career Compass

A single-file career-fit quiz (`flightway.html`) with a tiny Netlify
serverless function that emails users a link to their personalized results.

## How the email flow works

1. The user completes the quiz and reaches the **email-gate** screen.
   They can either go back to the home page or click **Email me my results**.
2. The popup asks for their email address.
3. On submit, the browser:
   - serializes the entire quiz state (answers, sliders, budget, GPA,
     school, resume boosts) to JSON,
   - base64url-encodes it into a compact token,
   - builds a results URL of the form
     `https://your-site.netlify.app/flightway.html#r=<token>`, and
   - `POST`s `{ email, results_url, token }` to
     `/.netlify/functions/send-results`.
4. The Netlify function (`netlify/functions/send-results.js`) calls the
   [Resend](https://resend.com) email API to send the user an email
   containing a **View My Results →** button linking to that URL.
5. When the user clicks the button, `flightway.html` loads with `#r=…`
   in the URL hash. The page detects this on load, decodes the state,
   and jumps straight to the results screen — no quiz to retake, no
   database lookup.

Because the entire state lives in the URL, the backend never needs to
store anything. The only job of the function is to send the email.

## Files

```
flightway.html                   # the whole site (intro, quiz, gate, results)
netlify.toml                      # publish dir + functions dir
netlify/functions/send-results.js # serverless email sender
```

## Deploying to Netlify

### 1. Create a Resend account and an API key

1. Sign up at <https://resend.com> (free tier covers 3,000 emails/month).
2. **Settings → API Keys → Create API Key** with **Send access**.
3. (Recommended) Add and verify your domain under **Domains** so emails
   come from `quiz@yourdomain.com` instead of the shared sandbox sender.
   Without a verified domain you can only send to your own verified
   email address using the default sender.

### 2. Push to Git and connect to Netlify

```bash
cd Flightway
git init && git add . && git commit -m "Initial commit"
# create a repo on GitHub/GitLab, then:
git remote add origin <your-repo-url>
git push -u origin main
```

In Netlify: **Add new site → Import from Git → pick your repo**. The
build settings come from `netlify.toml`, so leave them blank.

### 3. Set environment variables

In Netlify: **Site settings → Environment variables**, add:

| Key              | Value                                                            |
|------------------|------------------------------------------------------------------|
| `RESEND_API_KEY` | The key from step 1 (starts with `re_…`)                         |
| `FROM_EMAIL`     | `Flightway <quiz@yourdomain.com>` (or `onboarding@resend.dev`)  |
| `ALLOWED_ORIGIN` | (optional) `https://your-site.netlify.app` — defaults to `*`     |

Then **Deploys → Trigger deploy → Clear cache and deploy site** so the
function picks up the new vars.

### 4. Test it

1. Open `https://your-site.netlify.app/flightway.html`.
2. Finish the quiz, click **Email me my results**, enter your email.
3. You should receive an email within a few seconds with a **View My
   Results →** button. Click it and the results page should render
   immediately, populated with your real answers.

## Local development

```bash
npm install -g netlify-cli
cd Flightway
netlify dev
```

`netlify dev` serves `flightway.html` at <http://localhost:8888> and
proxies `/.netlify/functions/send-results` to the local function.
Create a `.env` file (or use `netlify env:set`) with `RESEND_API_KEY`
and `FROM_EMAIL` so the local function can actually send mail.

## Notes & limitations

- **No database.** Results are encoded in the URL hash, so the email
  link is the only place they're stored. If a user loses the email, the
  results are gone.
- **URL length.** The token is typically 1–3 KB. Email clients and
  browsers handle this fine, but be aware if you significantly expand
  the quiz.
- **Rate limiting.** The function does not rate-limit by itself. If you
  expect abuse, add Netlify's [edge rate limiting](https://docs.netlify.com/edge-functions/api/#rate-limit)
  or check the request IP and reject repeated submissions.
- **Privacy.** Emails are sent to Resend and are not stored by this
  app. Add a short privacy notice to the quiz if you intend to collect
  emails at scale.
