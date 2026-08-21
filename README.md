# Production Document Center

A real, hosted version of your document/template/submission tracker, built on
[Supabase](https://supabase.com) (database, auth, file storage) and deployed on
[Vercel](https://vercel.com).

Everything below uses free tiers. Total setup time: ~20–30 minutes.

---

## 1. Supabase — already set up ✅

Your Supabase project (`felixt115@gmail.com's Project`) is live and already configured:
- Tables (`profiles`, `documents`, `templates`, `submissions`) and their security rules are in place.
- The `attachments` storage bucket exists and is properly locked down.
- `.env` in this folder is already filled in with your real Project URL and public API key —
  nothing to copy-paste here.

One optional setting worth knowing about: **Authentication → Providers → Email → Confirm email**
is likely ON by default, meaning new sign-ups need to click a confirmation link before they can log
in. That's good practice for a real deployment. If you want to test faster right now, you can
temporarily switch it off in the Supabase dashboard, then turn it back on later.

---

## 2. Run it locally to confirm it works

You'll need [Node.js](https://nodejs.org) installed (any recent version).

```bash
cd production-doc-center
npm install
```

```bash
npm run dev
```

Open the URL it prints (usually `http://localhost:5173`). Sign up with your name, email, and a
password — **you'll automatically become Admin** since you're the first person to sign up.
From there, everything works the same as before: build documents, build templates, fill them out,
and manage other users' access from the Users tab once they've signed up too.

---

## 3. Put the code on GitHub

I don't have a GitHub connection available to push this for you, so this part is manual —
should only take a couple minutes:

1. Create a new repository at https://github.com/new (private is fine).
2. In this project folder:

```bash
git init
git add .
git commit -m "Production Document Center"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

(`.env` is already excluded via `.gitignore` — your Supabase keys won't be pushed. The anon key is
safe to expose in the deployed app itself since it's protected by the security rules already set
up in the database, but there's no reason to commit it to the repo.)

---

## 4. Deploy on Vercel

Also manual — no Vercel connection available here either.

1. Go to https://vercel.com, sign up (use "Continue with GitHub" — makes this step easier), and click **Add New → Project**.
2. Import the repository you just pushed.
3. Vercel will auto-detect Vite. Before deploying, expand **Environment Variables** and add:
   - `VITE_SUPABASE_URL` → `https://fztfapmpopnxdxsdlhie.supabase.co`
   - `VITE_SUPABASE_ANON_KEY` → `sb_publishable_ha2UlpEKDTWREHn5vPofRQ_hSe7M0Ng`
4. Click **Deploy**. In about a minute you'll get a live URL like `production-doc-center.vercel.app`.

That URL is your real, standalone website — share it with your team.

---

## Day-to-day usage

- **Sign-up is self-serve.** Send your team the site URL; each person creates their own account
  (name, email, password).
- **You grant access.** New accounts start as a Member with no build permissions. Go to the
  **Users** tab to turn on Doc Builder, Template Builder, or promote someone to Admin.
- **Deactivating someone** locks them out without deleting their account — a browser-based app
  can't safely delete login credentials outright, so this is the way to cut off access.
- **Forgot password** works from the login screen and sends a real reset email through Supabase.

## If something needs adjusting later

- **Storage limits**: the free Supabase tier includes 1GB of file storage and 500MB of database
  space — plenty for this kind of tool for a good while. If you outgrow it, Supabase's paid tier
  is pay-as-you-go from there.
- **Custom domain**: Vercel supports adding your own domain (e.g. `docs.thermobond.com`) for free
  under Project Settings → Domains.
- **Re-running the schema**: if you ever need to reset the database, you can re-run
  `supabase-schema.sql` — but note it will fail on tables that already exist rather than wiping
  them, which is a safety feature, not a bug.
