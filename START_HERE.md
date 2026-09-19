# Start here: Cursor → GitHub → Cloudflare Pages

## 1. Open the correct folder

Extract `Mint_Response_ORIGINAL_Cursor_GitHub.zip`. In Cursor, use File → Open Folder and select the folder containing `package.json`, `src/`, `worker/` and this guide. Do not start with the earlier v3.x modified-site package, and do not upload the ZIP itself as your repository's source.

Before asking Cursor to edit anything, give it this prompt:

> Read AGENTS.md, docs/CHANGE_LOG.md and docs/UNCHANGED_ITEMS.md. This is the ORIGINAL Mint Response website with the September 6 requested edits, not the earlier redesigned site. Do not change any wording or formatting without my explicit instruction. First install the dependencies, run npm run verify and show me the result. Fix technical build/backend failures without altering public copy or design. Do not deploy or modify DNS.

## 2. Install and review locally

Install Node.js 22 if needed. In Cursor's terminal, inside this project folder:

```sh
npm install
npm run verify
npm run dev
```

Open `http://127.0.0.1:3000` in Chrome or Edge. Check the original design and the requested removals/replacements. The page remains a single-page React website with its original section structure.

This handoff has no invented lockfile. The first `npm install` creates `package-lock.json`; review and commit that real file. Use `npm ci` after it exists. The preparation environment could not download npm packages, so a production Vite build, dependency-resolved type check and full visual review still need to run here.

## 3. Put the source in your GitHub account

Use a private repository under your owner account. Cursor's Source Control panel or GitHub Desktop can initialize the folder, create a first commit and publish the repository. Include the actual `package-lock.json` generated above. Exclude node_modules, real .env files and credentials; `.gitignore` is supplied.

If using a terminal, create an empty GitHub repository first, then use the actual repository URL returned by GitHub:

```sh
git init
git add .
git commit -m "Original Mint Response site with requested September 6 edits"
git branch -M main
git remote add origin YOUR_ACTUAL_GITHUB_REPOSITORY_URL
git push -u origin main
```

Replace YOUR_ACTUAL_GITHUB_REPOSITORY_URL with your repository's actual URL; do not paste the placeholder literally. The supplied GitHub workflow checks the project. It does not itself deploy the website. Configure repository permissions and any required approvals in your own account; those settings were not changed here.

For future edits, use a separate branch and review the diff before merging. Keep the owner in control of repository, domain, hosting, billing, form account and recovery methods.

## 4. Connect Cloudflare Pages only after approval

Create a Pages project with Git integration and select only the intended repository. Do not accidentally create a Workers-only project using different settings.

| Setting | Value |
|---|---|
| Production branch | main |
| Framework preset | None |
| Root directory | Repository root |
| Build command | npm run verify |
| Build output directory | dist |
| Node version | 22 |

The build produces the original React site plus `_worker.js` and `_routes.json` for `/api/service-request`. This is a Cloudflare Pages advanced-mode backend. A plain GitHub Pages deployment will not execute it.

A connected production-branch push/merge can deploy automatically. Use preview branches for changes and approve them before merging to main. No repository or hosting account was created or connected here.

Do not attach the live domain until the build, content, form and desktop/mobile review are approved. When connecting the website, preserve Google Workspace's MX, SPF, DKIM, DMARC and verification records; do not delete mail records as part of website setup. Domain ownership and existing DNS were not verified in this handoff.

## 5. Activate the form separately

Read `docs/FORM_SETUP.md`. Configure a real owner-controlled Formspree endpoint as the server variable `CONTACT_FORM_ENDPOINT` and run an owner-approved end-to-end test. No API key or provider ID is prefilled. Until configured, an attempted submission fails honestly, keeping the entered details.

Before launch, review `docs/UNCHANGED_ITEMS.md`: the original phone placeholder and some unmarked Kansas/renovation wording remain. No new copy was invented to fix those items without your instruction.

## Preview files

The standalone and before/after HTML files run the actual original/edited TSX compiled for review. They load React 19.0.1, Lucide 0.546.0 and Tailwind browser 4.1.14 from public CDNs, plus the original Google Fonts stylesheet. They need internet access and a modern browser. They contain no working submission path, and they are not production Vite builds.

The comparison controls are outside the website frame; they are not additions to your site. If the preview cannot load dependencies, use the Cursor preview above instead. These CDN-dependent previews could not be visually verified in the preparation environment.

## Official technical references

Cloudflare Pages Git integration: https://developers.cloudflare.com/pages/get-started/git-integration/

Cloudflare Pages advanced mode: https://developers.cloudflare.com/pages/functions/advanced-mode/

Cloudflare Pages routing: https://developers.cloudflare.com/pages/functions/routing/

Cursor project rules: https://prod.cursor.com/docs/rules
