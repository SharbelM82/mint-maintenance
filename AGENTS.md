# Mint Response: preserve the original website

The source of truth is the user's `remix-mint-response-llc.zip`, with ONLY the changes in `reference/2026-09-06 Requested Website Changes.docx` applied. This is NOT the previously redesigned/modified v3.x website.

Do not change customer-facing language, typography, colors, spacing, layout classes, logos, section order, page structure, or service claims unless the owner expressly instructs you to do so. Do not rewrite awkward grammar or reconcile inconsistent remaining references by assumption. Do not import any of the earlier modified site's design, positioning, pages, photographs, or copy.

Read `docs/CHANGE_LOG.md` and `docs/UNCHANGED_ITEMS.md` before editing. The changes were scoped to the screenshots in the instruction document. Additional Kansas/renovation references outside those marked blocks were intentionally preserved. The footer instruction was interpreted literally as removing its final sentence only. Obtain explicit direction for broader changes.

Backend bug fixes are allowed; keep the normal appearance of the form unchanged. Never fabricate submission success. The existing confirmation is displayed only after the server's Formspree adapter receives explicit acceptance. Do not replace this with a timer, localStorage, or a fake ticket. Do not add new public claims of performance, licensing, response times, or coverage.

Keep credentials server-side. CONTACT_FORM_ENDPOINT belongs in local .env.local or the Cloudflare Pages environment, never a VITE_ variable or committed credential. Do not send live test inquiries without owner approval. Public preview files must never submit requests.

Before proposing publication: run npm install once to generate a real package-lock.json, then npm run verify. Commit the resulting lockfile after review. Future installs should use npm ci. No production Vite build was possible in the preparation environment because package downloads were unavailable; do not treat the HTML review files as a verified Vite build.

Use a draft Git branch and show the diff before merging to main. Do not change DNS, publish, send invitations, or connect accounts without the owner's authorization. Do not automatically update protected-file hashes just to make a check pass; a change requires explicit approval and an updated change log.
