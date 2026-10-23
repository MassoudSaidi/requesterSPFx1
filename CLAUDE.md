# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

SharePoint Framework (SPFx 1.16.1) client-side web part, "BSUP Document Portal" (CAMP – Customer Asset Management Portal). Users search BlackStone documents, request access to them, and view their own requests. The README is the unfilled SPFx template; ignore it.

## Commands

Requires Node >=16.13 <17 (SPFx 1.16 constraint). `gulp` is a local devDependency; there is no `gulp-cli` script, so use `npx gulp ...`.

- `npx gulp serve` — local dev server (https, port 4321). Opens the hosted workbench at the tenant URL in `config/serve.json`, so it needs a signed-in SharePoint tenant; SharePoint data calls will not work from a plain localhost page.
- `npm run build` (`gulp bundle`) — dev bundle; `npx gulp bundle --ship && npx gulp package-solution --ship` — production `.sppkg` at `sharepoint/solution/BSUP-DocPortal-v1.sppkg` (path from `config/package-solution.json`). Bump `solution.version` there when deploying.
- `npm run clean` — `gulp clean`.
- `npm test` runs `gulp test`, but there are no tests or test config in the repo.
- Lint: ESLint config is `.eslintrc.js` (SPFx rules); it runs as part of the gulp build.

## Architecture

One web part, `src/webparts/bsupDocGen1/`:

- `BsupDocGen1WebPart.ts` — thin host. Passes `context` (WebPartContext), admin/manager emails from the property pane, and theme/Teams info into the React component.
- `components/BsupDocGen1.tsx` (~1.5k lines) — essentially the entire app in a single class component with a large `MyState`. UI is a mix of PrimeReact (DataTable, TabView, Dialog) and Fluent UI (ComboBox, TextField, buttons). Do not expect to find separated services, hooks, or stores.
- `components/CamlQuery.ts` — CAML query templates (`qry_documentSearch`, `qry_myRequests`) as Handlebars strings. The component compiles them with `Handlebars.compile(HTMLDecoder.decode(...))`, filling `{{minid}}/{{maxid}}` and per-field `{{Contains_XSearchKey}}` operator tags (Eq/Contains) plus `{{XSearchKey}}` values. `CamlQuery copy.ts` is a stale backup that differs from `CamlQuery.ts`; edit the non-copy file.

### Data layer (PnPjs v3, no custom backend)

All data is SharePoint lists, accessed with `@pnp/sp`:

- `sp = spfi().using(SPFx(context))` is the current site; `spSite = spfi(tenantUri + "/sites/" + targetDocumentSite).using(AssignFrom(sp.web))` targets the separate admin site (`targetDocumentSite`, currently `BlackStonePortalAdmin2`) where all lists live. Site, list and library names are hardcoded `readonly` fields at the top of the component (~line 255), not configuration. Several hardcoded tenant URLs also appear in `config/serve.json` and in comments/links in the component.
- Lookup lists (`sys_*`: Regions, Clients, Document_Classes, Document_Types, EquipmentTypes, Manufacturers, Models, Assets, Sites, Serials, Manager_Departments) populate the dropdowns in `loadDropDowns()`. The same fetch/sort/push-`TBD`-and-blank pattern is repeated per list.
- Document metadata is in the `BlackStone Document Data` list. Searches use `getItemsByCAMLQuery` and are paged by ID range (`_searchLargeList` / `_searchLargeRequestList`, `pageSize` 2000) to get around the list view threshold. Preserve this ID-window approach when changing search.
- Writes: access requests go to `BlackStone Document Requests`; a row is also added to `BlackStone WBP Requests` (the trigger list for a Power Automate sharing flow, so changing its fields can break that external flow); document views are logged to `log_Document_Views`. Notification emails go through `sp.utility.sendEmail`, addressed using the department→manager mapping and the admin/manager emails from the property pane.
- Document `Document Class`/`Document Type` carry an approval class (A/B/C) with time frames, which drives the request progress bars.

## Notes

- Internal SharePoint column names appear encoded in state keys (e.g. `DocumentClass_x002d_Discpline`), so match the existing internal names when adding fields.
- Strings live in `loc/en-us.js` + `loc/mystrings.d.ts`; the property pane declares `classA/B/CtimeFrameWbProb` props that are not yet exposed in the pane or passed to the component.
