# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MMObjectView ("Salesforce Object Explorer") is a Google Apps Script (GAS) add-on for Google Sheets that displays Salesforce object metadata (field definitions) in spreadsheet format. It authenticates via OAuth2 to Salesforce, fetches object/field metadata through the REST API, and generates formatted sheets with 24-column field detail tables.

The project is written in Japanese (comments, UI text, error messages).

## Deployment Commands

This project uses [clasp](https://github.com/nicholasq/clasp) for GAS deployment. There is no npm/build/test/lint toolchain.

```bash
# Push local code to Google Apps Script
clasp push

# Pull remote changes from GAS editor
clasp pull

# Open the GAS editor in browser
clasp open

# View logs (Stackdriver)
clasp logs
```

All clasp commands must be run from the `src/` directory (where `.clasp.json` lives).

## Architecture

All `.gs` files share a single global scope (GAS constraint). Private functions use a `_` suffix convention to avoid name collisions.

### Module Responsibilities

- **Code.gs** — Facade pattern entry point. All `google.script.run` calls from the sidebar route here. Delegates to other modules and returns unified `{ success, data, error }` responses.
- **Auth.gs** — OAuth2 flow using the `apps-script-oauth2` library (v43). Manages token storage via `PropertiesService` + `CacheService`, with `LockService` for concurrent refresh safety.
- **Config.gs** — All constants, Script Property keys, error codes, Japanese message mappings, field type translation table, and sheet formatting constants. Configurable values (client ID, API version, login URL) are read from Script Properties via `getXxx_()` helpers.
- **SalesforceApi.gs** — Salesforce REST API integration. Includes automatic 401 token refresh + retry, exponential backoff for 500/503, and 6-hour caching of object lists via `CacheService`.
- **SheetBuilder.gs** — Creates and formats new sheets with field data: header row, batch data writes, column widths, auto-filter, zebra striping, and summary metadata.
- **sidebar.html** — Single-file frontend with Material Design styling, embedded CSS/JS. Handles auth status display, environment switching (production/sandbox), object search/filter, and object expansion triggers.

### Data Flow

```
sidebar.html  →  google.script.run  →  Code.gs (facade)
                                           ├→ Auth.gs (OAuth2)
                                           ├→ SalesforceApi.gs (REST API calls)
                                           └→ SheetBuilder.gs (sheet generation)
```

### Key GAS Services Used

- `SpreadsheetApp` — Sheet creation/formatting
- `PropertiesService` (User scope) — OAuth token + config storage
- `CacheService` (User scope) — Object list caching (6h TTL)
- `LockService` — Token refresh concurrency control
- `UrlFetchApp` — HTTP requests to Salesforce
- `HtmlService` — Sidebar UI rendering

### External Dependency

- `apps-script-oauth2` library v43 (ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`), referenced in `appsscript.json`

## Configuration

Salesforce credentials are stored in GAS Script Properties (not in source code):
- `SF_CLIENT_ID` — Connected App Consumer Key (required)
- `SF_CLIENT_SECRET` — Connected App Consumer Secret (required)
- `SF_LOGIN_URL` — defaults to `https://login.salesforce.com` (optional)
- `SF_API_VERSION` — defaults to `v62.0` (optional)

## Conventions

- Facade functions in `Code.gs` are the only public API surface for `google.script.run`
- Private/internal functions end with `_` (e.g., `callSalesforceApi_()`, `getLoginUrl_()`)
- All public facade functions return `{ success: true, data: ... }` or `{ success: false, error: { code, message } }`
- Error codes are defined in `Config.gs` under `ERROR_CODES`
- Field column definitions (24 columns) are defined in `Config.gs` under `FIELD_COLUMNS`
