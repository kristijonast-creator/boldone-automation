# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

BoldOne Automation Suite is a collection of Google Apps Script systems for managing ad creative workflows. All systems are deployed as Google Apps Script projects attached to Google Sheets (except the web app).

## Development Commands

This is a Google Apps Script project. Use [clasp](https://github.com/google/clasp) for local development:

```bash
# Install clasp globally
npm install -g @google/clasp

# Login to Google
clasp login

# Clone a specific Apps Script project
clasp clone SCRIPT_ID --rootDir ./boldone-hub

# Push local changes to Apps Script
clasp push

# Pull remote changes
clasp pull

# Open the Apps Script editor in browser
clasp open
```

## Architecture

### Central Configuration

All 5 systems share a single Google Sheet for configuration:
- **Config Sheet ID**: `1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk`
- **Cache Duration**: 300 seconds (5 minutes) across all systems

Config tabs: Products, Team Members, Product Folders, Footage Folders, ClickUp Config, Valid Values, ClickUp Log, Error Log

### Systems

| System | Location | Purpose |
|--------|----------|---------|
| BoldOne Hub | `boldone-hub/` | Central web dashboard (code.gs + index.html SPA) |
| Ad Naming | `ad-naming/` | Batch creation, folder automation |
| Sprint→ClickUp | `sprint-clickup/` | Auto task creation from sprint sheets |
| Atria Reports | `atria-reports/` | Performance reporting, Slack notifications |
| Drive QA | `drive-qa/` | File naming validation |

### Two-Level Caching Pattern

All systems use this caching strategy—**both levels must be cleared when debugging**:

```javascript
let _inMemoryCache = null;  // Level 1: within execution

function getData_() {
  if (_inMemoryCache) return _inMemoryCache;

  const cache = CacheService.getScriptCache();  // Level 2: across executions
  const cached = cache.get('key');
  if (cached) {
    _inMemoryCache = JSON.parse(cached);
    return _inMemoryCache;
  }
  // ... fetch from source
}

function clearCache_() {
  CacheService.getScriptCache().remove('key');
  _inMemoryCache = null;  // Clear BOTH!
}
```

## Critical Conventions

### Function Naming
- `functionName()` — Public, callable from client-side `google.script.run`
- `functionName_()` — Private (underscore suffix), NOT callable from client-side

If `google.script.run.someFunction_()` silently fails, check for the underscore.

### ClickUp Configuration (Misleading Names)
The ClickUp Config tab has confusing naming that has caused debugging issues:

| Config Key | What It Actually Is |
|------------|---------------------|
| `FB_ADS_FOLDER_ID` (90150821271) | **The actual Space ID for API calls** |
| `SPACE_ID` (90151192866) | A folder within that space (NOT the space!) |

`getClickUpSpaceId_()` must return `config.FB_ADS_FOLDER_ID`, not `config.SPACE_ID`.

### ClickUp API
- IDs are strings—use `==` not `===` for comparisons
- Authorization header: just the API key, no "Bearer " prefix
- 401 errors often mean wrong Space/Folder IDs, not auth issues

### Sheet Operations
`insertRowAfter()` copies formatting from the row above. Always clear it:

```javascript
sheet.insertRowAfter(1);
const newRow = sheet.getRange(2, 1, 1, cols);
newRow.setValues([data]);
newRow.setBackground(null).setFontWeight('normal');  // Clear copied formatting
```

### Web App Client-Server Communication
```javascript
// Server returns success/error object
function serverFunction(data) {
  try {
    return { success: true, data: result };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
```

## BoldOne Hub Frontend

- Single-page app with left sidebar navigation
- CSS variables for theming (light/dark mode with `[data-theme="dark"]`)
- Glass morphism styling with `backdrop-filter`
- Optimistic UI: immediate visual feedback → server call → confirm/rollback

## Debugging Approach

1. Compare working code vs current code first (backups often exist)
2. Test API endpoints with different parameter values—don't trust error messages
3. Check Config sheet values against what code expects (naming is misleading)
4. Clear BOTH cache levels when debugging config issues
5. One fix at a time—test after each change
