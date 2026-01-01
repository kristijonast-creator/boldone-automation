# CommerceCore Development Context

## Overview
You are helping Chris maintain and improve CommerceCore's Google Apps Script automation systems and the BoldOne Creative Hub web application. There are 5 interconnected systems that share a central configuration sheet.

## Active Systems

| System | Purpose | Location |
|--------|---------|----------|
| **Ad Naming Sheet** | Batch creation, folder automation, ad naming conventions | Standalone Google Sheet |
| **Atria Reports** | Performance reporting, Slack notifications | Standalone Google Sheet |
| **Sprint → ClickUp** | Auto task creation from sprint sheets | Standalone Google Sheet |
| **Drive Naming QA** | File naming validation | Standalone Google Sheet |
| **BoldOne Creative Hub** | Central config web app, team/product management | Google Apps Script Web App |

## Critical Configuration

### Config Sheet ID (MEMORIZE THIS)
```
1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk
```

### ClickUp Config - MISLEADING NAMING (CRITICAL)
The ClickUp Config tab has naming that has caused major debugging issues:

| Config Key | Value | What It Actually Is |
|------------|-------|---------------------|
| `FB_ADS_FOLDER_ID` | 90150821271 | **The actual ClickUp Space ID for API calls** |
| `SPACE_ID` | 90151192866 | A folder within that space (NOT the space!) |
| `SPRINT_FOLDER_ID` | 901512650311 | PC SPRINT '26 folder |
| `LEGACY_SPRINT_FOLDER_ID` | 90154797247 | PC SPRINT 2025 folder |

**The `getClickUpSpaceId_()` function must return `config.FB_ADS_FOLDER_ID`, NOT `config.SPACE_ID`**

### Caching Architecture
Two levels of caching exist that MUST both be cleared when debugging:

1. **Script cache** (`CacheService.getScriptCache()`) - 300 second TTL
   - Keys: `atria_clickup`, `atria_folders`, `atria_cpa`, `atria_slack`, `atria_product_sheets`

2. **In-memory lazy loaders** (persist within script execution)
   - `_clickupConfigData`
   - `_productLocaleFolders`
   - `_targetCpa`

**When debugging config issues, ALWAYS clear BOTH:**
```javascript
CacheService.getScriptCache().remove('atria_clickup');
_clickupConfigData = null;
```

## Code Style Requirements

### JavaScript/Google Apps Script
- Use `function name() {}` syntax (NOT arrow functions for top-level)
- Use descriptive Logger.log messages with emojis for visibility
- Add comments for complex logic
- Keep functions focused and single-purpose
- Always wrap external API calls in try-catch
- ClickUp API IDs are strings - always use loose equality (`==` not `===`)

### Function Naming
- Public functions: `functionName()` - callable from client-side
- Private functions: `functionName_()` - underscore suffix makes them private, NOT callable from google.script.run

### HTML/CSS
- Use CSS custom properties (variables) for theming
- Support both light and dark modes
- Glass morphism styling with backdrop-filter
- Mobile-first responsive design

## BoldOne Creative Hub Architecture

### Backend (code.gs)
- `getDashboardData()` - Main data fetch for dashboard
- `getDashboardDataFresh()` - Clears cache then fetches (for after add/edit/delete)
- `getProducts_()` - Fetches products with footage folders
- `getTeamMembers_()` - Fetches team members
- All edit functions clear relevant caches

### Frontend (index.html)
- Single-page app with panel switching
- Left sidebar navigation (implemented 2024-12-31)
- Optimistic UI for add/delete operations
- Toast notifications for feedback
- Real-time validation in modals

### Config Sheet Tabs
- **Products** - Product keys, IDs, CPA targets, status
- **Team Members** - Users, roles, permissions
- **Product Folders** - Drive folder IDs per locale
- **Footage Folders** - Raw footage URLs per product
- **ClickUp Config** - API keys, space/folder IDs
- **Valid Values** - Dropdowns for platforms, locales, etc.
- **ClickUp Log** - Task creation history
- **Error Log** - System errors

## Debugging Approach (LEARNED FROM EXPERIENCE)

1. **Compare working code vs current code FIRST** - Chris often has working backup versions
2. **Test API endpoints directly with different parameter values** - Don't trust error messages
3. **Don't trust error messages literally** - "Oauth token not found" was actually wrong Space ID
4. **Check Config sheet values against what code expects** - Naming can be misleading
5. **One fix at a time** - Test after each change, don't stack multiple "fixes"
6. **When API returns 401** - Check if the ID parameters are correct BEFORE assuming auth issues
7. **When stuck for more than 2-3 attempts** - Step back and compare with working code

## Making Changes - Required Format

When providing code changes, ALWAYS use this format:

```
**FIND:**
[exact code to find]

**REPLACE WITH:**
[exact replacement code]
```

Also specify:
- Which file the change applies to
- Line number or function name if helpful
- What the change does and why

## Important Reminders

- Chris prefers exact code, not explanations of what to do
- Always mention which file/system a change applies to
- Suggest running diagnostic functions after changes
- Warn about potential issues before they happen
- When creating files, always put them in /mnt/user-data/outputs/
- Use present_files tool to share completed files

## Current UI/UX Patterns

### Modals
- Real-time validation with visual feedback
- Progress indicators for required fields
- Duplicate checking against existing data
- Success animations (confetti, banners, highlighting)

### Tables
- Sortable columns
- Optimistic UI (syncing/deleting badges)
- Mobile card view fallback
- Action buttons (edit, delete)

### Toasts
- Success (green), Error (red), Warning (yellow), Info (blue)
- Persistent toasts for long operations (syncing, deleting)
- Auto-dismiss for quick notifications

### Navigation
- Left sidebar with sections (Main, Insights, System)
- User profile and theme toggle at bottom
- Mobile: slide-out with overlay