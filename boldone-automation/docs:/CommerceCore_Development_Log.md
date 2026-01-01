# CommerceCore & BoldOne Development Log

## Overview
This document tracks all development work, changes, and known issues across the CommerceCore automation systems and BoldOne Creative Hub.

---

## 📅 2024-12-31 (Current Session)

### BoldOne Creative Hub - Major Updates

#### ✅ Left Sidebar Navigation (Option B)
- Replaced top navigation bar with left sidebar
- Sections: Main (Dashboard, Team, Products), Insights (Analytics, Goals, Activity), System (Settings)
- User profile and theme toggle at bottom
- Mobile: Slide-out sidebar with overlay
- Glass morphism styling maintained

#### ✅ Optimistic UI for Add/Delete Operations
- **Add operations**: Modal closes instantly → Optimistic row appears with blue "Syncing..." badge → Server confirms → Fresh data fetched → Confetti + success banner
- **Delete operations**: Red "Deleting..." badge → Server confirms → Row removed → Success toast
- State persistence across re-renders using JavaScript Sets
- `isDeleting` flag pauses auto-refresh during delete operations

#### ✅ Footage Folder Integration
- Added Footage Folder URL field to Add/Edit Product modals
- Quick-access button to open folder in new tab
- Reads from/writes to "Footage Folders" tab in Config sheet
- Icon: Video camera

#### ✅ Ad Locale Folders Rename
- Renamed "Google Drive Folders" → "Ad Locale Folders" for clarity

#### ✅ Beta Badges
- Added purple gradient "Beta" badges to Analytics and Goals nav items

#### ✅ Table Column Alignment Fix
- Fixed column alignment issues in Team and Products tables
- Added horizontal scroll for narrow screens (min-width: 800px)

#### ✅ Row Formatting Bug Fix (Sprint → ClickUp)
- Fixed `insertRowAfter()` copying header formatting to new rows
- Added `.setBackground(null).setFontWeight('normal')` after inserts
- Applied to: `logClickUpTasksCreated()`, `logErrorToHub_()`

#### ✅ Sprint → ClickUp Task Logging Fix
- Renamed `logClickUpTasksCreated_()` to `logClickUpTasksCreated()` (removed underscore)
- Private functions can't be called from client-side `google.script.run`
- Historical baseline of 250 tasks added to ClickUp Log

---

## 📅 2024-12-31 (Earlier Sessions)

### Session 3: Modal Overhaul
- Real-time validation for Add Product / Add Team Member modals
- Duplicate checking against existing data
- Progress indicators for required fields
- Success animations (confetti, banners, auto-scroll highlighting)
- Field hints and warnings

### Session 2: Tier Naming & Drive Folders
- Implemented Option H tier naming (Titans, Unicorns, Champions, Winners, Rising Stars)
- Added Google Drive folder management in Product modals
- Backend integration for folder CRUD operations

### Session 1: Health Checks & Foundation
- Implemented 6 health check fixes
- Activity logging system
- Copy-to-clipboard functionality
- Bold.One SVG logo integration
- Meta API discussion
- Tier naming options exploration

---

## 🐛 Known Issues

### Resolved
- [x] Field name mismatch in team member forms (clickupKey→apiKey, admin→isAdmin)
- [x] Duplicate `editTeam` case in saveModal (was missing `addTeam`)
- [x] Slow data refresh (12-15s → 6-8s with `getDashboardDataFresh()`)
- [x] Confetti appearing before data loaded
- [x] Deleting badge disappearing prematurely
- [x] Private function not callable from client-side
- [x] Row formatting copying from headers

### Open
- [ ] None currently tracked

---

## 📁 File Structure

### BoldOne Creative Hub
```
BoldOne Config Sheet (Google Sheet)
├── code.gs           # Backend - all server functions
├── index.html        # Frontend - full SPA with CSS/JS
└── Tabs:
    ├── Products
    ├── Team Members
    ├── Product Folders
    ├── Footage Folders
    ├── ClickUp Config
    ├── Valid Values
    ├── ClickUp Log
    ├── Error Log
    └── Analytics (external reference)
```

### Other Systems
```
Ad Naming Sheet
├── code.gs
├── autobatchid.gs
└── Various HTML modals

Atria Reports
├── AtriaSlackAutomation.gs
└── Modal HTMLs

Sprint → ClickUp
├── autobatchcreator.gs (with embedded HTML)

Drive QA
├── drivenamingchecker.gs
```

---

## 🔧 Configuration Reference

### Central Config Sheet
**ID**: `1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk`

### Cache Keys
| Key | Duration | Used By |
|-----|----------|---------|
| `hub_team_members` | 300s | BoldOne |
| `hub_products` | 300s | BoldOne |
| `hub_batches_count` | 300s | BoldOne |
| `hub_clickup_count` | 300s | BoldOne |
| `atria_clickup` | 300s | Atria Reports |
| `atria_folders` | 300s | Atria Reports |
| `atria_cpa` | 300s | Atria Reports |
| `dnc_products` | 300s | Drive QA |

### ClickUp Configuration
| Config Key | Value | Actually Is |
|------------|-------|-------------|
| `FB_ADS_FOLDER_ID` | 90150821271 | **Space ID** (use this!) |
| `SPACE_ID` | 90151192866 | A folder (misleading name) |
| `SPRINT_FOLDER_ID` | 901512650311 | PC SPRINT '26 folder |

---

## 📝 Code Conventions

### Function Naming
- `publicFunction()` - Callable from client
- `privateFunction_()` - Server-only (underscore suffix)

### Logging
```javascript
Logger.log('🚀 Starting operation');
Logger.log('✅ Success: ' + result);
Logger.log('❌ Error: ' + error);
Logger.log('⚠️ Warning: ' + message);
Logger.log('📋 Info: ' + details);
```

### Error Handling
```javascript
try {
  // risky operation
  return { success: true, data: result };
} catch (e) {
  Logger.log('❌ Error: ' + e.toString());
  return { success: false, error: e.toString() };
}
```

---

## 🚀 Planned Features

### BoldOne Creative Hub
- [ ] Analytics dashboard improvements
- [ ] Goals tracking system
- [ ] Bulk operations (multi-select delete/edit)
- [ ] Export functionality
- [ ] Keyboard shortcuts

### Automations
- [ ] Enhanced error notifications
- [ ] Performance optimizations
- [ ] Additional health checks

---

## 📚 Resources

### Project Files Location
All current code files are stored in Claude Project at `/mnt/project/`

### How to Update This Log
After each development session, add a new dated section with:
1. Features added (✅)
2. Bugs fixed (🐛)
3. Known issues updated
4. Any new conventions or patterns established