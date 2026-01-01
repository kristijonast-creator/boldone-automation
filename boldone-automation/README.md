# BoldOne Automation Suite

A collection of Google Apps Script automations for managing ad creative workflows at CommerceCore.

## 🏗️ Systems Overview

| System | Purpose | Status |
|--------|---------|--------|
| **BoldOne Creative Hub** | Central config web app, team/product management | ✅ Active |
| **Ad Naming Sheet** | Batch creation, folder automation, ad naming conventions | ✅ Active |
| **Sprint → ClickUp** | Auto task creation from sprint sheets | ✅ Active |
| **Atria Reports** | Performance reporting, Slack notifications | ✅ Active |
| **Drive Naming QA** | File naming validation | ✅ Active |

## 📁 Repository Structure

```
boldone-automation/
├── boldone-hub/          # Central web app
│   ├── code.gs           # Backend functions
│   └── index.html        # Frontend (SPA)
├── ad-naming/            # Ad naming automation
│   ├── code.gs           # Main backend
│   └── autobatchid.gs    # Batch ID generator
├── sprint-clickup/       # ClickUp integration
│   └── autobatchcreator.gs
├── atria-reports/        # Reporting system
│   └── AtriaSlackAutomation.gs
├── drive-qa/             # File validation
│   └── drivenamingchecker.gs
├── config/               # Setup scripts
│   └── MainConfig.gs
└── docs/                 # Documentation
    ├── CommerceCore_Project_Instructions.md
    ├── GoogleAppsScript_SKILL.md
    └── CommerceCore_Development_Log.md
```

## 🔧 Central Configuration

All systems connect to a central Google Sheet for configuration:

**Config Sheet ID:** `1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk`

### Config Tabs
- **Products** - Product keys, IDs, CPA targets
- **Team Members** - Users, roles, permissions
- **Product Folders** - Drive folder IDs per locale
- **Footage Folders** - Raw footage URLs
- **ClickUp Config** - API keys, space/folder IDs
- **Valid Values** - Dropdowns for platforms, locales

## 🚀 Getting Started

### For Developers

1. Clone this repo
2. Install [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
3. Login: `clasp login`
4. Clone the Apps Script project you want to work on:
   ```bash
   clasp clone SCRIPT_ID --rootDir ./boldone-hub
   ```

### For Users

Access the BoldOne Creative Hub at: [Web App URL]

## 📝 Development Notes

See `docs/CommerceCore_Development_Log.md` for detailed change history.

## ⚠️ Important Conventions

- **Private functions** end with underscore: `functionName_()`
- **Cache duration**: 300 seconds (5 minutes)
- **ClickUp IDs**: Always strings, use `==` not `===`
- **API calls**: Always wrap in try-catch

## 🔐 Security

This is a private repository. Never commit:
- API keys directly in code (use Config sheet)
- Personal access tokens
- Sensitive user data

---

*Maintained by CommerceCore Creative Team*