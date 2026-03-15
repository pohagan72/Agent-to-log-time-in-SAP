# SAP Hours Agent

Chrome/Edge extension that uses Claude AI to automate SAP Fiori time entry via natural language. Injects a chat panel into any SAP Fiori page, communicates with SAP's OData API, and auto-discovers user info at runtime.

## Architecture

```
User <-> Chat Panel (iframe) <-> Content Script <-> SAP OData API
                |
         Claude API (Azure)
```

- `content.js` — Injected at `document_idle`. Auto-discovers user identity from SAP, injects the chat panel iframe, handles all OData calls (`ZHR_TIME_ENTRY_SRV`) using session cookies, bridges iframe <-> SAP via `postMessage`.
- `sidepanel.html/js` — Chat UI in extension context. Calls Claude API directly. Implements the agentic loop: Claude emits JSON actions, extension executes them, results feed back into conversation.
- `background.js` — Toggles panel on extension icon click.
- `build.js` — Reads `.env`, generates `config.js` and `manifest.json` (both gitignored).

### Agent Actions

Claude can emit these JSON action blocks, which the extension executes and feeds back:

| Action | Description |
|--------|-------------|
| `SEARCH_PROJECTS` | Search `ProjectSearchSet` by ID or description |
| `GET_ACTIVITIES` | Fetch activities for a project from `ProjectActivitySet` |
| `GET_RECORDED_HOURS` | Read time entries from `TimeEntrySet` with date range filter |
| `GET_CALENDAR_STATUS` | Check which days have hours via `CalendarMarkingSet` |
| `GET_WEEK_TOTAL` | Current week total from `DynTileInfoSet` |
| `ENTER_TIME` | Create entries via `$batch` POST to `TimeEntrySet` |
| `DELETE_ENTRY` | Delete entry by counter ID |
| `COPY_WEEK` | Copy or move entries between weeks via `TimeEntryCopyToSet` |
| `ADD_FAVORITE` | Add project/activity to SAP favorites |
| `REMOVE_FAVORITE` | Remove project/activity from SAP favorites |

### Auto-Discovery

No manual user configuration required. On first load:

1. Reads SAP username from the FLP `<meta name="sap.ushellConfig.serverSideConfig">` tag
2. Queries `HRInfoSet` for personnel number, cost center, company, job title, role, work hours
3. Fetches favorites via `HRInfoSet?$expand=NavFavSet`
4. Caches in `chrome.storage.local` (7-day TTL)

Works from any SAP Fiori page, not just Time Entry.

### OData Details

All calls go through the content script (same origin, inherits session cookies):

- **Reads**: OData GET with `$filter`, `$expand`, `$format=json`
- **Writes**: `$batch` POST with changeset boundary, CSRF token via `HEAD` with `x-csrf-token: Fetch`
- **Deletes**: OData `DELETE` by entity key
- **Defaults**: Fetches `TimeEntryDetailSet` before creating entries to get required field defaults (Network, CompanyCode, billing flags, etc.)
- **Activity resolution**: Resolves activity names ("Coding") to SAP codes ("0010") via `ProjectActivitySet`

## Setup

```bash
cp .env.example .env    # Add API key, endpoint, SAP hostname
node build.js           # No dependencies required
# Load extension/ as unpacked in edge://extensions or chrome://extensions
```

## Configuration

`.env` requires 3 variables:

| Variable | Description |
|----------|-------------|
| `AZURE_CLAUDE_API_KEY` | Azure AI Foundry API key |
| `AZURE_CLAUDE_ENDPOINT` | Claude messages endpoint URL |
| `SAP_HOSTNAME` | SAP Fiori hostname |

## Usage

The panel appears on any SAP page. Quick action buttons and example prompts adapt to the user's actual favorite projects.

```
"8 hours on Agentics AI all of last week"
"4h AI Platform and 4h Translate every day, Friday off"
"Which days am I missing this month?"
"Delete the 4h entry on Wednesday"
"Copy last week to this week"
```

All destructive actions (create, delete, copy) require user confirmation before execution.

## Project Structure

```
sap-hours/
  .env.example        # Template
  build.js            # Generates config from .env
  extension/
    background.js     # Icon click handler
    content.js        # OData API + auto-discovery + panel injection
    sidepanel.html    # Chat UI
    sidepanel.js      # Claude API + agentic loop + state persistence
    icons/
    config.js         # Generated (gitignored)
    manifest.json     # Generated (gitignored)
```

## Security

- API keys in `.env` only (gitignored)
- User info discovered at runtime, not stored in config files
- Content script runs same-origin with SAP — no credentials sent to third parties except the configured Claude endpoint
- Claude API called from extension context with explicit `host_permissions`

## Requirements

- Node.js (build only, no packages)
- Edge or Chrome
- SAP Fiori instance with `ZHR_TIME_ENTRY_SRV` OData service
- Azure AI Foundry endpoint with Claude model
