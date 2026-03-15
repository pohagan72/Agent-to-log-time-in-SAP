# SAP Hours Agent

A Chrome/Edge browser extension that uses Claude AI to automate SAP Fiori time entry via natural language. Embeds a persistent chat panel directly into the SAP Time Entry page and interacts with SAP's OData API to read, search, and create time entries.

## Architecture

```
User ←→ Chat Panel (iframe) ←→ Content Script ←→ SAP OData API
                ↕
         Claude API (Azure)
```

- **Content script** (`content.js`) — Injected into SAP pages at `document_idle`. Injects a persistent side panel as an iframe, handles all SAP OData V2 API calls (`ZHR_TIME_ENTRY_SRV`) using the page's session cookies, and bridges communication between the iframe and SAP via `postMessage`.
- **Side panel** (`sidepanel.html` / `sidepanel.js`) — Runs inside the injected iframe in extension context. Manages the chat UI, calls the Claude API directly, and implements the agentic loop (Claude emits JSON actions → extension executes them → feeds results back to Claude).
- **Background worker** (`background.js`) — Handles extension icon click to toggle panel visibility.
- **Build step** (`build.js`) — Reads `.env` and generates `config.js`, `user-config.js`, and `manifest.json` with user-specific values. All generated files are gitignored.

### Agentic Loop

Claude's responses can contain JSON action blocks that the extension executes automatically:

| Action | Description |
|--------|-------------|
| `SEARCH_PROJECTS` | Searches SAP's `ProjectSearchSet` by ID or description |
| `GET_ACTIVITIES` | Fetches available activities for a project from `ProjectActivitySet` |
| `GET_RECORDED_HOURS` | Reads past time entries from `TimeEntrySet` with date range filter |
| `ENTER_TIME` | Creates entries via `$batch` POST to `TimeEntrySet` (requires CSRF token) |

Results are fed back into the conversation as system messages, allowing Claude to chain multiple actions before responding to the user.

### OData Details

All API calls go through the content script (same origin as SAP, inherits session cookies):

- **Reads**: Standard OData GET with `$filter`, `$expand`, `$format=json`
- **Writes**: Multipart `$batch` POST with changeset boundary, requires CSRF token fetched via `x-csrf-token: Fetch` header
- **Defaults**: Before creating an entry, fetches `TimeEntryDetailSet` to get required field defaults (Network, CompanyCode, billing flags, etc.)
- **Activity resolution**: Claude may send activity names ("Coding") — the extension resolves these to SAP activity codes ("0010") via `ProjectActivitySet` lookup

## Setup

```bash
# 1. Clone and configure
cp .env.example .env
# Edit .env with your API keys and SAP details

# 2. Build (no dependencies required)
node build.js

# 3. Load extension
# Edge:  edge://extensions → Developer mode → Load unpacked → select extension/
# Chrome: chrome://extensions → Developer mode → Load unpacked → select extension/
```

The build generates three gitignored files in `extension/`:
- `config.js` — API endpoint and key
- `user-config.js` — SAP hostname, personnel number, user details
- `manifest.json` — Dynamic host permissions and content script matching

## Configuration

All configuration lives in `.env` (see `.env.example`):

| Variable | Description |
|----------|-------------|
| `AZURE_CLAUDE_API_KEY` | Azure AI Foundry API key |
| `AZURE_CLAUDE_ENDPOINT` | Claude messages endpoint URL |
| `AZURE_CLAUDE_MODEL` | Model name (e.g., `claude-sonnet-4-5`) |
| `SAP_HOSTNAME` | Your SAP Fiori hostname |
| `SAP_CLIENT` | SAP client number |
| `SAP_PERS_NUMBER` | Your SAP personnel number |
| `SAP_USER_NAME` | Your SAP username |
| `SAP_USER_DISPLAY_NAME` | Your name (used in Claude's system prompt) |
| `SAP_COST_CENTER` | Your cost center (used in Claude's system prompt) |

## Usage

Navigate to your SAP Time Entry page. The agent panel appears automatically on the right side. Examples:

- *"I spent all week on Agentics AI, 8 hours a day"*
- *"Last week was 4h Agentics AI and 4h AI Platform every day, I took Friday off"*
- *"What hours did I record in February?"*
- *"Search for conference projects"*

The agent will propose entries with dates, projects, hours, and descriptions for your review before submitting anything to SAP.

## Project Structure

```
sap-hours/
  .env.example          # Template — copy to .env
  .gitignore
  build.js              # Generates config from .env
  extension/
    background.js       # Extension icon click handler
    content.js          # SAP OData API + panel injection + postMessage bridge
    sidepanel.html      # Chat UI (loaded as iframe)
    sidepanel.js        # Claude API + agentic loop + state persistence
    icons/              # Extension icons
    config.js           # Generated (gitignored)
    user-config.js      # Generated (gitignored)
    manifest.json       # Generated (gitignored)
```

## Security Notes

- API keys and personal info live exclusively in `.env` (gitignored)
- `config.js`, `user-config.js`, and `manifest.json` are generated and gitignored
- The content script runs in the SAP page context with same-origin access — no credentials are sent to any third party except the configured Claude API endpoint
- The Claude API call is made from the extension iframe context, which has explicit `host_permissions` for the configured Azure endpoint

## Requirements

- Node.js (build step only, no packages needed)
- Microsoft Edge or Google Chrome
- Access to a SAP Fiori Time Entry instance (`ZHR_TIME_ENTRY_SRV` OData service)
- Azure AI Foundry endpoint with Claude model deployed
