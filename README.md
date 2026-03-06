# Voxor Speaker Ready Agent

Minimal local agent for Speaker Ready desks. Runs on SR laptops alongside the browser-based Speaker Ready UI on PresenterHubV2.

## What it does

- **Open** — Downloads a presentation from PresenterHubV2, saves to local `Working/` folder, opens in PowerPoint
- **Publish** — Uploads the edited file from `Working/` back to PresenterHubV2
- **Heartbeat** — Registers as a `speakerready` agent with PresenterHubV2

## Requirements

- Node.js 18+ (no npm install needed — zero dependencies)
- Network access to the PresenterHubV2 server

## Setup

1. Copy this folder to `C:\SRAgent` on the SR laptop
2. Edit `sr-config.json`:
   ```json
   {
     "phServer": "http://<PHV2-SERVER-IP>:8088",
     "agentToken": "<your-agent-token>",
     "activeRoom": "Ballroom A"
   }
   ```
3. Double-click `start-sr.cmd`

## Usage

The agent runs in the background on port **8899**. Open `speakerready.html` on the PresenterHubV2 server in a browser — it auto-detects the local agent and shows **Open / Publish / Folder** buttons.

Without the agent (tablets, phones, remote access), the Speaker Ready page falls back to browser download/upload.

## File structure

```
C:\SRAgent\
  agent.js          — the micro-agent (Node.js, ~250 lines)
  sr-config.json    — server IP, agent token, active room
  start-sr.cmd      — double-click launcher
  Working/          — created automatically, local staging folder
    <Room>/
      <SessionFolder>/
        <SessionFolder>.pptx
```

## Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Agent status |
| GET | `/config` | Current config |
| POST | `/open` | Download PPTX from PHV2 → Working/ → open in PowerPoint |
| POST | `/publish` | Upload edited file from Working/ → PHV2 |
| POST | `/open-folder` | Open Working folder in Explorer |

## Environment variables (optional overrides)

| Variable | Default | Description |
|----------|---------|-------------|
| `PH_SERVER` | from sr-config.json | PresenterHubV2 server URL |
| `AGENT_TOKEN` | from sr-config.json | Agent authentication token |
| `PORT` | 8899 | Agent listen port |
| `SR_ROOT` | parent of agent.js | Root folder for Working/ and cache/ |
