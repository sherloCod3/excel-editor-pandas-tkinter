# SecOps SOC Agent — n8n Workflow

This directory contains the n8n workflow that powers the **THREAT_INTEL** page of the SecOps Portal.

---

## Architecture

```
POST /webhook/soc-scan
        ↓
[1] Webhook → Receive request from Streamlit
[2] Set     → Extract & sanitise "target" field
[3] IF      → Is it an IP address or a domain/URL?
        │
        ├─ [IP path]
        │   [4] HTTP → ip-api.com   (country, ISP, city)
        │   [5] HTTP → VirusTotal   (IP threat report)
        │
        └─ [Domain path]
            [6] HTTP → VirusTotal   (URL/domain report)
        │
[7] Merge   → Combine geo + VirusTotal data
[8] Set     → Normalise to canonical schema
        │
[9] IF      → OPENAI_API_KEY set?  ← env-var feature flag
        │
        ├─ [Yes] [10] LLM Chain → plain-English threat summary
        └─ [No]  [11] Set       → "AI summary disabled" message
        │
[12] Merge  → Attach summary to payload
[13] Set    → Build final response object
[14] Respond → Return JSON to Streamlit
```

**Response schema:**
```json
{
  "status": "success",
  "target": "8.8.8.8",
  "threat_score": 0,
  "location": "United States",
  "known_malicious": false,
  "summary": "8.8.8.8 is Google's public DNS server...",
  "details": "0 of 91 VirusTotal engines flagged this target. ISP: Google LLC"
}
```

---

## Prerequisites

| Tool | Version | Install |
|------|---------|---------|
| Docker + Compose | v2+ | [docs.docker.com](https://docs.docker.com/engine/install/) |
| VirusTotal API key | — | [virustotal.com/gui/join-us](https://www.virustotal.com/gui/join-us) (free) |
| OpenAI key *(optional)* | — | [platform.openai.com](https://platform.openai.com) |

---

## Setup

### 1. Configure environment variables

```bash
# From the project root
cp .env.example .env
```

Open `.env` and fill in:

```ini
N8N_USER=admin
N8N_PASSWORD=your_secure_password     # change this!
VIRUSTOTAL_API_KEY=your_vt_key_here

# Optional — leave blank to disable AI summaries
OPENAI_API_KEY=sk-...
```

### 2. Start n8n

```bash
# From the project root
docker compose up -d

# Verify it started
docker compose ps
docker compose logs -f n8n
```

n8n is ready when you see `Editor is now accessible via: http://localhost:5678`

### 3. Open the n8n UI

Navigate to **http://localhost:5678** and log in with the credentials from your `.env`.

### 4. Import the workflow

1. Click **+** → **Add workflow** (top right)
2. Click the **⋮** menu → **Import from file**
3. Select `n8n/soc_agent_workflow.json`
4. The workflow canvas will open with all nodes pre-wired.

### 5. Configure credentials

The workflow needs two credentials configured in n8n (they read from your env vars):

**VirusTotal** *(used inside HTTP Request nodes)*
- The API key is passed directly as a header using `{{ $env.VIRUSTOTAL_API_KEY }}` — no credential entry needed.

**OpenAI** *(only if AI summaries are enabled)*
- Go to **Settings → Credentials → Add credential → OpenAI**
- Name it exactly: `OpenAI - SecOps`
- API Key: `{{ $env.OPENAI_API_KEY }}`
- Base URL: `{{ $env.OPENAI_BASE_URL }}`

### 6. Activate the workflow

Click the **Inactive** toggle (top right of the workflow canvas) to activate it.

### 7. Copy the webhook URL

Click on the **Webhook - Receive Scan Request** node → copy the **Production URL**. It will look like:
```
http://localhost:5678/webhook/soc-scan
```

### 8. Connect to Streamlit

1. Open the SecOps Portal in your browser.
2. Navigate to **THREAT_INTEL**.
3. Paste the webhook URL into the **Webhook Config** field in the sidebar.
4. Click **💾 PERSIST URL**.
5. The badge changes from `MOCK MODE` to `LIVE MODE`.

### 9. Verify end-to-end

Run a test scan on `8.8.8.8` (Google DNS). You should see:
- **Threat Score:** 0–5
- **Location:** United States
- **Threat Status:** ✓ CLEAR
- **AI Summary** (if OpenAI key set): plain-English description

---

## Using Ollama Instead of OpenAI

Run AI summaries locally for free with [Ollama](https://ollama.com):

```bash
# Install and pull a model (llama3 is a good choice)
ollama pull llama3

# In your .env, set:
OPENAI_API_KEY=ollama        # any non-empty string enables the AI branch
OPENAI_BASE_URL=http://host.docker.internal:11434/v1
```

Restart n8n and update the OpenAI credential base URL. The `chainLlm` node is compatible with any OpenAI-spec API.

> **Note:** `host.docker.internal` resolves to your host machine from inside the Docker container on Linux you may need to add `--add-host=host.docker.internal:host-gateway` to the docker-compose extra_hosts.

---

## Stopping n8n

```bash
docker compose down          # stop containers, keep data
docker compose down -v       # stop and delete all data (destructive)
```

---

## Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| `LIVE MODE` shows but results look wrong | Workflow not activated | Click the Inactive toggle in n8n |
| VirusTotal returns 403 | Wrong API key | Re-check `VIRUSTOTAL_API_KEY` in `.env`, restart container |
| AI summary always disabled | Key not set or empty | Check `OPENAI_API_KEY` is non-empty in `.env` |
| `docker compose up` fails port conflict | Port 5678 in use | Change `N8N_PORT=5679` in `.env` |
| ip-api.com returns `fail` status | Private IP submitted | ip-api.com doesn't resolve RFC-1918 addresses (expected behaviour) |
