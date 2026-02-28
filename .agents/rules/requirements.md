---
trigger: always_on
---

# Backup Teams: Chrome Extension Specification 

## Context & Architecture
Our backend API (`api.backup-teams.com`) is a Python/FastAPI service that backups files from Microsoft Teams directly to an AWS S3 bucket.

Normally, the backend would use a standard OAuth 2.0 flow to request `Files.Read.All` from Microsoft Graph. However, the University IT admin has **blocked "User Consent for Applications"**. Furthermore, Microsoft enforces strict Conditional Access and IP Bot Protection, meaning our headless EC2 server cannot log in directly using Playwright or Selenium.

**The Solution: "Bring Your Own Token" (BYOT) via Chrome Extension.**
We need to build a lightweight Chrome Extension (Manifest V3) that the student installs. When they log into the official Microsoft Teams web app (`teams.microsoft.com`), the extension will quietly extract their cached MSAL (Microsoft Authentication Library) token from `localStorage` and `POST` it to our backend. 

---

## Extension Requirements

1. **Permissions:** The extension needs permissions to read `teams.microsoft.com` tabs and execute scripts (`activeTab`, `scripting`, `storage`).
2. **User Interface (Popup):** A simple UI where the user can:
   - Enter their Email.
   - Enter a `sync_secret` (which acts as an API key to our backend).
   - Click a "Sync Tokens" button.
3. **The Extraction Logic:** When the user clicks "Sync Tokens", the extension must find the open `teams.microsoft.com/v2/` tab, inject a content script to read `window.localStorage`, and parse the MSAL cache.
4. **The Network Request:** Once the token is extracted, the extension makes an HTTP `POST` request to `https://api.backup-teams.com/auth/sync-token`.

---

## 🏗️ Technical Details

### 1. The LocalStorage Extraction Script
The Teams Web App uses MSAL.js, which caches raw JWTs in `localStorage`. The keys are unpredictable variants of `<client_id>-login.windows.net...`.

You must iterate through `localStorage.length`, parse any JSON objects, and look for an object where the `target` or `scope` contains `graph.microsoft.com` and the `secret` starts with `ey` (a JWT).

**Here is the exact JS logic verified to work:**
```javascript
const now = Math.floor(Date.now() / 1000);
const candidates = [];

for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    try {
        const raw = localStorage.getItem(key);
        if (!raw || raw.length < 100) continue;
        const obj = JSON.parse(raw);
        if (!obj || typeof obj !== 'object') continue;

        const secret  = obj.secret  || obj.access_token || obj.token;
        const target  = obj.target  || obj.scope        || obj.scopes || '';
        const expires = obj.expiresOn || obj.expires_on || obj.ext_expires_on || 0;

        if (!secret || typeof secret !== 'string' || !secret.startsWith('ey'))  continue;
        if (!target || typeof target !== 'string') continue;
        if (!target.toLowerCase().includes('graph')) continue;
        if (expires && (parseInt(expires, 10) < now)) continue;  // expired

        candidates.push({ token: secret, scope: target, expires: expires });
    } catch(e) {}
}

// Sort to get the token that expires latest
candidates.sort((a, b) => (parseInt(b.expires) || 0) - (parseInt(a.expires) || 0));
const bestToken = candidates.length > 0 ? candidates[0].token : null;
```

### 2. The API Contract
Once the Extension has the token, it must fire a fetch request to the backend.

**Endpoint:** `POST https://api.backup-teams.com/auth/sync-token`
**Headers:** `Content-Type: application/json`
**Payload:**
```json
{
  "email": "student@university.edu",
  "access_token": "eyJ0eXAiOiJKV1QiLCJhb...",
  "refresh_token": null,
  "sync_secret": "THE_SECRET_ENTERED_IN_THE_POPUP"
}
```

### 3. Edge Cases to Handle
- The user might not have a Teams tab open. The extension should prompt: "Please open teams.microsoft.com and log in first."
- The user might be logged into Teams, but hasn't navigated around enough for MSAL to cache the Graph token. The extension should prompt: "No Graph token found. Click on a few files in Teams and try again."

---

## Your Task (For the Target AI Agent)
Please generate the complete source code for this Chrome Extension using vanilla HTML/JS/CSS (no build tools like React/Webpack required to make it simple). Provide:
1. `manifest.json` (Manifest V3)
2. `popup.html` (A clean, modern Senior-level UI)
3. `popup.js` (Handling the form, chrome.scripting API, and the fetch request to the backend)
4. Any required background service workers or content scripts.
