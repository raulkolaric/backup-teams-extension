document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('sync-form');
  const emailInput = document.getElementById('email');
  const secretInput = document.getElementById('secret');
  const submitBtn = document.getElementById('sync-btn');
  const btnText = submitBtn.querySelector('.btn-text');
  const spinner = submitBtn.querySelector('.spinner');
  
  const statusContainer = document.getElementById('status-container');
  const statusMessage = document.getElementById('status-message');

  // Load saved data from Chrome sync storage
  chrome.storage.local.get(['savedEmail', 'savedSecret'], (result) => {
    if (result.savedEmail) emailInput.value = result.savedEmail;
    if (result.savedSecret) secretInput.value = result.savedSecret;
  });

  const showStatus = (message, type) => {
    statusContainer.className = `status-container ${type}`;
    statusMessage.textContent = message;
  };

  const clearStatus = () => {
    statusContainer.className = 'status-container hidden';
    statusMessage.textContent = '';
  };

  const setLoading = (isLoading) => {
    submitBtn.disabled = isLoading;
    if (isLoading) {
      btnText.classList.add('hidden');
      spinner.classList.remove('hidden');
    } else {
      btnText.classList.remove('hidden');
      spinner.classList.add('hidden');
    }
  };

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    clearStatus();
    setLoading(true);

    const email = emailInput.value.trim();
    const secret = secretInput.value.trim();

    // Persist credentials locally
    chrome.storage.local.set({ savedEmail: email, savedSecret: secret });

    try {
      // 1. Get the current active tab
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

      if (!tab) {
         throw new Error("Cannot access the current tab.");
      }

      // 2. Validate we are on Microsoft Teams
      if (!tab.url || !tab.url.includes('teams.microsoft.com')) {
        throw new Error("Please open teams.microsoft.com and log in first.");
      }

      // 3. Inject our payload extractor script
      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: extractMsalToken
      });

      const bestToken = results[0]?.result;

      if (!bestToken) {
        throw new Error("No Graph token found. Click on a few files in Teams and try again.");
      }

      // 4. Send token to our Backend
      await syncWithBackend(email, bestToken, secret);

      showStatus("Token successfully synced with Backup Teams backend!", "success");

    } catch (err) {
      showStatus(err.message || "An unexpected error occurred.", "error");
    } finally {
      setLoading(false);
    }
  });
});

/**
 * -------------------------------------------------------------
 * CONTENT SCRIPT LOGIC (Executes in the context of the Teams tab)
 * -------------------------------------------------------------
 * Why keeping this strict and lean is important:
 * We don't want to load massive JWT parsing libraries into the user's tab.
 * We stick strictly to native APIs to parse localStorage.
 */
function extractMsalToken() {
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

      if (!secret || typeof secret !== 'string' || !secret.startsWith('ey')) continue;
      if (!target || typeof target !== 'string') continue;
      if (!target.toLowerCase().includes('graph')) continue;
      if (expires && (parseInt(expires, 10) < now)) continue;  // expired

      candidates.push({ token: secret, scope: target, expires: expires });
    } catch(e) {
      // Silent catch: localStorage might contain malformed JSON from other apps.
    }
  }

  // Sort to get the token that expires latest
  candidates.sort((a, b) => (parseInt(b.expires) || 0) - (parseInt(a.expires) || 0));
  return candidates.length > 0 ? candidates[0].token : null;
}

/**
 * -------------------------------------------------------------
 * API LAYER (Executes in popup context)
 * -------------------------------------------------------------
 */
async function syncWithBackend(email, accessToken, syncSecret) {
  const response = await fetch('https://api.backup-teams.com/auth/sync-token', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      email: email,
      access_token: accessToken,
      refresh_token: null,
      sync_secret: syncSecret
    })
  });

  if (!response.ok) {
    let errorMsg = 'Failed to sync with API backend.';
    try { 
      const data = await response.json(); 
      if (data.detail) errorMsg = data.detail;
    } catch (e) {}
    throw new Error(`Server Error: ${errorMsg}`);
  }

  return response.json();
}
