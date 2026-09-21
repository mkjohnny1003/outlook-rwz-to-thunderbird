/**
 * Background script for Outlook RWZ to Thunderbird Filter Add-on
 */

/* global messenger, browser */

const api = typeof messenger !== 'undefined' ? messenger : browser;

// Listen for action button click in toolbar
const actionApi = api.browserAction || api.action;

async function openAppTab() {
  try {
    const tabs = await api.tabs.query({});
    const existing = tabs.find(t => t.url && t.url.includes('ui/index.html'));
    if (existing) {
      await api.tabs.update(existing.id, { active: true });
      if (existing.windowId) {
        await api.windows.update(existing.windowId, { focused: true });
      }
      return;
    }
  } catch (_) {}

  try {
    await api.tabs.create({ url: 'ui/index.html' });
  } catch (_) {}
}

if (actionApi && actionApi.onClicked) {
  actionApi.onClicked.addListener(() => {
    openAppTab();
  });
}

// On install, open the options / tool tab safely
api.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    // Delay slightly to ensure UUID and contexts are ready
    setTimeout(() => {
      openAppTab();
    }, 500);
  }
});
