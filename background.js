/**
 * Background script for Outlook RWZ to Thunderbird Filter Add-on
 */

/* global messenger, browser */

const api = typeof messenger !== 'undefined' ? messenger : browser;

// Listen for action button click in toolbar
const actionApi = api.action || api.browserAction;

if (actionApi && actionApi.onClicked) {
  actionApi.onClicked.addListener(async () => {
    const url = api.runtime.getURL('ui/index.html');

    // Check if the tab is already open
    const tabs = await api.tabs.query({ url });
    if (tabs && tabs.length > 0) {
      // Focus existing tab
      await api.tabs.update(tabs[0].id, { active: true });
      if (tabs[0].windowId) {
        await api.windows.update(tabs[0].windowId, { focused: true });
      }
    } else {
      // Open new tab
      await api.tabs.create({ url });
    }
  });
}

// On install, log and optionally open the options / tool tab
api.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    const url = api.runtime.getURL('ui/index.html');
    api.tabs.create({ url }).catch(() => {});
  }
});
