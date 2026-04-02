// background.js — Service worker that opens the side panel when the extension icon is clicked.

// When the user clicks the extension icon, open the side panel
chrome.action.onClicked.addListener(async (tab) => {
  await chrome.sidePanel.open({ tabId: tab.id });
});

// Allow the side panel to open programmatically
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
