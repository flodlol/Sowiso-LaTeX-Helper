async function enableSidePanelOnActionClick() {
  if (!chrome.sidePanel || typeof chrome.sidePanel.setPanelBehavior !== "function") {
    return;
  }

  try {
    await chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
  } catch (error) {
    console.warn("Failed to configure side panel behavior.", error);
  }
}

chrome.runtime.onInstalled.addListener(() => {
  enableSidePanelOnActionClick();
});

chrome.runtime.onStartup.addListener(() => {
  enableSidePanelOnActionClick();
});

enableSidePanelOnActionClick();

// Only enable the side panel on Sowiso pages.
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (!chrome.sidePanel) {
    return;
  }
  const url = tab.url || changeInfo.url || "";
  const isSowiso = /^https:\/\/([^/]*\.)?sowiso\.nl\//.test(url);
  chrome.sidePanel.setOptions({ tabId, enabled: isSowiso });
});
