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

function isSupportedSidePanelUrl(url) {
  if (typeof url !== "string" || !url) {
    return false;
  }

  return /^(https?|file):\/\//i.test(url);
}

async function setSidePanelEnabledForTab(tabId, url) {
  if (!chrome.sidePanel || typeof chrome.sidePanel.setOptions !== "function") {
    return;
  }

  try {
    await chrome.sidePanel.setOptions({
      tabId,
      path: "popup.html",
      enabled: isSupportedSidePanelUrl(url)
    });
  } catch (error) {
    console.warn("Failed to sync side panel state for tab.", { tabId, url, error });
  }
}

function syncAllOpenTabs() {
  if (!chrome.tabs || typeof chrome.tabs.query !== "function") {
    return;
  }
  chrome.tabs.query({}, (tabs) => {
    if (!Array.isArray(tabs)) {
      return;
    }
    for (const tab of tabs) {
      if (tab && typeof tab.id === "number") {
        void setSidePanelEnabledForTab(tab.id, tab.url || "");
      }
    }
  });
}

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  const url = changeInfo.url || tab.url || "";
  void setSidePanelEnabledForTab(tabId, url);
});

chrome.tabs.onActivated.addListener((activeInfo) => {
  if (!activeInfo || typeof activeInfo.tabId !== "number") {
    return;
  }
  chrome.tabs.get(activeInfo.tabId, (tab) => {
    if (chrome.runtime.lastError || !tab) {
      return;
    }
    void setSidePanelEnabledForTab(activeInfo.tabId, tab.url || "");
  });
});

syncAllOpenTabs();
