// Clicking the toolbar icon injects (or toggles) the floating panel on the current tab.
chrome.action.onClicked.addListener((tab) => {
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ["xlsx.mini.min.js", "content.js"]
  });
});
