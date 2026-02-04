document.getElementById("downloadBtn").addEventListener("click", async () => {
    const statusDiv = document.getElementById("status");
    statusDiv.textContent = "Scanning...";
    statusDiv.className = "";

    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab) {
        statusDiv.textContent = "No active tab.";
        return;
    }

    chrome.scripting.executeScript({
        target: { tabId: tab.id, allFrames: true },
        files: ["xlsx.mini.min.js", "content.js"]
    }, () => {
        if (chrome.runtime.lastError) {
            statusDiv.textContent = "Error: " + chrome.runtime.lastError.message;
            statusDiv.className = "error";
            return;
        }
        
        chrome.tabs.sendMessage(tab.id, { action: "findAndExport" });
    });
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const statusDiv = document.getElementById("status");
    if (message.action === "downloadSuccess") {
        statusDiv.textContent = "Downloaded: " + message.filename;
        statusDiv.className = "success";
    } else if (message.action === "downloadError") {
        statusDiv.textContent = message.error;
        statusDiv.className = "error";
    }
});