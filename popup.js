document.getElementById('export').addEventListener('click', async () => {
    const [tab] = await chrome.tabs.query({active: true, currentWindow: true});
    if (!tab?.id) return;

    try {
        await chrome.tabs.sendMessage(tab.id, {type: 'EXPORT_ICS'});
        window.close();
    } catch (e) {
        alert("Couldn't reach the page. Make sure you're on the Workday courses page and it's fully loaded.");
    }
});
