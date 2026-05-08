Office.onReady((info) => {
    // This confirms the script is actually running
    console.log("Office.onReady triggered");
    
    const btn = document.getElementById("insertBtn");
    if (btn) {
        btn.onclick = insertText;
    }
});

function insertText() {
    try {
        const userText = document.getElementById("userInput").value;
        const combinedString = userText + " Hello World!";

        // ERROR CHECK 1: Is the Office API loaded?
        if (typeof Office === 'undefined' || !Office.context) {
            alert("Office.js is not loaded properly.");
            return;
        }

        // ERROR CHECK 2: Are we in a draft?
        if (!Office.context.mailbox.item) {
            alert("No active email found. Make sure you're in COMPOSE mode.");
            return;
        }

        // Perform the insertion
        Office.context.mailbox.item.body.setSelectedDataAsync(
            combinedString,
            { coercionType: Office.CoercionType.Text },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    alert("Error: " + asyncResult.error.message);
                } else {
                    console.log("Success!");
                }
            }
        );
    } catch (err) {
        alert("Catch Error: " + err.message);
    }
}
