// This function runs as soon as the Office Add-in engine is ready
Office.onReady((info) => {
    if (info.host === Office.Host.Outlook) {
        // Link the button click to our function
        document.getElementById("insertBtn").onclick = insertText;
    }
});

function insertText() {
    // 1. Get the value from the input box
    const userText = document.getElementById("userInput").value;
    
    // 2. Create the final string
    const combinedString = userText + " Hello World!";

    // 3. Check if we are in a compose window
    if (!Office.context.mailbox.item) {
        console.error("No item found. Make sure you are in a Compose window.");
        return;
    }

    // 4. Insert the data into the email body
    // This will insert at the current cursor position
    Office.context.mailbox.item.body.setSelectedDataAsync(
        combinedString,
        { coercionType: Office.CoercionType.Text },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Insertion failed: " + asyncResult.error.message);
            } else {
                // Success! Clear the input for the next use
                document.getElementById("userInput").value = "";
                console.log("Inserted successfully.");
            }
        }
    );
}
