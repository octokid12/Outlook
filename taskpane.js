Office.onReady((info) => {
    if (info.host === Office.Host.Outlook) {
        document.getElementById("insertBtn").onclick = insertText;
    }
});

function insertText() {
    const userValue = document.getElementById("userInput").value;
    // Combine the user string and "Hello World!"
    const fullText = userValue + " Hello World!";

    // Insert the text at the current cursor position in the email body
    Office.context.mailbox.item.body.setSelectedDataAsync(
        fullText,
        { coercionType: Office.CoercionType.Text },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            } else {
                console.log("Text inserted successfully!");
            }
        }
    );
}
