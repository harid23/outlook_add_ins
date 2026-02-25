Office.onReady(() => {
    const params = new URLSearchParams(window.location.search);
    document.getElementById("message").innerText = params.get("msg");
});

function sendAnyway() {
    Office.context.ui.messageParent("SEND_ANYWAY");
}

function cancelSend() {
    Office.context.ui.messageParent("CANCEL");
}