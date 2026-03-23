Office.initialize = () => {

    const params = new URLSearchParams(location.search);
    const raw = params.get("data");
    if (!raw) return;

    const recipients = JSON.parse(atob(raw));
    const list = document.getElementById("list");
    list.innerHTML = "";

    ["to", "cc", "bcc"].forEach(type => {
        recipients[type].forEach(r => {
            const div = document.createElement("div");
            div.className = "item";
            div.textContent = r.displayName || r.emailAddress || r.email;
            list.appendChild(div);
        });
    });
};

function confirmSend() {
    Office.context.ui.messageParent("confirmed");
}

function cancelSend() {
    Office.context.ui.messageParent("cancelled");
}
