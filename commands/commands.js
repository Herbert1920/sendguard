Office.initialize = () => {};

function onMessageSend(event) {

    const item = Office.context.mailbox.item;

    Promise.all([
        new Promise(res => item.to.getAsync(r => res(r.value || []))),
        new Promise(res => item.cc.getAsync(r => res(r.value || []))),
        new Promise(res => item.bcc.getAsync(r => res(r.value || [])))
    ]).then(values => {

        const data = {
            to: values[0],
            cc: values[1],
            bcc: values[2]
        };

        const encoded = btoa(JSON.stringify(data));

        // 🔥 הכתובת שלך + dialog
        const url =
          `https://Herbert1920.github.io/sendguard/dialog/dialog.html?data=${encoded}`;

        Office.context.ui.displayDialogAsync(
            url,
            { height: 55, width: 45, displayInIframe: false },
            result => {

                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    event.completed({ allowEvent: true });
                    return;
                }

                const dialog = result.value;

                dialog.addEventHandler(
                    Office.EventType.DialogMessageReceived,
                    args => {
                        dialog.close();
                        event.completed({ allowEvent: args.message === "confirmed" });
                    }
                );

                dialog.addEventHandler(
                    Office.EventType.DialogEventReceived,
                    () => {
                        event.completed({ allowEvent: false });
                    }
                );

            }
        );

    });
}
