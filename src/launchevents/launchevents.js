/* global Office analyzeMail */

let currentEvent = null;

Office.onReady(() => {
    console.log("Launch events runtime loaded");
});

function onMessageSendHandler(event) {

    currentEvent = event;/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 */

// ========================= EVENT HANDLERS =========================

// Handles the OnMessageSend event to encrypt the message body and attachments.
function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    { asyncContext: { event: event } },
    (result) => {
      const event = result.asyncContext.event;
      if (result.status === Office.AsyncResultStatus.Failed) {
        handleError(event,"Unable to encrypt the contents of this message.",`Failed to get body content: ${result.error.message}.`);
        return;
      }

      const body = result.value;
      const encryptedBody = encrypt(body);

      const placeholderMessage = `
        <div style="font-family: Segoe UI, Helvetica, Arial, sans-serif; font-size: 14px; color: #333;">
          <h2 style="color: #0078d4;">🔒 This message is encrypted</h2>
          <p>This message has been encrypted by the add-in.</p>
          <p>You must install the add-in to read it.</p>
        </div>
      `.trim();

      Office.context.mailbox.item.body.setAsync(
        placeholderMessage,
        { asyncContext: { event: event }, coercionType: Office.CoercionType.Html },
        (result) => {
          const event = result.asyncContext.event;
          if (result.status === Office.AsyncResultStatus.Failed) {
            handleError(event,"Unable to encrypt message.",`Failed to set placeholder: ${result.error.message}`);
            return;
          }

          Office.context.mailbox.item.addFileAttachmentFromBase64Async(
            encryptedBody,
            "encrypted_body.txt",
            { asyncContext: { event: event } },
            () => event.completed({ allowEvent: true })
          );
        }
      );
    }
  );
}

// Handles the OnMessageRead event to decrypt message.
function onMessageReadHandler(event) {
  const attachments = Office.context.mailbox.item.attachments;
  if (!attachments.length) {
    event.completed({ allowEvent: true });
    return;
  }

  Office.context.mailbox.item.getAttachmentContentAsync(
    attachments[0].id,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        event.completed({ allowEvent: false });
        return;
      }

      const decrypted = decrypt(result.value.content);

      event.completed({
        allowEvent: true,
        emailBody: {
          coercionType: Office.CoercionType.Html,
          content: decrypted
        }
      });
    }
  );
}

// ========================= ENCRYPTION =========================

function customBtoa(str) {
  return btoa(unescape(encodeURIComponent(str)));
}

function customAtob(str) {
  return decodeURIComponent(escape(atob(str)));
}

function encrypt(input, key = "OfficeAddInSampleKey") {
  let result = "";
  for (let i = 0; i < input.length; i++) {
    result += String.fromCharCode(input.charCodeAt(i) ^ key.charCodeAt(i % key.length));
  }
  return customBtoa(result);
}

function decrypt(input, key = "OfficeAddInSampleKey") {
  const decoded = customAtob(input);
  let result = "";
  for (let i = 0; i < decoded.length; i++) {
    result += String.fromCharCode(decoded.charCodeAt(i) ^ key.charCodeAt(i % key.length));
  }
  return result;
}

// ========================= HELPERS =========================

function handleError(event, message, log) {
  console.log(log);
  event.completed({ allowEvent: false, errorMessage: message });
}

// ========================= 🔥 SAFE INITIALIZATION (FIX) =========================

let handlersRegistered = false;

function registerEventHandlers() {
  if (handlersRegistered) return;

  try {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    Office.actions.associate("onMessageReadHandler", onMessageReadHandler);
    handlersRegistered = true;
    console.log("Handlers registered");
  } catch (err) {
    console.error("Registration failed:", err);
  }
}

// Modern Outlook runtime
if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      registerEventHandlers();
    }
  });
}

// Cold start fallback (VERY important for event-based add-ins)
(function () {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.mailbox) {
      registerEventHandlers();
    }
  } catch (e) {}
})();

    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Text,
        function (bodyResult) {

            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({ allowEvent: true });
                return;
            }

            const body = bodyResult.value;

            Office.context.mailbox.item.to.getAsync(function (recResult) {

                let recipients = [];

                if (recResult.status === Office.AsyncResultStatus.Succeeded)
                    recipients = recResult.value.map(r => r.emailAddress);

                const decision = analyzeMail(body, recipients);

                if (decision.action === "ALLOW") {
                    event.completed({ allowEvent: true });
                    return;
                }

                if (decision.action === "BLOCK") {
                    event.completed({
                        allowEvent: false,
                        errorMessage: decision.message
                    });
                    return;
                }

                openWarningDialog(decision.message);
            });
        }
    );
}

function openWarningDialog(message) {

    Office.context.ui.displayDialogAsync(
        "https://localhost:3000/warn.html?msg=" + encodeURIComponent(message),
        { height: 40, width: 30, displayInIframe: true },
        function (result) {

            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                currentEvent.completed({ allowEvent: false });
                return;
            }

            const dialog = result.value;

            dialog.addEventHandler(
                Office.EventType.DialogMessageReceived,
                function (arg) {

                    if (arg.message === "SEND_ANYWAY")
                        currentEvent.completed({ allowEvent: true });
                    else
                        currentEvent.completed({ allowEvent: false });

                    dialog.close();
                }
            );
        }
    );
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);