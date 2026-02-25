/*********************** DLP ENGINE ************************/

const COMPANY_DOMAIN = "company.com";
const PAYROLL_MAILS = ["payroll@company.com", "finance@company.com"];

// Detect full PAN
function containsPAN(text) {
  const panRegex = /\b[A-Z]{5}[0-9]{4}[A-Z]\b/i;
  return panRegex.test(text);
}

// Detect masked PAN
function containsMaskedPAN(text) {
  const maskedRegex = /\b([A-Z]{4}\*{4}[A-Z])|(\*{5}[0-9]{4}[A-Z])\b/i;
  return maskedRegex.test(text);
}

// classify recipients
function classifyRecipients(recipients) {
  let external = false;
  let payroll = false;
  let internal = false;

  recipients.forEach(mail => {
    mail = mail.toLowerCase();

    if (PAYROLL_MAILS.includes(mail)) payroll = true;
    else if (mail.endsWith("@" + COMPANY_DOMAIN)) internal = true;
    else external = true;
  });

  return { external, payroll, internal };
}

// Decision engine
function analyzeMail(body, recipients) {

  if (!containsPAN(body)) return { action: "ALLOW" };

  if (containsMaskedPAN(body)) return { action: "ALLOW" };

  const target = classifyRecipients(recipients);

  if (target.payroll) return { action: "ALLOW" };

  if (target.external) {
    return {
      action: "BLOCK",
      message: "DLP Policy: PAN cannot be sent to external recipients."
    };
  }

  return {
    action: "WARN",
    message: "DLP Warning: You are sending PAN inside organization. Remove or justify."
  };
}

/*********************** EVENT HANDLER ************************/

function onMessageSendHandler(event) {

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
        if (recResult.status === Office.AsyncResultStatus.Succeeded) {
          recipients = recResult.value.map(r => r.emailAddress);
        }

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

        // WARN case → simulate override
        event.completed({
          allowEvent: false,
          errorMessage: decision.message + "\n\nEdit the mail or remove PAN to continue."
        });

      });
    }
  );
}

/*********************** READ HANDLER (disabled) ************************/

function onMessageReadHandler(event) {
  event.completed({ allowEvent: true });
}

/*********************** SAFE INITIALIZATION ************************/

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

if (typeof Office !== "undefined") {
  Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) registerEventHandlers();
  });
}

(function () {
  try {
    if (Office?.context?.mailbox) registerEventHandlers();
  } catch {}
})();