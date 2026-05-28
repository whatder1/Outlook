Office.onReady();

// Matches "Elizabeth" only when it appears right after a sign-off line
// (Best, / Thanks, / Regards, etc.) — keeps the add-in from renaming
// other people named Elizabeth that you're emailing about.
const SIGNATURE_PATTERN = /((?:Best|Thanks|Thank you|Regards|Best regards|Kind regards|Sincerely|Cheers|Warmly|Take care|Talk soon|All the best)[,!.]?\s*(?:<br\s*\/?>|<\/p>\s*<p[^>]*>|<\/div>\s*<div[^>]*>|\n|\r)+\s*(?:<[^>]+>\s*)*)Elizabeth\b/gi;

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Html, (getResult) => {
    if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
      // Couldn't read the body — let the send proceed rather than block.
      event.completed({ allowEvent: true });
      return;
    }

    const body = getResult.value;

    if (!SIGNATURE_PATTERN.test(body)) {
      event.completed({ allowEvent: true });
      return;
    }

    SIGNATURE_PATTERN.lastIndex = 0; // reset after .test()
    const fixedBody = body.replace(SIGNATURE_PATTERN, "$1Eli");

    item.body.setAsync(
      fixedBody,
      { coercionType: Office.CoercionType.Html },
      (setResult) => {
        if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed({
            allowEvent: false,
            errorMessage: "Couldn't fix the signature automatically. Edit the email and send again."
          });
          return;
        }
        event.completed({ allowEvent: true });
      }
    );
  });
}

// Register the handler so the manifest's FunctionName resolves.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
