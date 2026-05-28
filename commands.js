Office.onReady();

// Sign-off words. Match "Elizabeth" only when it appears shortly after one of
// these, so we don't rename an actual person named Elizabeth that you're emailing.
const SIGN_OFFS =
  "Best(?:\\s+regards)?|Thanks(?:\\s+again)?|Thank\\s+you|Regards|" +
  "Kind\\s+regards|Sincerely|Cheers|Warmly|Take\\s+care|Talk\\s+soon|All\\s+the\\s+best";

// Group 1 = the sign-off (preserved). Group 2 = everything between the sign-off
// and "Elizabeth" — any mix of whitespace, HTML tags, nbsp, or entities, up to
// ~500 chars (non-greedy). This handles nested <div><font><span> wrappers that
// New Outlook produces between paragraphs.
const PATTERN = new RegExp(
  "(\\b(?:" + SIGN_OFFS + ")[,!.]?)" +
    "((?:\\s|<[^>]+>|&nbsp;|&#160;|&[a-zA-Z]+;){1,500}?)" +
    "Elizabeth\\b",
  "gi"
);

function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    item.body.getAsync(Office.CoercionType.Html, function (getResult) {
      try {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed({ allowEvent: true });
          return;
        }

        const body = getResult.value || "";

        // Cheap pre-check — bail fast if no "Elizabeth" anywhere.
        if (!/\bElizabeth\b/i.test(body)) {
          event.completed({ allowEvent: true });
          return;
        }

        PATTERN.lastIndex = 0;
        const fixed = body.replace(PATTERN, "$1$2Eli");

        if (fixed === body) {
          // "Elizabeth" appeared but not after a sign-off — leave it alone.
          event.completed({ allowEvent: true });
          return;
        }

        item.body.setAsync(
          fixed,
          { coercionType: Office.CoercionType.Html },
          function () {
            // Whether setAsync succeeded or not, allow the send through.
            event.completed({ allowEvent: true });
          }
        );
      } catch (innerErr) {
        event.completed({ allowEvent: true });
      }
    });
  } catch (outerErr) {
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
