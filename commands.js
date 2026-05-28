Office.onReady();

// Zero-width / invisible separator characters that New Outlook's text prediction
// inserts between the user-typed prefix ("Eli") and the suggested suffix
// ("zabeth"). Includes ZWSP, ZWNJ, ZWJ, and BOM.
const ZW = "[\\u200B-\\u200D\\uFEFF]";

// Pattern 1 — the unambiguous autocomplete artifact: "Eli" + one-or-more
// invisible chars + "zabeth". No one types this on purpose, so it's always
// safe to collapse it back to "Eli".
// Also strip a trailing "Tab" token if the prediction overlay's keyboard hint
// somehow leaked into the body (e.g., via copy-paste from the compose window).
const ARTIFACT_PATTERN = new RegExp(
  "Eli" + ZW + "+zabeth(?:(?:\\s|<[^>]+>|&nbsp;|&#160;)*Tab\\b)?",
  "g"
);

// Pattern 2 — plain "Elizabeth" right after a sign-off, as a backstop in case
// the prediction lands without invisible separators in some Outlook version.
const SIGN_OFFS =
  "Best(?:\\s+regards)?|Thanks(?:\\s+again)?|Thank\\s+you|Regards|" +
  "Kind\\s+regards|Sincerely|Cheers|Warmly|Take\\s+care|Talk\\s+soon|" +
  "All\\s+the\\s+best";
const SIGNOFF_PATTERN = new RegExp(
  "(\\b(?:" + SIGN_OFFS + ")[,!.]?)" +
    "((?:\\s|<[^>]+>|&nbsp;|&#160;|&[a-zA-Z]+;){1,500}?)" +
    "Elizabeth(?:(?:\\s|<[^>]+>|&nbsp;|&#160;)*Tab)?\\b",
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
        let fixed = body;

        // First pass: zap "Eli<invisible>zabeth" artifacts anywhere.
        ARTIFACT_PATTERN.lastIndex = 0;
        fixed = fixed.replace(ARTIFACT_PATTERN, "Eli");

        // Second pass: clean plain "Elizabeth" near a sign-off (backstop).
        SIGNOFF_PATTERN.lastIndex = 0;
        fixed = fixed.replace(SIGNOFF_PATTERN, "$1$2Eli");

        if (fixed === body) {
          event.completed({ allowEvent: true });
          return;
        }

        item.body.setAsync(
          fixed,
          { coercionType: Office.CoercionType.Html },
          function () {
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
