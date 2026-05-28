Office.onReady();

// === DIAGNOSTIC BUILD v1.0.0.5 ===
// This build prepends a visible marker "[SIG-FIX v5 RAN]" to every outgoing
// email body so we can confirm the OnMessageSend handler is actually firing.
// If you send a test and the recipient sees the marker, the add-in is loaded
// and running — and the prior versions' regex was the problem. If the marker
// is missing, the add-in itself is not being invoked (cache / install issue).

const MARKER = "[SIG-FIX v5 RAN] ";

// Matches the entire entity wrapper that contains the prediction.
const SUGGESTED_COMPLETION =
  /<span\b[^>]*class="[^"]*_EType_suggestedCompletion[^"]*"[^>]*>[\s\S]*?<\/span>\s*<\/span>/gi;

const ENTITY_DELIMITER =
  /<span\b[^>]*class="[^"]*entityDelimiter(?:Before|After)[^"]*"[^>]*>[\s\S]*?<\/span>/gi;

const ZW = "[\\u200B-\\u200D\\uFEFF]";
const ARTIFACT_PATTERN = new RegExp(
  "Eli" + ZW + "+zabeth(?:(?:\\s|<[^>]+>|&nbsp;|&#160;)*Tab\\b)?",
  "g"
);

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

        // Strip the prediction widget if present.
        fixed = fixed.replace(SUGGESTED_COMPLETION, "");
        fixed = fixed.replace(ENTITY_DELIMITER, "");
        ARTIFACT_PATTERN.lastIndex = 0;
        fixed = fixed.replace(ARTIFACT_PATTERN, "Eli");
        SIGNOFF_PATTERN.lastIndex = 0;
        fixed = fixed.replace(SIGNOFF_PATTERN, "$1$2Eli");

        // Always prepend a visible marker so we know the handler ran.
        // Inject at the very top of the body so it's the first thing
        // recipients see.
        const bodyOpen = fixed.search(/<body\b[^>]*>/i);
        if (bodyOpen >= 0) {
          const tagEnd = fixed.indexOf(">", bodyOpen) + 1;
          fixed =
            fixed.slice(0, tagEnd) +
            "<div style='color:red;font-weight:bold'>" +
            MARKER +
            "</div>" +
            fixed.slice(tagEnd);
        } else {
          fixed = MARKER + fixed;
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
