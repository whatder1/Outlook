Office.onReady();

// New Outlook embeds unaccepted text predictions directly in the email body as
// an HTML widget:
//
//   <span class="entityDelimiterBefore">​</span>
//   <span class="_Entity _EType_suggestedCompletion ...">
//     <span id="textPredictionParent" class="suggestedCompletion" ...>
//       ​‌<span id="suggestionText">zabeth</span><span id="textPredictionTabHint">Tab</span>​‌
//     </span>
//   </span>
//   <span class="entityDelimiterAfter">​</span>
//
// If the user hits Send without pressing Tab, the entire widget ships with the
// message. We strip it on send so only the user's actually-typed text remains.

// Matches the entity wrapper that contains the prediction. Non-greedy body with
// "</span>\s*</span>" terminator — the only place two </span> appear with only
// whitespace between is the wrapper's closing pair.
const SUGGESTED_COMPLETION =
  /<span\b[^>]*class="[^"]*_EType_suggestedCompletion[^"]*"[^>]*>[\s\S]*?<\/span>\s*<\/span>/gi;

// Surrounding zero-width delimiter spans Outlook adds before/after the widget.
const ENTITY_DELIMITER =
  /<span\b[^>]*class="[^"]*entityDelimiter(?:Before|After)[^"]*"[^>]*>[\s\S]*?<\/span>/gi;

// Backstop 1: raw "Eli<invisible>zabeth(Tab)" if the artifact shows up
// without the HTML wrapper.
const ZW = "[\\u200B-\\u200D\\uFEFF]";
const ARTIFACT_PATTERN = new RegExp(
  "Eli" + ZW + "+zabeth(?:(?:\\s|<[^>]+>|&nbsp;|&#160;)*Tab\\b)?",
  "g"
);

// Backstop 2: plain "Elizabeth" right after a sign-off.
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

        fixed = fixed.replace(SUGGESTED_COMPLETION, "");
        fixed = fixed.replace(ENTITY_DELIMITER, "");
        ARTIFACT_PATTERN.lastIndex = 0;
        fixed = fixed.replace(ARTIFACT_PATTERN, "Eli");
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
