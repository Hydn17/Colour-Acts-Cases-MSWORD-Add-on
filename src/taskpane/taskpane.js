/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Word.run(async (context) => {
      // Load the whole document body text
      const body = context.document.body;
      body.load("text");
      await context.sync();
      const text = body.text || "";

  // Regex to match patterns like:
  // "Lee v The Minister of Foreign Affairs [2003]" or "Mills v Harris (1975)"
  // - Party names: start with upper-case letter, allow words and common punctuation
  // - 'v' or 'v.' or 'versus' variants
  // - year in square brackets or parentheses [YYYY] or (YYYY)
  // add the 'i' flag so the regex is case-insensitive
  const citationRegex = /([A-Z][A-Za-z0-9'’\.\-,&() ]{1,160}?)\s+v(?:\.|ersus)?\s+([A-Z][A-Za-z0-9'’\.\-,&() ]{1,160}?)\s*[\[\(]([0-9]{4})[\]\)]/gi;

      const matches = [];
      let m;
      while ((m = citationRegex.exec(text)) !== null) {
        matches.push({
          fullMatch: m[0],
          claimant: m[1].trim(),
          respondent: m[2].trim(),
          year: m[3],
          index: m.index
        });
        // Avoid infinite loops for zero-length matches (not expected here)
        if (citationRegex.lastIndex === m.index) citationRegex.lastIndex++;
      }

      // For each match, find the exact occurrence(s) in the document and
      // colour the text before the date (i.e., the parties and 'v') red.
      // We do this by searching for the full match, then searching within
      // that found range for the pre-date substring and applying formatting
      // to that sub-range. This avoids changing other identical text elsewhere.
      for (const mm of matches) {
        try {
          // Derive the substring that precedes the date/parenthesis/bracket
          const preDate = mm.fullMatch.replace(/\s*[\[\(][0-9]{4}[\]\)]\s*$/u, "");

          if (!preDate) continue;

          // Search for the full match first (literal search). This returns RangeCollection
          // use case-insensitive search so matches are found regardless of capitalization
          const fullRanges = body.search(mm.fullMatch, {matchCase: false, matchWholeWord: false});
          fullRanges.load("items");
          await context.sync();

          // For each found full match, search within it for the preDate substring
          // and apply font colour to those sub-ranges only.
          for (const r of fullRanges.items) {
            // search the pre-date substring case-insensitively as well
            const subRanges = r.search(preDate, {matchCase: false, matchWholeWord: false});
            // Load the font properties for each sub-range so we can inspect italic
            subRanges.load("items/font");
            await context.sync();

            for (const sr of subRanges.items) {
              try {
                // If the sub-range is not italicised, set it to italic
                if (!sr.font.italic) {
                  sr.font.italic = true;
                }
                sr.font.color = "red";
              } catch (e) {
                // If setting properties fails for some reason, continue with others
                console.error("Failed to set color/italic for range:", e);
              }
            }
          }
          // sync once per match batch to commit font changes
          await context.sync();
        } catch (innerErr) {
          // Non-fatal per-match errors shouldn't stop the process
          console.error("Error processing match:", mm, innerErr);
        }
      }

          // --- New: match Acts like "The Interpretation Act (1971)" ---
          // Match patterns where the name ends with the word 'Act' followed by a year
          // in parentheses or square brackets. We'll colour the pre-date part blue.
          const actsRegex = /([A-Z][A-Za-z0-9'’\.\-,&() ]{1,200}?\bAct)\s*[\[\(]([0-9]{4})[\]\)]/gi;
          const actsMatches = [];
          while ((m = actsRegex.exec(text)) !== null) {
            actsMatches.push({
              fullMatch: m[0],
              preDate: m[1].trim(),
              year: m[2],
              index: m.index
            });
            if (actsRegex.lastIndex === m.index) actsRegex.lastIndex++;
          }

          for (const am of actsMatches) {
            try {
              // Search for the full match in the document (case-insensitive)
              const fullRanges = body.search(am.fullMatch, {matchCase: false, matchWholeWord: false});
              fullRanges.load("items");
              await context.sync();

              for (const r of fullRanges.items) {
                const subRanges = r.search(am.preDate, {matchCase: false, matchWholeWord: false});
                // Load the font properties for each sub-range so we can inspect italic
                subRanges.load("items/font");
                await context.sync();

                for (const sr of subRanges.items) {
                  try {
                    // If the sub-range is not italicised, set it to italic
                    if (!sr.font.italic) {
                      sr.font.italic = true;
                    }
                    sr.font.color = "blue";
                  } catch (e) {
                    console.error("Failed to set color/italic for act range:", e);
                  }
                }
              }
              await context.sync();
            } catch (e) {
              console.error("Error processing act match:", am, e);
            }
          }

      // Render results: look for an element with id="results" in the taskpane
      const outEl = document.getElementById("results");
      if (outEl) {
        if (matches.length === 0) {
          outEl.textContent = "No matches found.";
        } else {
          // Build a simple HTML list of matches
          outEl.innerHTML = matches.map((x, i) =>
            `<div class="match"><strong>#${i + 1}</strong>: ${escapeHtml(x.fullMatch)}<div class="meta">Claimant: ${escapeHtml(x.claimant)} | Respondent: ${escapeHtml(x.respondent)} | Year: ${x.year}</div></div>`
          ).join("");
        }
      } else {
        // If no results element is present, log to console
        console.log("Citation matches:", matches);
      }
    });
  } catch (err) {
    console.error("Error searching document:", err);
    const outEl = document.getElementById("results");
    if (outEl) outEl.textContent = `Error: ${err.message}`;
  }

  // Helper: escape HTML for safe display
  function escapeHtml(str) {
    return str.replace(/[&<>"']/g, (s) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[s]));
  }
}
