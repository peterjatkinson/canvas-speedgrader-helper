// content.js — Injected into the SpeedGrader page.
// Inspects the DOM to find the current student's name and returns it.
//
// Canvas updates its UI periodically, so these selectors may need adjustment.
// The script tries multiple strategies and reports what it found so we can adapt.

(() => {
  const result = {
    url: window.location.href,
    studentName: null,
    method: null           // which detection strategy worked
  };

  // Strategy 1 (confirmed working on Imperial Canvas, Apr 2026):
  // The newer React-based SpeedGrader has a button with
  // data-testid="student-select-trigger". Inside it, the student name
  // is in a <span> that is NOT the avatar-initials span.
  const trigger = document.querySelector('[data-testid="student-select-trigger"]');
  if (trigger) {
    // Get all leaf text spans, skip avatar initials and status indicator dots.
    // Some students have a ● (U+25CF) status dot between the avatar and name.
    const spans = trigger.querySelectorAll("span");
    for (const span of spans) {
      const text = span.textContent.trim();
      if (span.children.length === 0 &&
          !span.className.includes("avatar") &&
          text.length > 1 &&
          /[a-zA-Z0-9]/.test(text)) {
        result.studentName = text;
        result.method = "student-select-trigger";
        break;
      }
    }
  }

  // Strategy 2: Older Canvas — <select id="students_selectmenu">
  if (!result.studentName) {
    const studentSelect = document.getElementById("students_selectmenu");
    if (studentSelect) {
      const selectedOption = studentSelect.options[studentSelect.selectedIndex];
      if (selectedOption) {
        result.studentName = selectedOption.textContent.trim();
        result.method = "students_selectmenu";
      }
    }
  }

  // Strategy 3: Other known selectors from various Canvas versions
  if (!result.studentName) {
    const nameEl =
      document.querySelector(".ui-selectmenu-status .ui-selectmenu-item-header") ||
      document.querySelector("#combo_box_container .ui-selectmenu-status");
    if (nameEl) {
      result.studentName = nameEl.textContent.trim();
      result.method = "header_element";
    }
  }

  return result;
})();
