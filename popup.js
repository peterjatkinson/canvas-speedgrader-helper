// popup.js — Extension popup logic.
// Phase 1: Detect current student from SpeedGrader DOM.
// Phase 2: Load CSV, match student, preview comment.
// Phase 3: Fill, submit, next student, batch processing, logging.

// ── CSV parsing ────────────────────────────────────────────────

function parseCSV(text) {
  const lines = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { inQuotes = !inQuotes; current += ch; }
    else if (ch === "\n" && !inQuotes) { lines.push(current.trim()); current = ""; }
    else if (ch === "\r" && !inQuotes) { /* skip */ }
    else { current += ch; }
  }
  if (current.trim()) lines.push(current.trim());
  if (lines.length < 2) return [];
  const headers = splitCSVLine(lines[0]).map(h => h.toLowerCase().trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i]) continue;
    const values = splitCSVLine(lines[i]);
    const row = {};
    headers.forEach((h, idx) => { row[h] = (values[idx] || "").trim(); });
    rows.push(row);
  }
  return rows;
}

function splitCSVLine(line) {
  const fields = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
      else { inQuotes = !inQuotes; }
    } else if (ch === "," && !inQuotes) { fields.push(current); current = ""; }
    else { current += ch; }
  }
  fields.push(current);
  return fields;
}

// ── Student matching ───────────────────────────────────────────

function normalizeName(name) {
  return name.normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/\s+/g, " ").trim();
}

// Extract the ID token from a CSV id value like "anon_98DDB_5726" or "LATE_anon_xBzzu_text".
// Returns Group 2 from the regex, e.g. "98DDB", or null if no match.
function extractCsvId(value) {
  const match = value.match(/^(?:.+_)?([a-zA-Z0-9]+)_(text|link|\d+)(?=_|$)/);
  return match ? match[1] : null;
}

// Extract the last =value from a SpeedGrader URL.
// e.g. "...?assignment_id=431&anonymous_id=98DDB" → "98DDB"
function extractUrlId(url) {
  const params = url.split("?")[1];
  if (!params) return null;
  const pairs = params.split("&");
  const last = pairs[pairs.length - 1];
  const eqIdx = last.indexOf("=");
  return eqIdx >= 0 ? last.substring(eqIdx + 1) : null;
}

// Display label for a student: prefer the URL ID, fall back to name
function studentLabel(name, url) {
  const id = url ? extractUrlId(url) : null;
  return id ? id : name;
}

function findMatch(studentName, csvRows, pageUrl) {
  if (!csvRows.length) return null;
  const headers = Object.keys(csvRows[0]);

  // Identify ID columns and name columns separately
  const idColumns = headers.filter(h => /^(id|student.?id)$/i.test(h));
  const nameColumns = headers.filter(h =>
    /^(name|student|student.?name|full.?name|learner|display.?name)$/i.test(h)
  );

  // Strategy 1: ID-based matching (CSV id column → URL parameter)
  if (idColumns.length > 0 && pageUrl) {
    const urlId = extractUrlId(pageUrl);
    if (urlId) {
      for (const row of csvRows) {
        for (const col of idColumns) {
          const csvId = extractCsvId(row[col] || "");
          if (csvId && csvId === urlId) {
            return { row, matchedOn: col + " (id: " + csvId + ")" };
          }
        }
      }
    }
  }

  // Strategy 2: Name-based matching (existing logic)
  const nameNorm = normalizeName(studentName);
  const columnsToSearch = nameColumns.length > 0 ? nameColumns : headers;

  for (const row of csvRows) {
    for (const col of columnsToSearch) {
      const val = normalizeName(row[col] || "");
      if (!val) continue;
      if (val === nameNorm) return { row, matchedOn: col };
      if (val.includes(nameNorm) || nameNorm.includes(val)) return { row, matchedOn: col };
    }
  }
  const nameParts = nameNorm.split(/[\s,]+/).filter(Boolean);
  if (nameParts.length >= 2) {
    const reversed = nameParts.reverse().join(" ");
    for (const row of csvRows) {
      for (const col of columnsToSearch) {
        const val = normalizeName(row[col] || "");
        if (val === reversed || val.includes(reversed) || reversed.includes(val)) {
          return { row, matchedOn: col + " (reversed)" };
        }
      }
    }
  }
  return null;
}

function findCommentColumn(headers) {
  const commentPatterns = [
    /^(comment|comments|feedback|feedback.?comment|response|note|notes)$/i
  ];
  for (const h of headers) {
    for (const pat of commentPatterns) {
      if (pat.test(h)) return h;
    }
  }
  return null;
}

// ── Storage helpers ────────────────────────────────────────────

async function saveCSV(rows, filename) {
  await chrome.storage.local.set({ csvRows: rows, csvFilename: filename });
}
async function loadCSV() {
  const data = await chrome.storage.local.get(["csvRows", "csvFilename"]);
  return { rows: data.csvRows || null, filename: data.csvFilename || null };
}
async function clearCSVStorage() {
  await chrome.storage.local.remove(["csvRows", "csvFilename"]);
}

// ── Page interaction functions (injected into SpeedGrader) ─────

// Find a SpeedGrader tab fresh (ignoring any cached ID).
// Prefers the active tab if it's SpeedGrader, otherwise searches all tabs.
async function findSpeedGraderTab() {
  // Prefer the active tab — this is the one the user is looking at
  const [active] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (active && active.url && (active.url.includes("speed_grader") || active.url.includes("speedgrader"))) {
    speedGraderTabId = active.id;
    return active;
  }

  // Fall back to searching all tabs in the current window
  const tabs = await chrome.tabs.query({ currentWindow: true });
  const sgTab = tabs.find(t => t.url && (t.url.includes("speed_grader") || t.url.includes("speedgrader")));
  if (sgTab) {
    speedGraderTabId = sgTab.id;
    return sgTab;
  }

  throw new Error("No SpeedGrader tab found. Open SpeedGrader first.");
}

// Get the SpeedGrader tab, reusing the cached ID if valid.
// Used mid-batch so it keeps targeting the same tab even if user switches away.
async function getSpeedGraderTab() {
  if (speedGraderTabId !== null) {
    try {
      const tab = await chrome.tabs.get(speedGraderTabId);
      if (tab && tab.url && (tab.url.includes("speed_grader") || tab.url.includes("speedgrader"))) {
        return tab;
      }
    } catch (e) {
      // Tab was closed or doesn't exist anymore
    }
    speedGraderTabId = null;
  }
  // Cache expired or invalid — discover fresh
  return findSpeedGraderTab();
}

async function detectStudent() {
  const tab = await getSpeedGraderTab();
  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ["content.js"]
  });
  const data = results[0]?.result;
  if (!data) throw new Error("Script ran but returned no data.");
  return data;
}

// Fill the GENERAL comment box (not rubric) via TinyMCE.
// All helper functions are defined INSIDE this function because it gets
// injected into the page via chrome.scripting.executeScript and cannot
// access anything else defined in popup.js.
function fillCommentOnPage(commentText) {
  // ── Find the correct editor ──
  function findEditor() {
    if (typeof tinymce === "undefined" || !tinymce.editors.length) return null;
    const submitBtn = document.getElementById("comment_submit_button");
    if (submitBtn) {
      let container = submitBtn.parentElement;
      while (container) {
        const iframe = container.querySelector('iframe[id$="_ifr"]');
        if (iframe) {
          return tinymce.get(iframe.id.replace("_ifr", ""));
        }
        container = container.parentElement;
      }
    }
    for (const ed of tinymce.editors) {
      const iframe = document.getElementById(ed.id + "_ifr");
      if (iframe && !iframe.closest("table")) return ed;
    }
    return tinymce.editors[0];
  }

  // ── Inline formatting: bold, italic ──
  function inlineFormat(text) {
    text = text.replace(/\*\*\*(.+?)\*\*\*/g, "<strong><em>$1</em></strong>");
    text = text.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
    text = text.replace(/__(.+?)__/g, "<strong>$1</strong>");
    text = text.replace(/\*(.+?)\*/g, "<em>$1</em>");
    text = text.replace(/(?<!\w)_(.+?)_(?!\w)/g, "<em>$1</em>");
    return text;
  }

  // ── Markdown to HTML ──
  function markdownToHtml(md) {
    let html = "";
    const lines = md.split("\n");
    let inUl = false;
    let inOl = false;

    for (const rawLine of lines) {
      const trimmed = rawLine.trim();

      if (/^---+$/.test(trimmed) || /^\*\*\*+$/.test(trimmed)) {
        if (inUl) { html += "</ul>"; inUl = false; }
        if (inOl) { html += "</ol>"; inOl = false; }
        html += "<hr>";
        continue;
      }
      const headingMatch = trimmed.match(/^(#{1,6})\s+(.+)/);
      if (headingMatch) {
        if (inUl) { html += "</ul>"; inUl = false; }
        if (inOl) { html += "</ol>"; inOl = false; }
        const level = headingMatch[1].length;
        html += "<h" + level + ">" + inlineFormat(headingMatch[2]) + "</h" + level + ">";
        continue;
      }
      const ulMatch = trimmed.match(/^[-*]\s+(.*)/);
      if (ulMatch) {
        if (inOl) { html += "</ol>"; inOl = false; }
        if (!inUl) { html += "<ul>"; inUl = true; }
        html += "<li>" + inlineFormat(ulMatch[1]) + "</li>";
        continue;
      }
      const olMatch = trimmed.match(/^\d+\.\s+(.*)/);
      if (olMatch) {
        if (inUl) { html += "</ul>"; inUl = false; }
        if (!inOl) { html += "<ol>"; inOl = true; }
        html += "<li>" + inlineFormat(olMatch[1]) + "</li>";
        continue;
      }
      if (trimmed === "") {
        if (inUl) { html += "</ul>"; inUl = false; }
        if (inOl) { html += "</ol>"; inOl = false; }
        continue;
      }
      if (inUl) { html += "</ul>"; inUl = false; }
      if (inOl) { html += "</ol>"; inOl = false; }
      html += "<p>" + inlineFormat(trimmed) + "</p>";
    }
    if (inUl) html += "</ul>";
    if (inOl) html += "</ol>";
    return html;
  }

  // ── Main logic ──
  const editor = findEditor();
  if (!editor) return { success: false, error: "Could not find general comment editor." };

  const html = markdownToHtml(commentText);
  editor.setContent(html);
  editor.fire("change");
  return { success: true };
}

async function fillComment(commentText) {
  const tab = await getSpeedGraderTab();
  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id }, world: "MAIN",
    func: fillCommentOnPage, args: [commentText]
  });
  const data = results[0]?.result;
  if (!data) throw new Error("Fill script returned no data.");
  if (!data.success) throw new Error(data.error);
}

// Click the "Submit comment" button, wait for it to post, then clear the editor.
// Helper functions are defined inside because this runs on the page context.
function submitAndClearOnPage() {
  function findEditor() {
    if (typeof tinymce === "undefined" || !tinymce.editors.length) return null;
    const submitBtn = document.getElementById("comment_submit_button");
    if (submitBtn) {
      let container = submitBtn.parentElement;
      while (container) {
        const iframe = container.querySelector('iframe[id$="_ifr"]');
        if (iframe) {
          return tinymce.get(iframe.id.replace("_ifr", ""));
        }
        container = container.parentElement;
      }
    }
    for (const ed of tinymce.editors) {
      const iframe = document.getElementById(ed.id + "_ifr");
      if (iframe && !iframe.closest("table")) return ed;
    }
    return tinymce.editors[0];
  }

  // Count all comments (submitted + draft) by counting "Delete comment" buttons.
  function countAllComments() {
    const btns = document.querySelectorAll("button");
    let count = 0;
    for (const b of btns) {
      if (b.textContent.includes("Delete comment")) count++;
    }
    return count;
  }

  // Count draft comments by looking for the "Draft" pill (data-testid="draft-pill").
  function countDraftComments() {
    return document.querySelectorAll('[data-testid="draft-pill"]').length;
  }

  return new Promise((resolve) => {
    const btn = document.getElementById("comment_submit_button");
    if (!btn) { resolve({ success: false, error: "Submit button not found." }); return; }
    if (btn.disabled) { resolve({ success: false, error: "Submit button is disabled." }); return; }

    // Snapshot counts BEFORE clicking submit
    const totalBefore = countAllComments();
    const draftsBefore = countDraftComments();

    btn.click();

    // Poll until we confirm the comment was truly submitted (not just drafted):
    //   - Total comments (Delete buttons) must increase by 1
    //   - Draft count must NOT increase (if it did, it was saved as draft, not posted)
    // Check every 500ms. Two timeouts:
    //   - 5s after a draft is detected (it's not going to transition to posted)
    //   - 15s total if nothing appears at all
    let elapsed = 0;
    let draftDetectedAt = null;
    const interval = setInterval(() => {
      elapsed += 500;
      const totalNow = countAllComments();
      const draftsNow = countDraftComments();

      const newCommentAppeared = totalNow > totalBefore;
      const noDraftAdded = draftsNow <= draftsBefore;

      if (newCommentAppeared && noDraftAdded) {
        // Success: comment posted, not a draft
        clearInterval(interval);
        const editor = findEditor();
        if (editor) { editor.setContent(""); editor.fire("change"); }
        resolve({ success: true });

      } else if (newCommentAppeared && !noDraftAdded) {
        // Comment appeared but as a draft
        if (!draftDetectedAt) draftDetectedAt = elapsed;
        // Give it 5s from when draft was first detected to transition to posted
        if (elapsed - draftDetectedAt >= 5000) {
          clearInterval(interval);
          const editor = findEditor();
          if (editor) { editor.setContent(""); editor.fire("change"); }
          resolve({ success: false, error: "Comment saved as draft but not posted." });
        }

      } else if (elapsed >= 15000) {
        // Nothing appeared at all
        clearInterval(interval);
        const editor = findEditor();
        if (editor) { editor.setContent(""); editor.fire("change"); }
        resolve({ success: false, error: "Timed out waiting for comment to appear (15s)." });
      }
    }, 500);
  });
}

async function submitComment() {
  const tab = await getSpeedGraderTab();
  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id }, world: "MAIN",
    func: submitAndClearOnPage
  });
  const data = results[0]?.result;
  if (!data) throw new Error("Submit script returned no data.");
  if (!data.success) throw new Error(data.error);
}

// Click the "Next student" button
function nextStudentOnPage() {
  const btn = document.getElementById("next-student-button");
  if (!btn) return { success: false, error: "Next student button not found." };
  if (btn.disabled) return { success: false, error: "No next student (end of list)." };
  btn.click();
  return { success: true };
}

async function goNextStudent() {
  const tab = await getSpeedGraderTab();
  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id }, world: "MAIN",
    func: nextStudentOnPage
  });
  const data = results[0]?.result;
  if (!data) throw new Error("Next-student script returned no data.");
  if (!data.success) throw new Error(data.error);
}

// ── Logging ────────────────────────────────────────────────────

const logEl = document.getElementById("log");
const logSection = document.getElementById("logSection");

function log(message, type = "ok") {
  logSection.classList.remove("hidden");
  const line = document.createElement("div");
  line.className = type;
  const time = new Date().toLocaleTimeString();
  line.textContent = `[${time}] ${message}`;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

// ── Helpers ────────────────────────────────────────────────────

function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

async function detectStudentWithRetry(retries = 3, delayMs = 1500) {
  for (let i = 0; i < retries; i++) {
    if (i > 0) await wait(delayMs);
    try {
      const data = await detectStudent();
      if (data.studentName) return data;
    } catch (e) {
      if (i === retries - 1) throw e;
    }
  }
  throw new Error("Could not detect student after retries.");
}

function getCommentForStudent(studentName, pageUrl) {
  if (!csvRows) return null;
  const match = findMatch(studentName, csvRows, pageUrl);
  if (!match) return null;
  const headers = Object.keys(match.row);
  const commentCol = findCommentColumn(headers);
  if (!commentCol) return null;
  return match.row[commentCol];
}

// ── UI Elements ───────────────────────────────────────────────

const csvInfoEl = document.getElementById("csvInfo");
const csvFileInput = document.getElementById("csvFile");
const statusEl = document.getElementById("status");
const detailsEl = document.getElementById("details");
const matchEl = document.getElementById("matchResult");
const fillBtn = document.getElementById("fillBtn");
const submitBtn = document.getElementById("submitBtn");
const nextBtn = document.getElementById("nextBtn");
const fillStatusEl = document.getElementById("fillStatus");
const batchBtn = document.getElementById("batchBtn");
const pauseBtn = document.getElementById("pauseBtn");

// Sections
const modeSection = document.getElementById("modeSection");
const oneByOneSection = document.getElementById("oneByOneSection");
const batchSection = document.getElementById("batchSection");
const matchSection = document.getElementById("matchSection");
const batchProgress = document.getElementById("batchProgress");
const batchCurrentItem = document.getElementById("batchCurrentItem");
const batchSummary = document.getElementById("batchSummary");

// Mode buttons
const modeOneBtn = document.getElementById("modeOneBtn");
const modeBatchBtn = document.getElementById("modeBatchBtn");

let currentStudentName = null;
let currentUrl = null;
let currentComment = null;
let csvRows = null;
let batchRunning = false;
let batchPaused = false;
let speedGraderTabId = null;

// ── Show / hide helpers ───────────────────────────────────────

function showSection(el) { el.classList.remove("hidden"); }
function hideSection(el) { el.classList.add("hidden"); }

function selectMode(mode) {
  if (mode === "one") {
    modeOneBtn.classList.add("active");
    modeBatchBtn.classList.remove("active");
    showSection(oneByOneSection);
    hideSection(batchSection);
  } else {
    modeBatchBtn.classList.add("active");
    modeOneBtn.classList.remove("active");
    showSection(batchSection);
    hideSection(oneByOneSection);
  }
}

// ── On load: restore CSV ──────────────────────────────────────

(async () => {
  const saved = await loadCSV();
  if (saved.rows && saved.rows.length > 0) {
    csvRows = saved.rows;
    const headers = Object.keys(csvRows[0]);
    csvInfoEl.textContent = `Loaded: ${saved.filename} (${csvRows.length} rows, columns: ${headers.join(", ")})`;
    document.getElementById("clearBtn").disabled = false;
    showSection(modeSection);
  }
})();

// ── Mode buttons ──────────────────────────────────────────────

modeOneBtn.addEventListener("click", () => selectMode("one"));
modeBatchBtn.addEventListener("click", () => selectMode("batch"));

// ── Detect button (one-at-a-time) ─────────────────────────────

document.getElementById("detectBtn").addEventListener("click", async () => {
  speedGraderTabId = null;  // always re-discover on user action
  statusEl.textContent = "Detecting...";
  statusEl.className = "status-box";
  detailsEl.textContent = "";
  try {
    const data = await detectStudent();
    if (data.studentName) {
      currentStudentName = data.studentName;
      currentUrl = data.url;
      statusEl.textContent = studentLabel(data.studentName, data.url) + " — " + data.studentName;
      statusEl.className = "status-box success";
      tryMatch();
    } else {
      statusEl.textContent = "On SpeedGrader, but could not find student name.";
      statusEl.className = "status-box error";
    }
  } catch (err) {
    statusEl.textContent = err.message;
    statusEl.className = "status-box error";
  }
});

// ── CSV upload ─────────────────────────────────────────────────

document.getElementById("uploadBtn").addEventListener("click", () => { csvFileInput.click(); });

csvFileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async (ev) => {
    const rows = parseCSV(ev.target.result);
    if (rows.length === 0) { csvInfoEl.textContent = "Error: CSV has no data rows."; return; }
    csvRows = rows;
    const headers = Object.keys(csvRows[0]);
    csvInfoEl.textContent = `Loaded: ${file.name} (${csvRows.length} rows, columns: ${headers.join(", ")})`;
    document.getElementById("clearBtn").disabled = false;
    showSection(modeSection);
    await saveCSV(csvRows, file.name);
    tryMatch();
  };
  reader.readAsText(file);
});

document.getElementById("clearBtn").addEventListener("click", async () => {
  csvRows = null;
  currentStudentName = null;
  currentUrl = null;
  currentComment = null;
  csvInfoEl.textContent = "No CSV loaded. Upload a CSV with an ID or name column and a comment column.";
  fillBtn.disabled = true;
  submitBtn.disabled = true;
  nextBtn.disabled = true;
  fillStatusEl.textContent = "";
  document.getElementById("clearBtn").disabled = true;
  hideSection(modeSection);
  hideSection(oneByOneSection);
  hideSection(batchSection);
  hideSection(logSection);
  modeOneBtn.classList.remove("active");
  modeBatchBtn.classList.remove("active");
  await clearCSVStorage();
});

// ── Step-by-step buttons ───────────────────────────────────────

fillBtn.addEventListener("click", async () => {
  if (!currentComment) return;
  fillStatusEl.textContent = "Filling...";
  fillStatusEl.style.color = "#555";
  submitBtn.disabled = true;
  try {
    await fillComment(currentComment);
    fillStatusEl.textContent = "Filled. Review in SpeedGrader, then Submit.";
    fillStatusEl.style.color = "green";
    submitBtn.disabled = false;
  } catch (err) {
    fillStatusEl.textContent = "Error: " + err.message;
    fillStatusEl.style.color = "red";
  }
});

submitBtn.addEventListener("click", async () => {
  fillStatusEl.textContent = "Submitting...";
  fillStatusEl.style.color = "#555";
  try {
    await submitComment();
    fillStatusEl.textContent = "Submitted! Click Next Student to continue.";
    fillStatusEl.style.color = "green";
    submitBtn.disabled = true;
    nextBtn.disabled = false;
    log(`Submitted: ${studentLabel(currentStudentName, currentUrl)} (${currentStudentName})`);
  } catch (err) {
    fillStatusEl.textContent = "Submit error: " + err.message;
    fillStatusEl.style.color = "red";
    log(`Submit failed for ${studentLabel(currentStudentName, currentUrl)} (${currentStudentName}): ${err.message}`, "err");
  }
});

nextBtn.addEventListener("click", async () => {
  fillStatusEl.textContent = "Moving to next student...";
  fillStatusEl.style.color = "#555";
  try {
    await goNextStudent();
    nextBtn.disabled = true;
    await wait(2000);
    const data = await detectStudentWithRetry();
    currentStudentName = data.studentName;
    currentUrl = data.url;
    statusEl.textContent = studentLabel(data.studentName, data.url) + " — " + data.studentName;
    statusEl.className = "status-box success";
    fillStatusEl.textContent = "";
    tryMatch();
  } catch (err) {
    fillStatusEl.textContent = "Next error: " + err.message;
    fillStatusEl.style.color = "red";
    log(`Next student failed: ${err.message}`, "err");
  }
});

// ── Batch processing ───────────────────────────────────────────

batchBtn.addEventListener("click", async () => {
  if (!csvRows) return;
  speedGraderTabId = null;  // always re-discover on user action
  batchRunning = true;
  batchPaused = false;
  batchBtn.classList.add("hidden");
  showSection(batchProgress);
  hideSection(batchSummary);
  pauseBtn.classList.remove("hidden");

  let processed = 0;
  let skipped = 0;
  let failed = 0;
  let drafts = 0;
  const unmatched = [];
  const seen = new Set();
  let firstStudent = null;

  log("Batch started");

  while (batchRunning) {
    if (batchPaused) {
      batchCurrentItem.textContent = `Paused. ${processed} done, ${skipped} skipped, ${drafts} draft, ${failed} failed.`;
      await wait(500);
      continue;
    }

    // 1. Detect current student
    let studentName, studentUrl;
    try {
      batchCurrentItem.innerHTML = '<span class="spinner"></span> Detecting student...';
      const data = await detectStudentWithRetry();
      studentName = data.studentName;
      studentUrl = data.url;
      if (!studentName) throw new Error("No name detected");
    } catch (err) {
      log(`Detection failed: ${err.message}`, "err");
      failed++;
      batchRunning = false;
      break;
    }

    // Loop detection
    const loopKey = studentUrl || studentName;
    if (!firstStudent) {
      firstStudent = loopKey;
    } else if (seen.has(loopKey)) {
      const label = studentLabel(studentName, studentUrl);
      log(`Back to ${label} — full loop complete, stopping.`);
      batchRunning = false;
      break;
    }
    seen.add(loopKey);

    const label = studentLabel(studentName, studentUrl);
    batchCurrentItem.innerHTML = `<span class="spinner"></span> Processing: ${label} (${studentName})`;

    // 2. Match against CSV
    const comment = getCommentForStudent(studentName, studentUrl);
    if (!comment) {
      log(`No CSV match: ${label} (${studentName})`, "skip");
      skipped++;
      unmatched.push(`${label} (${studentName})`);
    } else {
      // 3. Fill + submit
      try {
        await fillComment(comment);
        await wait(500);
        await submitComment();
        log(`Submitted: ${label} (${studentName})`);
        processed++;
      } catch (err) {
        if (err.message && err.message.includes("draft")) {
          log(`Draft only: ${label} (${studentName}) — comment saved as draft, not posted.`, "skip");
          drafts++;
          unmatched.push(`${label} (${studentName}) [draft]`);
        } else {
          log(`Failed for ${label} (${studentName}): ${err.message}`, "err");
          failed++;
        }
      }
    }

    // 4. Move to next student
    try {
      await goNextStudent();
    } catch (err) {
      log("Reached end of student list (or next-student failed).");
      batchRunning = false;
      break;
    }

    // 5. Wait for SpeedGrader to load
    await wait(2500);
  }

  // Summary
  batchRunning = false;
  hideSection(batchProgress);
  pauseBtn.classList.add("hidden");
  batchBtn.classList.remove("hidden");

  const total = processed + skipped + drafts + failed;
  let summaryHtml = '<div class="summary-box">';
  summaryHtml += `<div><span class="stat">${total}</span> students processed</div>`;
  summaryHtml += `<div><span class="stat good">${processed}</span> submitted</div>`;
  if (skipped > 0) summaryHtml += `<div><span class="stat warn">${skipped}</span> skipped (no CSV match)</div>`;
  if (drafts > 0) summaryHtml += `<div><span class="stat warn">${drafts}</span> saved as draft only</div>`;
  if (failed > 0) summaryHtml += `<div><span class="stat bad">${failed}</span> failed</div>`;
  summaryHtml += '</div>';

  batchSummary.innerHTML = summaryHtml;
  showSection(batchSummary);

  const summaryText = `Batch done: ${processed} submitted, ${skipped} skipped, ${drafts} draft only, ${failed} failed.`;
  log(summaryText);
  if (unmatched.length > 0) {
    log(`Unmatched: ${unmatched.join(", ")}`, "skip");
  }

  // Re-detect current student after batch
  try {
    const data = await detectStudentWithRetry();
    currentStudentName = data.studentName;
    currentUrl = data.url;
  } catch (e) { /* ignore */ }
});

pauseBtn.addEventListener("click", () => {
  if (batchPaused) {
    batchPaused = false;
    pauseBtn.textContent = "Pause";
    log("Batch resumed");
  } else {
    batchPaused = true;
    pauseBtn.textContent = "Resume";
    log("Batch paused");
  }
});

// ── Clear log ──────────────────────────────────────────────────

document.getElementById("clearLogBtn").addEventListener("click", () => {
  logEl.innerHTML = "";
});

// ── Match logic ────────────────────────────────────────────────

function tryMatch() {
  currentComment = null;
  fillBtn.disabled = true;
  submitBtn.disabled = true;
  nextBtn.disabled = !currentStudentName;  // always allow Next if a student is detected
  fillStatusEl.textContent = "";

  if (!currentStudentName || !csvRows) {
    if (currentStudentName && !csvRows) {
      matchEl.textContent = "Load a CSV to match against " + studentLabel(currentStudentName, currentUrl) + " (" + currentStudentName + ").";
      matchEl.className = "status-box warn";
    }
    return;
  }

  const match = findMatch(currentStudentName, csvRows, currentUrl);
  if (!match) {
    matchEl.textContent = "No match found for " + studentLabel(currentStudentName, currentUrl) + " (" + currentStudentName + ") in CSV.";
    matchEl.className = "status-box error";
    showSection(matchSection);
    return;
  }

  const headers = Object.keys(match.row);
  const commentCol = findCommentColumn(headers);
  if (commentCol) {
    const comment = match.row[commentCol];
    matchEl.textContent = "Matched on: " + match.matchedOn + "\nComment column: " + commentCol + "\n\n" + comment;
    matchEl.className = "status-box success";
    currentComment = comment;
    fillBtn.disabled = false;
    showSection(matchSection);
  } else {
    matchEl.textContent = "Matched on: " + match.matchedOn +
      "\nCould not auto-detect comment column.\nRow data:\n" + JSON.stringify(match.row, null, 2);
    matchEl.className = "status-box warn";
    showSection(matchSection);
  }
}
