import {
  DEFAULT_THEMES,
  FIXED_RULES,
  QUESTION_KEYS,
  QUESTION_META,
  categoryForFixed,
  cloneDefaultThemes,
  countCategories,
  countMissing,
  normalizeRows,
  parseThemeText,
  rowMatchesCategory,
  themeMatches
} from "./logic.js";

const DEMO_PATH = "./data/AI_Sovereignty_35_Synthetic_Responses.xlsx";
const PRIVACY_ORDER = ["Low concern", "Moderately concerned", "Highly concerned", "Uncertain", "Other / review", "Missing answer"];
const PREFERENCE_ORDER = ["Use both local and foreign AI", "Develop local capability and retain foreign access", "Build local AI", "Rely on foreign AI", "Unsure", "Other / review", "Missing answer"];

const state = {
  workbook: null,
  data: null,
  filename: "",
  sheetName: "",
  message: "Files are read only in this browser. Nothing is sent to an account, database or upload server.",
  themes: cloneDefaultThemes(),
  filters: { privacy: "", preference: "", search: "", chart: null }
};

const $ = (selector) => document.querySelector(selector);
const escapeHtml = (value) => String(value ?? "").replace(/[&<>'"]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[character]);
const isMissing = (value) => String(value ?? "").trim() === "";
const percentage = (count, denominator) => denominator ? Math.round((count / denominator) * 100) : 0;
const motion = () => window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth";

function setMessage(message, tone = "") {
  state.message = message;
  const target = $("#workbook-message");
  target.textContent = message;
  target.className = `help-text ${tone}`.trim();
}

async function ensureXlsx() {
  for (let index = 0; index < 25; index += 1) {
    if (window.XLSX) return window.XLSX;
    await new Promise((resolve) => window.setTimeout(resolve, 120));
  }
  throw new Error("The Excel reader could not load. Check your internet connection and try again.");
}

function currentRows() {
  if (!state.data) return [];
  const search = state.filters.search.trim().toLowerCase();
  return state.data.rows.filter((row) => {
    if (state.filters.privacy && categoryForFixed(row.answers.q3, "privacy") !== state.filters.privacy) return false;
    if (state.filters.preference && categoryForFixed(row.answers.q4, "preference") !== state.filters.preference) return false;
    if (state.filters.chart && !rowMatchesCategory(row, state.filters.chart, state.themes)) return false;
    if (!search) return true;
    return [row.id, ...QUESTION_KEYS.map((key) => row.answers[key])].join(" ").toLowerCase().includes(search);
  });
}

function categoryOrder(question) {
  if (question === "q3") return PRIVACY_ORDER;
  if (question === "q4") return PREFERENCE_ORDER;
  return [...state.themes[question].map((theme) => theme.name), "Unmatched", "Missing answer"];
}

function questionColumn(question) {
  return state.data?.questionColumns[question];
}

function sourceLine(question) {
  const column = questionColumn(question);
  const meta = QUESTION_META[question];
  if (!column) return "";
  return `<div class="source-details"><p><strong>Source · Survey Question ${meta.number}</strong></p><p class="source-question">${escapeHtml(column.header)}</p><p>${escapeHtml(state.filename)} · ${escapeHtml(state.sheetName)} · Excel column ${column.letter}</p></div>`;
}

function chartRuleCopy(question) {
  if (question === "q3") return "The dashboard checks the answer against the fixed wording rules in order. The first match becomes a low, moderate, high or uncertain concern summary. Unfamiliar wording and missing answers stay separate.";
  if (question === "q4") return "The dashboard checks the answer against fixed wording rules in order. These are summaries of expressed positions, not rankings or correctness scores. Mixed and uncertain answers are kept distinct.";
  return "These are provisional keyword themes. A response counts once per matching theme, can match several themes, and can therefore make theme percentages total more than 100%. Unmatched and missing answers are separate.";
}

function chartExamples(question, rows, labels) {
  if (!["q2", "q5"].includes(question)) return "";
  const items = labels.flatMap((label) => {
    if (["Unmatched", "Missing answer"].includes(label)) return [];
    const example = rows.find((row) => themeMatches(row.answers[question], state.themes[question]).names.includes(label));
    if (!example) return [];
    return [`<li><strong>${escapeHtml(label)}:</strong> “${escapeHtml(example.answers[question])}” <span class="example-meta">Response ${escapeHtml(example.id)}, row ${example.originalRow}</span></li>`];
  });
  if (!items.length) return "";
  return `<details class="chart-examples"><summary>Original answer examples</summary><p>Each example is the first matching answer in the currently selected rows. It is evidence from the worksheet, not a separate dataset or a claim about every answer in the theme.</p><ul>${items.join("")}</ul></details>`;
}

function renderChart(question, target) {
  const rows = currentRows();
  const counts = countCategories(rows, question, state.themes);
  const labels = categoryOrder(question);
  const denominator = rows.length;
  const column = questionColumn(question);
  if (!column) {
    target.innerHTML = `<p class="empty-state">This worksheet does not contain the recognised source column for this chart.</p>`;
    return;
  }
  const missingCount = counts.get("Missing answer") || 0;
  const visibleLabels = labels.filter((label) => counts.has(label) || question === "q3" || question === "q4");
  const max = Math.max(1, ...visibleLabels.map((label) => counts.get(label) || 0));
  const bars = visibleLabels.map((label) => {
    const count = counts.get(label) || 0;
    const selected = state.filters.chart?.question === question && state.filters.chart?.category === label;
    return `<button class="chart-bar ${selected ? "is-selected" : ""}" type="button" data-chart-question="${question}" data-chart-category="${escapeHtml(label)}" aria-pressed="${selected}"><span class="chart-label">${escapeHtml(label)}</span><span class="bar-track"><span class="bar-fill" style="width:${(count / max) * 100}%"></span></span><span class="chart-value">${count} <small>${percentage(count, denominator)}%</small></span></button>`;
  }).join("");
  const countDescription = question === "q2" || question === "q5"
    ? `Each percentage uses all ${denominator} selected response${denominator === 1 ? "" : "s"} as its denominator, including missing answers. ${denominator - missingCount} answered cell${denominator - missingCount === 1 ? "" : "s"}; ${missingCount} missing cell${missingCount === 1 ? "" : "s"}.`
    : `Counts and percentages use all ${denominator} selected response${denominator === 1 ? "" : "s"}. ${denominator - missingCount} answered cell${denominator - missingCount === 1 ? "" : "s"}; ${missingCount} missing cell${missingCount === 1 ? "" : "s"}.`;
  target.innerHTML = `${sourceLine(question)}<p class="chart-method">${chartRuleCopy(question)}</p><div class="chart-bars" role="group" aria-label="${escapeHtml(QUESTION_META[question].short)} categories">${bars || `<p class="empty-state">No selected responses match the current filters.</p>`}</div><p class="chart-note">${countDescription} Click a category to filter the entire dashboard and jump to its original responses. Click the selected category again or use Reset filters to clear it.</p>${chartExamples(question, rows, visibleLabels)}`;
}

function renderWorkbook() {
  const hasData = Boolean(state.data);
  $("#dataset-badge").textContent = state.data?.synthetic ? "Synthetic classroom data" : hasData ? "Workbook data" : "No readable data";
  $("#dataset-badge").classList.toggle("synthetic", Boolean(state.data?.synthetic));
  $("#file-summary").textContent = hasData ? `${state.filename} · ${state.sheetName} · ${state.data.rows.length} response rows · ${countMissing(state.data.rows)} missing answer cells` : "Choose an .xlsx workbook with the five survey question columns.";
  const selector = $("#worksheet-select");
  selector.disabled = !state.workbook;
  selector.innerHTML = state.workbook ? state.workbook.SheetNames.map((sheet) => `<option value="${escapeHtml(sheet)}" ${sheet === state.sheetName ? "selected" : ""}>${escapeHtml(sheet)}</option>`).join("") : "";
}

function renderExplanation() {
  const target = $("#data-explanation");
  if (!state.data) {
    target.innerHTML = `<p>Load the demonstration workbook or upload a readable .xlsx file. The dashboard will look for the five question headers, not fixed column positions.</p>`;
    return;
  }
  const { rows } = state.data;
  target.innerHTML = `<p><strong>${escapeHtml(state.filename)}</strong>, worksheet <strong>${escapeHtml(state.sheetName)}</strong>, is being analysed. Each spreadsheet row is one respondent; each of the five question columns is that respondent’s answer. The dashboard adds counts, percentages, shorter headings, category labels and themes, but does not invent or rewrite original answers.</p><p>IDs are retained so you can trace a response, while names, email addresses, timestamps, points and feedback are excluded. A <strong>response</strong> means one person’s row; an <strong>answer</strong> means one question cell, so one response can supply five answers. ${rows.length} response rows are currently available before filters. Read a chart’s exact source question, then click a bar to inspect matching original answers in the table.</p>${state.data.synthetic ? `<p class="synthetic-callout"><strong>Synthetic classroom data:</strong> these are made-up teaching examples, not findings about real students.</p>` : ""}`;
}

function renderFilterControls() {
  const privacy = $("#privacy-filter");
  const preference = $("#preference-filter");
  privacy.innerHTML = `<option value="">All privacy concern summaries</option>${PRIVACY_ORDER.map((category) => `<option value="${escapeHtml(category)}" ${state.filters.privacy === category ? "selected" : ""}>${escapeHtml(category)}</option>`).join("")}`;
  preference.innerHTML = `<option value="">All AI preference summaries</option>${PREFERENCE_ORDER.map((category) => `<option value="${escapeHtml(category)}" ${state.filters.preference === category ? "selected" : ""}>${escapeHtml(category)}</option>`).join("")}`;
  $("#search-input").value = state.filters.search;
  const active = [];
  if (state.filters.privacy) active.push(`privacy: ${state.filters.privacy}`);
  if (state.filters.preference) active.push(`AI preference: ${state.filters.preference}`);
  if (state.filters.chart) active.push(`${QUESTION_META[state.filters.chart.question].short}: ${state.filters.chart.category}`);
  if (state.filters.search) active.push(`search: “${state.filters.search}”`);
  const rows = currentRows();
  $("#active-filters").textContent = active.length ? `Active filters: ${active.join(" · ")}. ${rows.length} respondent row${rows.length === 1 ? "" : "s"} remain.` : `No active filters. ${rows.length} respondent row${rows.length === 1 ? "" : "s"} remain.`;
}

function topSummary(rows, question) {
  const countMap = countCategories(rows, question, state.themes);
  const labels = categoryOrder(question).filter((label) => !["Missing answer", "Unmatched", "Other / review"].includes(label));
  const max = Math.max(0, ...labels.map((label) => countMap.get(label) || 0));
  if (!max) return { labels: [], count: 0 };
  return { labels: labels.filter((label) => (countMap.get(label) || 0) === max), count: max };
}

function renderOverview() {
  const rows = currentRows();
  const total = state.data.rows.length;
  const missing = countMissing(rows);
  $("#overview-totals").innerHTML = [
    [rows.length, "Selected responses", `of ${total} worksheet response rows`],
    [total, "Worksheet responses", "each row is one respondent"],
    [missing, "Missing answers", "blank cells in the five survey questions only"]
  ].map(([value, label, note]) => `<article class="metric"><strong>${value}</strong><span>${label}</span><small>${note}</small></article>`).join("");
  if (!rows.length) {
    $("#observations").innerHTML = `<p class="empty-state">No respondent rows match the current filters. Reset filters to return to the worksheet view.</p>`;
    return;
  }
  const summaries = [["q3", "#privacy"], ["q4", "#investment"], ["q2", "#reliance"]].map(([question, href]) => {
    const top = topSummary(rows, question);
    if (!top.count) return `<a href="${href}">No recognised ${QUESTION_META[question].short.toLowerCase()} category appears in this view.</a>`;
    const phrase = top.labels.length > 1 ? `a tie between ${top.labels.join(" and ")}` : top.labels[0];
    return `<a href="${href}"><strong>${top.count} of ${rows.length} (${percentage(top.count, rows.length)}%)</strong> fall into ${escapeHtml(phrase)} for ${QUESTION_META[question].short.toLowerCase()}.</a>`;
  });
  $("#observations").innerHTML = `<p class="observation-intro">These observations describe this selected workbook view only. They do not show causes or support claims beyond these responses.</p>${summaries.join("")}`;
}

function formatMatchKeywords(matches) {
  return Object.entries(matches).map(([theme, keywords]) => `${theme}: ${keywords.join(", ")}`).join("; ");
}

function renderThemeMethod() {
  const rows = currentRows();
  let example = null;
  ["q2", "q5"].some((question) => rows.some((row) => {
    const match = themeMatches(row.answers[question], state.themes[question]);
    if (match.names[0] !== "Unmatched" && match.names[0] !== "Missing answer") {
      example = { question, row, match };
      return true;
    }
    return false;
  }));
  const exampleHtml = example ? `<div class="worked-example"><h3>Worked example from this view</h3><p><strong>Question ${QUESTION_META[example.question].number} · response ${escapeHtml(example.row.id)} · Excel row ${example.row.originalRow}</strong></p><blockquote>“${escapeHtml(example.row.answers[example.question])}”</blockquote><p><strong>Matched keyword or phrase:</strong> ${escapeHtml(formatMatchKeywords(example.match.keywords))}</p><p><strong>Assigned theme${example.match.names.length === 1 ? "" : "s"}:</strong> ${escapeHtml(example.match.names.join(", "))}. This answer contributes one count to each listed theme.</p></div>` : `<p class="empty-state">There is no suitable matched Question 2 or Question 5 answer in the current view, so a worked example cannot be shown.</p>`;
  $("#theme-method").innerHTML = `<p>A <strong>theme</strong> is a topic label added by the dashboard. It is not a survey question or an original spreadsheet column. Theme grouping applies only to the written answers to Questions 2 and 5.</p><p>Each theme has a list of keywords or phrases. An answer matches when it contains at least one entry, ignoring capital letters; matching can happen inside longer words. An answer counts once per matching theme even when several keywords for that theme occur. One answer can match several themes, so percentages can add to more than 100%.</p><p>A theme percentage is its matching-answer count divided by all selected responses, including respondents whose relevant cell is missing. <strong>Unmatched</strong> means an answered cell matched no theme keywords. <strong>Missing answer</strong> means the question cell was empty. Keyword matching does not understand meaning, context or negation: “not secure” can still match “secure.” These are provisional discussion groupings, not definitive qualitative analysis or correctness judgements.</p>${exampleHtml}`;
}

function ruleText(theme) {
  return theme.keywords.join(", ");
}

function renderRules() {
  const controls = ["q2", "q5"].flatMap((question) => state.themes[question].map((theme, index) => `<label for="rule-${question}-${index}"><span>Question ${QUESTION_META[question].number} · ${escapeHtml(theme.name)}</span><textarea id="rule-${question}-${index}" name="${question}:${index}" rows="2">${escapeHtml(ruleText(theme))}</textarea></label>`));
  $("#rules-form").innerHTML = controls.join("");
  const privacyRules = FIXED_RULES.privacy.map(([label, pattern]) => `<li><strong>${label}:</strong> ${escapeHtml(pattern.source)}</li>`).join("");
  const preferenceRules = FIXED_RULES.preference.map(([label, pattern]) => `<li><strong>${label}:</strong> ${escapeHtml(pattern.source)}</li>`).join("");
  $("#fixed-rules-content").innerHTML = `<p>These groups are separate from the editable theme rules. The dashboard evaluates each list in order and uses the first matching rule. A blank cell becomes Missing answer; an answered cell that matches no rule becomes Other / review.</p><h3>Privacy concern</h3><ol>${privacyRules}</ol><h3>AI preference</h3><ol>${preferenceRules}</ol>`;
}

function definitionCards(rows) {
  const seen = new Set();
  return rows.filter((row) => {
    const value = row.answers.q1;
    if (isMissing(value) || seen.has(value.toLowerCase())) return false;
    seen.add(value.toLowerCase());
    return true;
  }).map((row) => `<figure class="definition-card"><blockquote>“${escapeHtml(row.answers.q1)}”</blockquote><figcaption>Response ${escapeHtml(row.id)} · Excel row ${row.originalRow}</figcaption></figure>`).join("");
}

function answerCell(value) {
  if (isMissing(value)) return `<span class="missing-answer">Missing answer</span>`;
  const answer = String(value);
  if (answer.length <= 165) return escapeHtml(answer);
  return `<details class="full-answer"><summary>${escapeHtml(answer.slice(0, 165))}… <span>Read full answer</span></summary><p>${escapeHtml(answer)}</p></details>`;
}

function renderResponses() {
  const rows = currentRows();
  const guide = QUESTION_KEYS.map((question) => {
    const column = questionColumn(question);
    return `<article><strong>${QUESTION_META[question].short}</strong><span>Question ${QUESTION_META[question].number} · column ${column?.letter || "?"}</span><p>${escapeHtml(column?.header || "Source column not found")}</p></article>`;
  }).join("");
  $("#column-guide").innerHTML = `<p>Question-to-column guide: the five shortened headings below map directly to the current workbook headers.</p><div>${guide}</div>`;
  $("#response-head").innerHTML = `<tr><th scope="col">Workbook ID</th><th scope="col">Excel row</th>${QUESTION_KEYS.map((key) => `<th scope="col">${QUESTION_META[key].short}<span>Q${QUESTION_META[key].number} · ${questionColumn(key)?.letter || "?"}</span></th>`).join("")}</tr>`;
  $("#response-body").innerHTML = rows.length ? rows.map((row) => `<tr><th scope="row">${escapeHtml(row.id)}</th><td>${row.originalRow}</td>${QUESTION_KEYS.map((key) => `<td>${answerCell(row.answers[key])}</td>`).join("")}</tr>`).join("") : `<tr><td colspan="7" class="empty-table">No respondent rows match the current filters. Reset filters to restore the worksheet view.</td></tr>`;
  $("#table-status").textContent = `${rows.length} of ${state.data.rows.length} response rows shown`;
}

function renderDefinitions() {
  const rows = currentRows();
  $("#definitions-source").innerHTML = sourceLine("q1");
  $("#definition-cards").innerHTML = definitionCards(rows) || `<p class="empty-state">No non-blank Question 1 answers match the current filters.</p>`;
}

function renderAll() {
  renderWorkbook();
  renderExplanation();
  if (!state.data) {
    ["#overview-totals", "#observations", "#privacy-chart", "#preference-chart", "#reliance-chart", "#business-chart", "#theme-method", "#definition-cards", "#response-body"].forEach((selector) => { $(selector).innerHTML = `<p class="empty-state">Load a worksheet to see this part of the dashboard.</p>`; });
    return;
  }
  renderFilterControls();
  renderOverview();
  renderChart("q3", $("#privacy-chart"));
  renderChart("q4", $("#preference-chart"));
  renderChart("q2", $("#reliance-chart"));
  renderChart("q5", $("#business-chart"));
  renderThemeMethod();
  renderRules();
  renderDefinitions();
  renderResponses();
}

function applySheet(sheetName) {
  if (!state.workbook) return;
  const values = window.XLSX.utils.sheet_to_json(state.workbook.Sheets[sheetName], { header: 1, defval: "", raw: false, blankrows: false });
  const parsed = normalizeRows(values);
  state.sheetName = sheetName;
  if (parsed.missingQuestions.length) {
    state.data = null;
    setMessage(`This worksheet is missing recognised source columns for ${parsed.missingQuestions.map((key) => `Question ${QUESTION_META[key].number}`).join(", ")}. Choose another worksheet or upload a workbook that has all five questions.`, "error");
    renderAll();
    return;
  }
  state.data = { ...parsed, synthetic: values.flat().some((value) => String(value).toLowerCase().includes("synthetic data")) };
  state.filters = { privacy: "", preference: "", search: "", chart: null };
  setMessage(`Workbook read successfully. ${parsed.rows.length} respondent rows were found; blanks in the five question columns remain visible as Missing answer.`, "success");
  renderAll();
}

async function loadWorkbook(arrayBuffer, filename) {
  try {
    await ensureXlsx();
    const workbook = window.XLSX.read(arrayBuffer, { type: "array" });
    if (!workbook.SheetNames?.length) throw new Error("The workbook does not contain any worksheets.");
    state.workbook = workbook;
    state.filename = filename;
    applySheet(workbook.SheetNames[0]);
  } catch (error) {
    state.workbook = null;
    state.data = null;
    setMessage(`Unable to read this workbook. Use a valid .xlsx file with the five survey question columns. ${error.message}`, "error");
    renderAll();
  }
}

async function loadDemo() {
  try {
    setMessage("Loading the supplied demonstration workbook…");
    const response = await fetch(DEMO_PATH);
    if (!response.ok) throw new Error(`The demonstration file returned ${response.status}.`);
    await loadWorkbook(await response.arrayBuffer(), DEMO_PATH.split("/").pop());
  } catch (error) {
    setMessage(`The demonstration workbook could not be loaded. ${error.message}`, "error");
  }
}

function chooseFile(file) {
  if (!file) return;
  if (!/\.xlsx$/i.test(file.name)) {
    setMessage("Unsupported file. Please choose an .xlsx workbook.", "error");
    return;
  }
  file.arrayBuffer().then((buffer) => loadWorkbook(buffer, file.name));
}

function selectChart(question, category) {
  const existing = state.filters.chart;
  state.filters.chart = existing?.question === question && existing?.category === category ? null : { question, category };
  renderAll();
  if (state.filters.chart) {
    window.setTimeout(() => {
      const section = $("#responses");
      section.scrollIntoView({ behavior: motion(), block: "start" });
      $("#responses-heading").focus({ preventScroll: true });
    }, 0);
  }
}

function applyRules() {
  const edited = {};
  ["q2", "q5"].forEach((question) => {
    edited[question] = state.themes[question].map((theme, index) => ({ ...theme, keywords: parseThemeText($("#rule-" + question + "-" + index).value) }));
  });
  state.themes = edited;
  if (state.filters.chart && ["q2", "q5"].includes(state.filters.chart.question) && !categoryOrder(state.filters.chart.question).includes(state.filters.chart.category)) state.filters.chart = null;
  setMessage("Theme rules applied. Charts, examples and any active theme filter now use the edited keyword lists. Original workbook answers are unchanged.", "success");
  renderAll();
}

function attachEvents() {
  $("#workbook-input").addEventListener("change", (event) => chooseFile(event.target.files[0]));
  $("#load-demo").addEventListener("click", loadDemo);
  $("#worksheet-select").addEventListener("change", (event) => applySheet(event.target.value));
  $("#privacy-filter").addEventListener("change", (event) => { state.filters.privacy = event.target.value; renderAll(); });
  $("#preference-filter").addEventListener("change", (event) => { state.filters.preference = event.target.value; renderAll(); });
  $("#search-input").addEventListener("input", (event) => { state.filters.search = event.target.value; renderAll(); });
  $("#reset-filters").addEventListener("click", () => { state.filters = { privacy: "", preference: "", search: "", chart: null }; renderAll(); });
  document.addEventListener("click", (event) => {
    const button = event.target.closest("[data-chart-question]");
    if (button) selectChart(button.dataset.chartQuestion, button.dataset.chartCategory);
  });
  $("#rules-form").addEventListener("submit", (event) => { event.preventDefault(); applyRules(); });
  $("#restore-rules").addEventListener("click", () => { state.themes = cloneDefaultThemes(); setMessage("Default theme rules restored. Original workbook answers are unchanged.", "success"); renderAll(); });
  const dropTarget = $("#drop-target");
  ["dragenter", "dragover"].forEach((eventName) => dropTarget.addEventListener(eventName, (event) => { event.preventDefault(); dropTarget.classList.add("is-dragging"); }));
  ["dragleave", "drop"].forEach((eventName) => dropTarget.addEventListener(eventName, (event) => { event.preventDefault(); dropTarget.classList.remove("is-dragging"); }));
  dropTarget.addEventListener("drop", (event) => chooseFile(event.dataTransfer.files[0]));
  dropTarget.addEventListener("click", () => $("#workbook-input").click());
  dropTarget.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); $("#workbook-input").click(); } });
  $("#back-to-top").addEventListener("click", () => { $("#page-title").scrollIntoView({ behavior: motion(), block: "start" }); $("#page-title").focus({ preventScroll: true }); });
}

attachEvents();
renderAll();
loadDemo();
