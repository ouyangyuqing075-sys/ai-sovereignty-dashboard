export const QUESTION_KEYS = ["q1", "q2", "q3", "q4", "q5"];

export const QUESTION_META = {
  q1: { short: "AI sovereignty", number: 1, matcher: (header) => header.includes("ai sovereignty") },
  q2: { short: "Everyday reliance", number: 2, matcher: (header) => header.includes("foreign ai apps") || (header.includes("daily student") && header.includes("difficult")) },
  q3: { short: "Privacy concern", number: 3, matcher: (header) => header.includes("personal data") && (header.includes("country") || header.includes("company")) },
  q4: { short: "AI preference", number: 4, matcher: (header) => header.includes("foreign tech giants") || (header.includes("local ai technology") && header.includes("rely")) },
  q5: { short: "Business considerations", number: 5, matcher: (header) => header.includes("australian business") && (header.includes("local ai") || header.includes("foreign ai")) }
};

export const FIXED_RULES = {
  privacy: [
    ["Low concern", /\b(worry a little|a little bit|not worry|do not worry|don't worry|not much|convenience)\b/i],
    ["Moderately concerned", /\b(moderately|somewhat|moderate concern)\b/i],
    ["Highly concerned", /\b(worry a lot|very concerned|highly concerned|a great deal)\b/i],
    ["Uncertain", /\b(unsure|not sure|don't know|do not know)\b/i]
  ],
  preference: [
    ["Use both local and foreign AI", /\b(use both|both local and foreign|mix of local and foreign|hybrid)\b/i],
    ["Develop local capability and retain foreign access", /\b(invest in local|local capability|local infrastructure|local skills|retain foreign|keep access|do not try to build everything)\b/i],
    ["Build local AI", /\b(build (our )?own local|build local ai|own local ai|national interests)\b/i],
    ["Rely on foreign AI", /\b(rely on foreign|foreign providers|better tools at lower cost)\b/i],
    ["Unsure", /\b(unsure|not sure|do not know|don't know)\b/i]
  ]
};

export const DEFAULT_THEMES = {
  q2: [
    { name: "Coding and technical tools", keywords: ["coding", "debug", "python", "programming", "code"] },
    { name: "Understanding course material", keywords: ["lecture", "readings", "explain unfamiliar", "understand", "course material"] },
    { name: "Writing and communication", keywords: ["essay", "grammar", "email", "writing", "outline"] },
    { name: "Translation and language", keywords: ["translat", "mandarin", "language"] },
    { name: "Study preparation", keywords: ["practice questions", "exam", "study", "notes"] },
    { name: "Research and ideas", keywords: ["research", "sources", "ideas", "brainstorm"] },
    { name: "Limited reliance", keywords: ["could still", "not rely", "limited", "rarely"] }
  ],
  q5: [
    { name: "Privacy and data handling", keywords: ["data", "stored", "hosting", "personal information"] },
    { name: "Security and confidentiality", keywords: ["security", "confidential", "secrets", "leak", "trade secret"] },
    { name: "Legal obligations", keywords: ["legal", "rules", "contract", "compliance", "obligations", "audit"] },
    { name: "Supplier dependence", keywords: ["provider", "supplier", "prices", "depend", "foreign provider"] },
    { name: "Reliability and continuity", keywords: ["outage", "backup", "access", "dispute", "continuity"] },
    { name: "Cost and performance", keywords: ["cost", "price", "performance", "better tools"] },
    { name: "Local context and jobs", keywords: ["local", "australian", "jobs", "skills"] }
  ]
};

const excludedHeader = (header) => /(^|\s|-)(points?|feedback)(\s|-|$)/i.test(header);
const missing = (value) => value === null || value === undefined || String(value).trim() === "";

export function excelColumnLetter(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

export function identifyColumns(headers) {
  const normalized = headers.map((header) => String(header ?? "").trim());
  const questionColumns = {};
  QUESTION_KEYS.forEach((key) => {
    const meta = QUESTION_META[key];
    const index = normalized.findIndex((header) => !excludedHeader(header) && meta.matcher(header.toLowerCase()));
    if (index !== -1) questionColumns[key] = { index, header: normalized[index], letter: excelColumnLetter(index) };
  });
  const idIndex = normalized.findIndex((header) => /^(response )?id$/i.test(header) || /^id\b/i.test(header));
  return { questionColumns, idIndex };
}

export function categoryForFixed(answer, type) {
  if (missing(answer)) return "Missing answer";
  const matched = FIXED_RULES[type].find(([, pattern]) => pattern.test(String(answer)));
  return matched ? matched[0] : "Other / review";
}

export function themeMatches(answer, themes) {
  if (missing(answer)) return { names: ["Missing answer"], keywords: {} };
  const text = String(answer).toLowerCase();
  const keywords = {};
  themes.forEach((theme) => {
    const found = theme.keywords.filter((keyword) => keyword.trim() && text.includes(keyword.trim().toLowerCase()));
    if (found.length) keywords[theme.name] = found;
  });
  const names = Object.keys(keywords);
  return { names: names.length ? names : ["Unmatched"], keywords };
}

export function normalizeRows(values) {
  const headers = values[0] || [];
  const { questionColumns, idIndex } = identifyColumns(headers);
  const rows = values.slice(1).map((source, rowIndex) => {
    const hasContent = source.some((value) => !missing(value));
    if (!hasContent) return null;
    const answers = Object.fromEntries(QUESTION_KEYS.map((key) => [key, questionColumns[key] ? String(source[questionColumns[key].index] ?? "").trim() : ""]));
    const id = idIndex >= 0 && !missing(source[idIndex]) ? String(source[idIndex]).trim() : `Row ${rowIndex + 2}`;
    return { id, originalRow: rowIndex + 2, answers };
  }).filter(Boolean);
  return { headers, questionColumns, idIndex, rows, missingQuestions: QUESTION_KEYS.filter((key) => !questionColumns[key]) };
}

export function countMissing(rows) {
  return rows.reduce((total, row) => total + QUESTION_KEYS.filter((key) => missing(row.answers[key])).length, 0);
}

export function countCategories(rows, questionKey, themes) {
  const counts = new Map();
  rows.forEach((row) => {
    let labels;
    if (questionKey === "q3") labels = [categoryForFixed(row.answers.q3, "privacy")];
    else if (questionKey === "q4") labels = [categoryForFixed(row.answers.q4, "preference")];
    else labels = themeMatches(row.answers[questionKey], themes[questionKey]).names;
    labels.forEach((label) => counts.set(label, (counts.get(label) || 0) + 1));
  });
  return counts;
}

export function rowMatchesCategory(row, selection, themes) {
  if (!selection) return true;
  if (selection.question === "q3") return categoryForFixed(row.answers.q3, "privacy") === selection.category;
  if (selection.question === "q4") return categoryForFixed(row.answers.q4, "preference") === selection.category;
  return themeMatches(row.answers[selection.question], themes[selection.question]).names.includes(selection.category);
}

export function cloneDefaultThemes() {
  return JSON.parse(JSON.stringify(DEFAULT_THEMES));
}

export function parseThemeText(text) {
  return String(text).split(",").map((keyword) => keyword.trim()).filter(Boolean);
}
