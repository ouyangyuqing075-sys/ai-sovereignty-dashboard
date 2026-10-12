import test from "node:test";
import assert from "node:assert/strict";
import { categoryForFixed, countCategories, excelColumnLetter, identifyColumns, normalizeRows, rowMatchesCategory, themeMatches, cloneDefaultThemes } from "../src/logic.js";

test("identifies only usable question headers and preserves their real positions", () => {
  const headers = ["ID", "Points - How would you define AI Sovereignty?", "How would you define AI Sovereignty?", "When you use an AI tool, how much do you worry about personal data country company owns?", "Should Australia rely on foreign tech giants or build local AI technology?", "If foreign AI apps were banned, what daily student task becomes difficult?", "Australian business local AI or foreign AI"];
  const { questionColumns, idIndex } = identifyColumns(headers);
  assert.equal(idIndex, 0);
  assert.equal(questionColumns.q1.letter, "C");
  assert.equal(questionColumns.q2.letter, "F");
  assert.equal(questionColumns.q5.letter, "G");
});

test("fixed groups keep uncertainty, missing values and mixed positions separate", () => {
  assert.equal(categoryForFixed("I am unsure about where data goes", "privacy"), "Uncertain");
  assert.equal(categoryForFixed("", "privacy"), "Missing answer");
  assert.equal(categoryForFixed("Use both local and foreign tools", "preference"), "Use both local and foreign AI");
});

test("themes overlap without double-counting repeated keywords within one theme", () => {
  const themes = cloneDefaultThemes();
  const matched = themeMatches("I use Python coding for research ideas", themes.q2);
  assert.deepEqual(matched.names, ["Coding and technical tools", "Research and ideas"]);
  assert.deepEqual(matched.keywords["Coding and technical tools"], ["coding", "python"]);
  const rows = [{ answers: { q2: "Python coding and coding" } }];
  assert.equal(countCategories(rows, "q2", themes).get("Coding and technical tools"), 1);
});

test("normalised rows retain workbook IDs and original Excel row numbers", () => {
  const values = [["ID", "How would you define AI Sovereignty?"], ["7", "Control over systems"], ["", ""]];
  const data = normalizeRows(values);
  assert.equal(data.rows.length, 1);
  assert.equal(data.rows[0].id, "7");
  assert.equal(data.rows[0].originalRow, 2);
  assert.equal(excelColumnLetter(26), "AA");
  assert.equal(rowMatchesCategory({ answers: { q2: "Python coding" } }, { question: "q2", category: "Coding and technical tools" }, cloneDefaultThemes()), true);
});
