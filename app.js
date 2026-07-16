const state = {
  rows: [],
  questions: [],
  currentIndex: 0,
  score: 0,
  selectedChoice: "",
  answers: [],
};

const els = {
  appLayout: document.querySelector("#appLayout"),
  sidebarToggle: document.querySelector("#sidebarToggle"),
  navItems: document.querySelectorAll(".nav-item"),
  pagePanels: document.querySelectorAll("[data-page-panel]"),
  pageTitle: document.querySelector("#pageTitle"),
  pageEyebrow: document.querySelector("#pageEyebrow"),
  fileInput: document.querySelector("#fileInput"),
  statusLine: document.querySelector("#statusLine"),
  mappingState: document.querySelector("#mappingState"),
  goToQuizButton: document.querySelector("#goToQuizButton"),
  questionColumn: document.querySelector("#questionColumn"),
  answerColumn: document.querySelector("#answerColumn"),
  optionStartColumn: document.querySelector("#optionStartColumn"),
  optionEndColumn: document.querySelector("#optionEndColumn"),
  explanationColumn: document.querySelector("#explanationColumn"),
  categoryColumn: document.querySelector("#categoryColumn"),
  modeSelect: document.querySelector("#modeSelect"),
  limitInput: document.querySelector("#limitInput"),
  categoryFilter: document.querySelector("#categoryFilter"),
  startButton: document.querySelector("#startButton"),
  questionCount: document.querySelector("#questionCount"),
  sidebarQuestionCount: document.querySelector("#sidebarQuestionCount"),
  scoreSummary: document.querySelector("#scoreSummary"),
  emptyState: document.querySelector("#emptyState"),
  quizCard: document.querySelector("#quizCard"),
  resultCard: document.querySelector("#resultCard"),
  progressText: document.querySelector("#progressText"),
  categoryText: document.querySelector("#categoryText"),
  questionText: document.querySelector("#questionText"),
  choiceList: document.querySelector("#choiceList"),
  answerInput: document.querySelector("#answerInput"),
  feedback: document.querySelector("#feedback"),
  explanation: document.querySelector("#explanation"),
  nextButton: document.querySelector("#nextButton"),
  finalScore: document.querySelector("#finalScore"),
  finalDetail: document.querySelector("#finalDetail"),
  wrongList: document.querySelector("#wrongList"),
  restartButton: document.querySelector("#restartButton"),
};

const pages = {
  quiz: { title: "開始測驗", eyebrow: "Quiz" },
  settings: { title: "題庫設定", eyebrow: "Question bank" },
};

const emptyOption = { label: "不使用", value: "" };
const aliases = {
  question: ["題目", "問題", "question", "q", "題幹", "題庫"],
  answer: ["答案", "正解", "answer", "a", "解答", "正確答案"],
  optionA: ["選項a", "選項1", "choicea", "choice1", "optiona", "option1", "a"],
  optionB: ["選項b", "選項2", "choiceb", "choice2", "optionb", "option2", "b"],
  optionC: ["選項c", "選項3", "choicec", "choice3", "optionc", "option3", "c"],
  optionD: ["選項d", "選項4", "choiced", "choice4", "optiond", "option4", "d"],
  explanation: ["說明", "解析", "詳解", "explanation", "description", "note", "備註"],
  category: ["分類", "類別", "category", "章節", "單元"],
};

els.fileInput.addEventListener("change", handleFile);
els.startButton.addEventListener("click", startQuiz);
els.goToQuizButton.addEventListener("click", () => switchPage("quiz"));
els.sidebarToggle.addEventListener("click", toggleSidebar);
els.navItems.forEach((item) => {
  item.addEventListener("click", () => switchPage(item.dataset.page));
});
document.querySelectorAll("[data-go-page]").forEach((item) => {
  item.addEventListener("click", () => switchPage(item.dataset.goPage));
});
window.addEventListener("hashchange", () => {
  switchPage(location.hash.slice(1), false);
});
els.answerInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    submitAnswer();
  }
});
els.nextButton.addEventListener("click", nextQuestion);
els.restartButton.addEventListener("click", startQuiz);

if (localStorage.getItem("quizSidebarCollapsed") === "true") {
  els.appLayout.classList.add("sidebar-collapsed");
  els.sidebarToggle.title = "展開側欄";
  els.sidebarToggle.setAttribute("aria-label", "展開側欄");
}
switchPage(location.hash.slice(1) || "quiz", false);

for (const select of [
  els.questionColumn,
  els.answerColumn,
  els.optionStartColumn,
  els.optionEndColumn,
  els.explanationColumn,
  els.categoryColumn,
]) {
  select.addEventListener("change", rebuildCategoryFilter);
}

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  setStatus("讀取題庫中...");
  const reader = new FileReader();
  reader.onload = (loadEvent) => {
    try {
      const rows = readRows(file, loadEvent.target.result);
      loadRows(rows);
      setStatus(`已載入 ${state.rows.length} 題`);
    } catch (error) {
      state.rows = [];
      setupColumns();
      rebuildCategoryFilter();
      updateCounts();
      els.startButton.disabled = true;
      updateMappingState();
      setStatus(error.message, true);
    }
  };
  reader.onerror = () => setStatus("檔案讀取失敗，請重新選擇檔案。", true);
  reader.readAsArrayBuffer(file);
}

function readRows(file, buffer) {
  if (file.name.toLowerCase().endsWith(".csv")) {
    const text = new TextDecoder("utf-8").decode(buffer);
    return parseCsv(text);
  }

  if (typeof XLSX === "undefined") {
    throw new Error("Excel 解析套件尚未載入，請確認網路連線後重新整理頁面，或先改用 CSV 檔。");
  }

  const data = new Uint8Array(buffer);
  const workbook = XLSX.read(data, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
}

function loadRows(rows) {
  state.rows = rows.map(normalizeRow).filter((row) => Object.values(row).some(Boolean));
  setupColumns();
  rebuildCategoryFilter();
  updateCounts();
  els.startButton.disabled = !canStart();
  updateMappingState();
}

function switchPage(page, updateHash = true) {
  const targetPage = pages[page] ? page : "quiz";
  els.navItems.forEach((item) => item.classList.toggle("active", item.dataset.page === targetPage));
  els.pagePanels.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.pagePanel === targetPage);
  });
  els.pageTitle.textContent = pages[targetPage].title;
  els.pageEyebrow.textContent = pages[targetPage].eyebrow;
  if (updateHash && location.hash !== `#${targetPage}`) {
    history.pushState(null, "", `#${targetPage}`);
  }
}

function toggleSidebar() {
  const collapsed = els.appLayout.classList.toggle("sidebar-collapsed");
  localStorage.setItem("quizSidebarCollapsed", String(collapsed));
  els.sidebarToggle.title = collapsed ? "展開側欄" : "收縮側欄";
  els.sidebarToggle.setAttribute("aria-label", els.sidebarToggle.title);
}

function parseCsv(text) {
  const rows = [];
  let cell = "";
  let row = [];
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && quoted && next === '"') {
      cell += '"';
      i += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell);
  rows.push(row);

  const [headers = [], ...body] = rows.filter((items) => items.some((item) => item.trim()));
  return body.map((items) =>
    Object.fromEntries(headers.map((header, index) => [header.trim(), (items[index] || "").trim()]))
  );
}

function normalizeRow(row) {
  return Object.fromEntries(
    Object.entries(row).map(([key, value]) => [String(key).trim(), String(value ?? "").trim()])
  );
}

function setupColumns() {
  const columns = Object.keys(state.rows[0] || {});
  setOptions(els.questionColumn, columns, guessColumn(columns, aliases.question));
  setOptions(els.answerColumn, columns, guessColumn(columns, aliases.answer));
  const optionStart = guessColumn(columns, aliases.optionA);
  const optionEnd = guessColumn(columns, aliases.optionD) || optionStart;
  setOptions(els.optionStartColumn, [emptyOption, ...columns], optionStart);
  setOptions(els.optionEndColumn, [emptyOption, ...columns], optionEnd);
  setOptions(
    els.explanationColumn,
    [emptyOption, ...columns],
    guessColumn(columns, aliases.explanation)
  );
  setOptions(els.categoryColumn, [emptyOption, ...columns], guessColumn(columns, aliases.category));
}

function setOptions(select, items, selectedValue) {
  select.innerHTML = "";
  for (const item of items) {
    const option = document.createElement("option");
    option.value = typeof item === "string" ? item : item.value;
    option.textContent = typeof item === "string" ? item : item.label;
    select.append(option);
  }
  select.value = selectedValue || "";
}

function guessColumn(columns, names) {
  return columns.find((column) => names.includes(normalizeColumnName(column))) || "";
}

function normalizeColumnName(value) {
  return String(value).trim().replace(/\s+/g, "").replace(/[._-]/g, "").toLowerCase();
}

function rebuildCategoryFilter() {
  const categoryColumn = els.categoryColumn.value;
  const previous = els.categoryFilter.value;
  els.categoryFilter.innerHTML = '<option value="">全部</option>';

  if (categoryColumn) {
    const categories = [...new Set(state.rows.map((row) => row[categoryColumn]).filter(Boolean))].sort();
    for (const category of categories) {
      const option = document.createElement("option");
      option.value = category;
      option.textContent = category;
      els.categoryFilter.append(option);
    }
  }

  els.categoryFilter.value = [...els.categoryFilter.options].some((option) => option.value === previous)
    ? previous
    : "";
  els.startButton.disabled = !canStart();
  els.goToQuizButton.disabled = !canStart();
  updateMappingState();
}

function updateMappingState() {
  const ready = canStart();
  els.mappingState.textContent = ready ? "設定完成" : state.rows.length ? "請確認欄位" : "等待上傳";
  els.mappingState.classList.toggle("ready", ready);
}

function canStart() {
  return state.rows.length > 0 && els.questionColumn.value && els.answerColumn.value;
}

function startQuiz() {
  if (!canStart()) return;

  switchPage("quiz");

  const categoryColumn = els.categoryColumn.value;
  const category = els.categoryFilter.value;
  const questionColumn = els.questionColumn.value;
  const answerColumn = els.answerColumn.value;
  const explanationColumn = els.explanationColumn.value;
  const columns = Object.keys(state.rows[0] || {});
  const optionColumns = getColumnRange(
    columns,
    els.optionStartColumn.value,
    els.optionEndColumn.value
  );

  let pool = state.rows
    .filter((row) => row[questionColumn] && row[answerColumn])
    .filter((row) => !category || row[categoryColumn] === category)
    .map((row) => ({
      question: row[questionColumn],
      answer: row[answerColumn],
      options: optionColumns.map((column) => row[column]).filter(Boolean),
      explanation: explanationColumn ? row[explanationColumn] : "",
      category: categoryColumn ? row[categoryColumn] : "",
    }));

  if (els.modeSelect.value === "random") {
    pool = shuffle(pool);
  }

  const limit = Math.max(1, Number(els.limitInput.value) || 10);
  state.questions = pool.slice(0, limit);
  state.currentIndex = 0;
  state.score = 0;
  state.answers = [];
  updateCounts();
  showQuestion();
}

function getColumnRange(columns, startColumn, endColumn) {
  const startIndex = columns.indexOf(startColumn);
  const endIndex = columns.indexOf(endColumn);
  if (startIndex < 0 || endIndex < startIndex) return [];
  return columns.slice(startIndex, endIndex + 1);
}

function shuffle(items) {
  const copy = [...items];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function showQuestion() {
  if (!state.questions.length) {
    showEmpty("沒有符合條件的題目");
    return;
  }

  const current = state.questions[state.currentIndex];
  state.selectedChoice = "";

  els.emptyState.classList.add("hidden");
  els.resultCard.classList.add("hidden");
  els.quizCard.classList.remove("hidden");
  els.feedback.className = "feedback hidden";
  els.feedback.textContent = "";
  els.explanation.className = "explanation hidden";
  els.explanation.textContent = "";
  els.answerInput.value = "";
  els.answerInput.classList.toggle("hidden", current.options.length > 0);
  els.nextButton.disabled = true;
  els.progressText.textContent = `第 ${state.currentIndex + 1} / ${state.questions.length} 題`;
  els.categoryText.textContent = current.category || "未分類";
  els.questionText.textContent = current.question;
  renderChoices(current);
  updateCounts();
}

function renderChoices(current) {
  els.choiceList.innerHTML = "";
  els.choiceList.classList.toggle("hidden", current.options.length === 0);

  current.options.forEach((choice, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "choice";
    button.innerHTML = `<span class="key">${String.fromCharCode(65 + index)}</span><span>${escapeHtml(choice)}</span>`;
    button.addEventListener("click", () => {
      if (state.answers.length > state.currentIndex) return;
      state.selectedChoice = choice;
      document.querySelectorAll(".choice").forEach((item) => item.classList.remove("selected"));
      button.classList.add("selected");
      submitAnswer(index);
    });
    els.choiceList.append(button);
  });
}

function submitAnswer(selectedIndex = -1) {
  const current = state.questions[state.currentIndex];
  if (!current || state.answers.length > state.currentIndex) return;
  const userAnswer = current.options.length ? state.selectedChoice : els.answerInput.value.trim();
  if (!userAnswer) return;

  const correct = current.options.length
    ? isCorrectOption(selectedIndex, current.answer)
    : isCorrectText(userAnswer, current.answer);
  if (correct) state.score += 1;

  state.answers.push({
    ...current,
    userAnswer,
    correct,
  });

  els.feedback.className = `feedback ${correct ? "correct" : "wrong"}`;
  els.feedback.textContent = correct
    ? "答對了"
    : `答錯了\n正確答案：${formatCorrectAnswer(current)}`;
  if (current.explanation) {
    els.explanation.className = "explanation";
    els.explanation.textContent = current.explanation;
  }
  document.querySelectorAll(".choice").forEach((button) => {
    button.disabled = true;
  });
  els.nextButton.disabled = false;
  updateCounts();
}

function isCorrectOption(selectedIndex, answer) {
  const answerNumber = Number.parseInt(String(answer).trim(), 10);
  return Number.isInteger(answerNumber) && selectedIndex === answerNumber - 1;
}

function isCorrectText(userAnswer, answer) {
  const normalize = (value) => String(value).trim().replace(/\s+/g, "").toLowerCase();
  const acceptedAnswers = String(answer)
    .split(/[|；;]/)
    .map(normalize)
    .filter(Boolean);
  return acceptedAnswers.includes(normalize(userAnswer));
}

function formatCorrectAnswer(question) {
  if (!question.options.length) return question.answer;
  const answerNumber = Number.parseInt(String(question.answer).trim(), 10);
  const option = question.options[answerNumber - 1];
  return option ? `${answerNumber}. ${option}` : question.answer;
}

function nextQuestion() {
  if (state.currentIndex >= state.questions.length - 1) {
    showResult();
    return;
  }

  state.currentIndex += 1;
  showQuestion();
}

function showResult() {
  const total = state.questions.length;
  const percent = total ? Math.round((state.score / total) * 100) : 0;
  els.quizCard.classList.add("hidden");
  els.resultCard.classList.remove("hidden");
  els.finalScore.textContent = `${percent}%`;
  els.finalDetail.textContent = `${state.score} / ${total}`;
  els.wrongList.innerHTML = "";

  const wrongAnswers = state.answers.filter((answer) => !answer.correct);
  if (!wrongAnswers.length) {
    els.wrongList.innerHTML = '<div class="wrong-item"><strong>全部答對</strong><span>這次很穩。</span></div>';
  } else {
    for (const item of wrongAnswers) {
      const div = document.createElement("div");
      div.className = "wrong-item";
      div.innerHTML = `
        <strong>${escapeHtml(item.question)}</strong>
        <div>你的答案：${escapeHtml(item.userAnswer)}</div>
        <div>正確答案：${escapeHtml(item.answer)}</div>
      `;
      els.wrongList.append(div);
    }
  }

  updateCounts();
}

function showEmpty(message) {
  els.quizCard.classList.add("hidden");
  els.resultCard.classList.add("hidden");
  els.emptyState.classList.remove("hidden");
  els.emptyState.querySelector("h2").textContent = message;
}

function updateCounts() {
  els.questionCount.textContent = `${state.rows.length} 題`;
  els.sidebarQuestionCount.textContent = String(state.rows.length);
  els.scoreSummary.textContent = state.questions.length
    ? `得分 ${state.score} / ${state.questions.length}`
    : "尚未開始";
}

function setStatus(message, isError = false) {
  els.statusLine.textContent = message;
  els.statusLine.classList.toggle("hidden", !message);
  els.statusLine.classList.toggle("error", isError);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
