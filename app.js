const state = {
  rows: [],
  questions: [],
  currentIndex: 0,
  score: 0,
  selectedChoice: "",
  answers: [],
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  statusLine: document.querySelector("#statusLine"),
  questionColumn: document.querySelector("#questionColumn"),
  answerColumn: document.querySelector("#answerColumn"),
  optionColumn: document.querySelector("#optionColumn"),
  optionColumnA: document.querySelector("#optionColumnA"),
  optionColumnB: document.querySelector("#optionColumnB"),
  optionColumnC: document.querySelector("#optionColumnC"),
  optionColumnD: document.querySelector("#optionColumnD"),
  categoryColumn: document.querySelector("#categoryColumn"),
  modeSelect: document.querySelector("#modeSelect"),
  limitInput: document.querySelector("#limitInput"),
  categoryFilter: document.querySelector("#categoryFilter"),
  startButton: document.querySelector("#startButton"),
  questionCount: document.querySelector("#questionCount"),
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
  submitButton: document.querySelector("#submitButton"),
  nextButton: document.querySelector("#nextButton"),
  finalScore: document.querySelector("#finalScore"),
  finalDetail: document.querySelector("#finalDetail"),
  wrongList: document.querySelector("#wrongList"),
  restartButton: document.querySelector("#restartButton"),
};

const emptyOption = { label: "不使用", value: "" };
const aliases = {
  question: ["題目", "問題", "question", "q", "題幹", "題庫"],
  answer: ["答案", "正解", "answer", "a", "解答", "正確答案"],
  options: ["選項", "choices", "options", "choice", "選擇題選項"],
  optionA: ["選項a", "選項1", "choicea", "choice1", "optiona", "option1", "a"],
  optionB: ["選項b", "選項2", "choiceb", "choice2", "optionb", "option2", "b"],
  optionC: ["選項c", "選項3", "choicec", "choice3", "optionc", "option3", "c"],
  optionD: ["選項d", "選項4", "choiced", "choice4", "optiond", "option4", "d"],
  category: ["分類", "類別", "category", "章節", "單元"],
};

els.fileInput.addEventListener("change", handleFile);
els.startButton.addEventListener("click", startQuiz);
els.submitButton.addEventListener("click", submitAnswer);
els.nextButton.addEventListener("click", nextQuestion);
els.restartButton.addEventListener("click", startQuiz);

for (const select of [
  els.questionColumn,
  els.answerColumn,
  els.optionColumn,
  els.optionColumnA,
  els.optionColumnB,
  els.optionColumnC,
  els.optionColumnD,
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
  setOptions(els.optionColumn, [emptyOption, ...columns], guessColumn(columns, aliases.options));
  setOptions(els.optionColumnA, [emptyOption, ...columns], guessColumn(columns, aliases.optionA));
  setOptions(els.optionColumnB, [emptyOption, ...columns], guessColumn(columns, aliases.optionB));
  setOptions(els.optionColumnC, [emptyOption, ...columns], guessColumn(columns, aliases.optionC));
  setOptions(els.optionColumnD, [emptyOption, ...columns], guessColumn(columns, aliases.optionD));
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
}

function canStart() {
  return state.rows.length > 0 && els.questionColumn.value && els.answerColumn.value;
}

function startQuiz() {
  if (!canStart()) return;

  const categoryColumn = els.categoryColumn.value;
  const category = els.categoryFilter.value;
  const questionColumn = els.questionColumn.value;
  const answerColumn = els.answerColumn.value;
  const optionColumn = els.optionColumn.value;
  const optionColumns = [
    els.optionColumnA.value,
    els.optionColumnB.value,
    els.optionColumnC.value,
    els.optionColumnD.value,
  ].filter(Boolean);

  let pool = state.rows
    .filter((row) => row[questionColumn] && row[answerColumn])
    .filter((row) => !category || row[categoryColumn] === category)
    .map((row) => ({
      question: row[questionColumn],
      answer: row[answerColumn],
      options: collectOptions(row, optionColumn, optionColumns),
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

function parseOptions(value) {
  if (!value) return [];
  return String(value)
    .split(/\r?\n|[|；;]/)
    .map((option) => option.replace(/^[A-Z]\s*[.)、:：]\s*/i, "").trim())
    .filter(Boolean);
}

function collectOptions(row, combinedColumn, separateColumns) {
  const separateOptions = separateColumns.map((column) => row[column]).filter(Boolean);
  if (separateOptions.length > 1) {
    return separateOptions;
  }

  const combinedOptions = parseOptions(row[combinedColumn]);
  if (combinedOptions.length > 1) {
    return combinedOptions;
  }

  return separateOptions.length ? separateOptions : combinedOptions;
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
  els.answerInput.value = "";
  els.answerInput.classList.toggle("hidden", current.options.length > 0);
  els.submitButton.disabled = false;
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
      state.selectedChoice = choice;
      document.querySelectorAll(".choice").forEach((item) => item.classList.remove("selected"));
      button.classList.add("selected");
    });
    els.choiceList.append(button);
  });
}

function submitAnswer() {
  const current = state.questions[state.currentIndex];
  const userAnswer = current.options.length ? state.selectedChoice : els.answerInput.value.trim();
  if (!userAnswer) return;

  const correct = isCorrect(userAnswer, current.answer);
  if (correct) state.score += 1;

  state.answers.push({
    ...current,
    userAnswer,
    correct,
  });

  els.feedback.className = `feedback ${correct ? "correct" : "wrong"}`;
  els.feedback.textContent = correct ? "答對了" : `答錯了\n正確答案：${current.answer}`;
  els.submitButton.disabled = true;
  els.nextButton.disabled = false;
  updateCounts();
}

function isCorrect(userAnswer, answer) {
  const normalize = (value) => String(value).trim().replace(/\s+/g, "").toLowerCase();
  const acceptedAnswers = String(answer)
    .split(/[|；;]/)
    .map(normalize)
    .filter(Boolean);
  return acceptedAnswers.includes(normalize(userAnswer));
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
