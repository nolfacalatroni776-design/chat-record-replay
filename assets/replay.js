const A = window.MessageRecordAnalytics;

const state = {
  analysis: null,
  selectedProjectId: "",
  selectedGroupName: "",
  combo: {
    project: { options: [], query: "", activeIndex: 0 },
    group: { options: [], query: "", activeIndex: 0 }
  }
};

const elements = {
  fileInput: document.getElementById("fileInput"),
  loadDemoBtn: document.getElementById("loadDemoBtn"),
  dropzone: document.getElementById("dropzone"),
  loadState: document.getElementById("loadState"),
  fileMeta: document.getElementById("fileMeta"),
  restoreBtn: document.getElementById("restoreBtn"),
  toast: document.getElementById("toast"),
  projectCombo: document.getElementById("projectCombo"),
  projectComboButton: document.getElementById("projectComboButton"),
  projectComboLabel: document.getElementById("projectComboLabel"),
  projectComboPopover: document.getElementById("projectComboPopover"),
  projectComboSearch: document.getElementById("projectComboSearch"),
  projectComboOptions: document.getElementById("projectComboOptions"),
  groupCombo: document.getElementById("groupCombo"),
  groupComboButton: document.getElementById("groupComboButton"),
  groupComboLabel: document.getElementById("groupComboLabel"),
  groupComboPopover: document.getElementById("groupComboPopover"),
  groupComboSearch: document.getElementById("groupComboSearch"),
  groupComboOptions: document.getElementById("groupComboOptions")
};

function $(id) {
  return document.getElementById(id);
}

function safeText(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#039;"
  }[char]));
}

function shortText(value, max = 150) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return text.length > max ? `${text.slice(0, max - 1)}...` : text;
}

function normalize(value) {
  return String(value || "").trim();
}

function normalizeProjectId(value) {
  return normalize(value);
}

function groupName(record) {
  return record["群名"] || "未识别群";
}

function sameDay(a, b) {
  return a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function formatChatDate(date, fallback) {
  if (!date) return fallback || "-";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  const second = String(date.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
}

function formatChatTime(date, fallback) {
  if (!date) {
    const match = String(fallback || "").match(/(\d{1,2}:\d{2}(?::\d{2})?)$/);
    return match ? match[1] : "-";
  }
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  const second = String(date.getSeconds()).padStart(2, "0");
  return `${hour}:${minute}:${second}`;
}

function avatarColor(name) {
  let hash = 0;
  const text = String(name || "?");
  for (let index = 0; index < text.length; index += 1) {
    hash = (hash * 31 + text.charCodeAt(index)) % 360;
  }
  return `hsl(${hash}, 72%, 28%)`;
}

function showToast(message) {
  elements.toast.textContent = message;
  elements.toast.classList.add("show");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => elements.toast.classList.remove("show"), 2600);
}

function setLoadState(text, isGood = false) {
  elements.loadState.textContent = text;
  elements.loadState.classList.toggle("good", isGood);
}

function readWorkbookFromBuffer(buffer, fileName) {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: false, dense: false });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const originalRef = worksheet["!ref"] || "";
  const repairedRef = A.repairWorksheetRange(XLSX, worksheet);
  const rows = A.rowsFromWorksheet(XLSX, worksheet);
  const parsed = A.recordsFromRows(rows);
  if (!parsed.headerOk) {
    throw new Error(`表头不匹配：缺少 ${parsed.headerMissing.join("、")}`);
  }
  return A.analyzeRecords(parsed.records, {
    fileName,
    sheetName,
    originalRef,
    repairedRef,
    rowsRead: rows.length
  });
}

function loadAnalysis(analysis) {
  state.analysis = analysis;
  state.selectedProjectId = "";
  state.selectedGroupName = "";
  elements.fileMeta.textContent = `${analysis.source.fileName || "消息记录"} · ${A.formatNumber(analysis.totals.messages)} 条消息 · ${A.formatNumber(analysis.totals.projects)} 个项目`;
  renderProjectOptions();
  selectProject(state.combo.project.options[0] && state.combo.project.options[0].value, { silent: true });
  renderReplay();
}

async function loadFile(file) {
  if (!file) return;
  try {
    setLoadState("解析中");
    const buffer = await file.arrayBuffer();
    loadAnalysis(readWorkbookFromBuffer(buffer, file.name));
    setLoadState("已加载", true);
    showToast(`已解析 ${A.formatNumber(state.analysis.totals.messages)} 条消息。`);
  } catch (error) {
    setLoadState("解析失败");
    showToast(error.message);
  }
}

async function loadDemo() {
  try {
    setLoadState("加载演示");
    const response = await fetch("data/demo-records.json");
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const records = await response.json();
    const analysis = A.analyzeRecords(records, {
      fileName: "脱敏演示数据",
      sheetName: "demo-records.json",
      rowsRead: records.length
    });
    loadAnalysis(analysis);
    setLoadState("演示已加载", true);
    showToast("脱敏演示数据已加载。");
  } catch (error) {
    setLoadState("演示不可用");
    showToast(`演示加载失败：${error.message}`);
  }
}

function recordsForProject(projectId) {
  const target = normalizeProjectId(projectId).toLowerCase();
  if (!state.analysis || !target) return [];
  return state.analysis.records.filter((record) => normalizeProjectId(record["项目ID"]).toLowerCase() === target);
}

function recordsForCurrentGroup() {
  return recordsForProject(state.selectedProjectId)
    .filter((record) => groupName(record) === state.selectedGroupName)
    .slice()
    .sort((a, b) => {
      const at = a.parsedDate ? a.parsedDate.getTime() : Number.POSITIVE_INFINITY;
      const bt = b.parsedDate ? b.parsedDate.getTime() : Number.POSITIVE_INFINITY;
      return at - bt || a.__rowNumber - b.__rowNumber;
    });
}

function groupCounts(records) {
  const map = new Map();
  records.forEach((record) => {
    const name = groupName(record);
    if (!map.has(name)) {
      map.set(name, { name, messages: 0, users: new Set(), firstDate: null, lastDate: null });
    }
    const item = map.get(name);
    item.messages += 1;
    if (record["发送用户"]) item.users.add(record["发送用户"]);
    if (record.parsedDate) {
      if (!item.firstDate || record.parsedDate < item.firstDate) item.firstDate = record.parsedDate;
      if (!item.lastDate || record.parsedDate > item.lastDate) item.lastDate = record.parsedDate;
    }
  });
  return Array.from(map.values())
    .map((item) => ({ ...item, users: item.users.size }))
    .sort((a, b) => b.messages - a.messages || a.name.localeCompare(b.name));
}

function renderProjectOptions() {
  if (!state.analysis) {
    setComboOptions("project", []);
    return;
  }
  const options = state.analysis.projectStats.map((project) => ({
    value: project.name,
    label: project.name,
    meta: `${A.formatNumber(project.messages)} 条 · ${A.formatNumber(project.groups)} 群 · ${A.formatNumber(project.users)} 用户`,
    search: `${project.name} ${project.messages} ${project.groups} ${project.users}`
  }));
  setComboOptions("project", options);
}

function renderGroupOptions(projectId) {
  const groups = groupCounts(recordsForProject(projectId));
  const options = groups.map((group) => ({
    value: group.name,
    label: group.name,
    meta: `${A.formatNumber(group.messages)} 条 · ${A.formatNumber(group.users)} 用户 · ${formatDateRange(group.firstDate, group.lastDate)}`,
    search: `${group.name} ${group.messages} ${group.users}`
  }));
  setComboOptions("group", options);
}

function formatDateRange(start, end) {
  if (!start || !end) return "-";
  const startText = formatChatDate(start).slice(0, 10);
  const endText = formatChatDate(end).slice(0, 10);
  return startText === endText ? startText : `${startText} 至 ${endText}`;
}

function setComboOptions(type, options) {
  state.combo[type].options = options;
  state.combo[type].query = "";
  state.combo[type].activeIndex = 0;
  elements[`${type}ComboSearch`].value = "";
  renderComboOptions(type);
}

function filteredComboOptions(type) {
  const comboState = state.combo[type];
  const query = comboState.query.trim().toLowerCase();
  if (!query) return comboState.options;
  return comboState.options.filter((option) => option.search.toLowerCase().includes(query));
}

function renderComboOptions(type) {
  const container = elements[`${type}ComboOptions`];
  const selectedValue = type === "project" ? state.selectedProjectId : state.selectedGroupName;
  const options = filteredComboOptions(type);
  if (!options.length) {
    container.innerHTML = "<div class=\"combo-empty\">没有匹配选项</div>";
    return;
  }
  state.combo[type].activeIndex = Math.min(state.combo[type].activeIndex, options.length - 1);
  container.innerHTML = options.map((option, index) => `
    <button class="combo-option${option.value === selectedValue || index === state.combo[type].activeIndex ? " active" : ""}" type="button" role="option" data-type="${type}" data-value="${safeText(option.value)}" aria-selected="${option.value === selectedValue ? "true" : "false"}">
      <strong>${safeText(option.label)}</strong>
      <em>${safeText(option.meta)}</em>
    </button>
  `).join("");
}

function setComboLabel(type, text) {
  elements[`${type}ComboLabel`].textContent = text || elements[`${type}Combo`].dataset.placeholder;
}

function openCombo(type) {
  closeCombo(type === "project" ? "group" : "project");
  elements[`${type}Combo`].classList.add("open");
  elements[`${type}ComboButton`].setAttribute("aria-expanded", "true");
  elements[`${type}ComboSearch`].focus();
  elements[`${type}ComboSearch`].select();
  renderComboOptions(type);
}

function closeCombo(type) {
  elements[`${type}Combo`].classList.remove("open");
  elements[`${type}ComboButton`].setAttribute("aria-expanded", "false");
}

function selectComboOption(type, value) {
  if (type === "project") {
    selectProject(value);
  } else {
    selectGroup(value);
  }
  closeCombo(type);
}

function selectProject(projectId, options = {}) {
  const value = normalizeProjectId(projectId);
  state.selectedProjectId = value;
  const option = state.combo.project.options.find((item) => item.value === value);
  setComboLabel("project", option ? option.label : "先加载数据");
  renderGroupOptions(value);
  const nextGroup = state.combo.group.options[0] && state.combo.group.options[0].value;
  selectGroup(nextGroup || "", { silent: true });
  if (!options.silent) renderReplay();
}

function selectGroup(group, options = {}) {
  state.selectedGroupName = group || "";
  const option = state.combo.group.options.find((item) => item.value === state.selectedGroupName);
  setComboLabel("group", option ? option.label : "先选择项目ID");
  renderComboOptions("group");
  if (!options.silent) renderReplay();
}

function handleComboKeydown(type, event) {
  const options = filteredComboOptions(type);
  if (event.key === "ArrowDown") {
    event.preventDefault();
    state.combo[type].activeIndex = options.length ? (state.combo[type].activeIndex + 1) % options.length : 0;
    renderComboOptions(type);
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    state.combo[type].activeIndex = options.length ? (state.combo[type].activeIndex - 1 + options.length) % options.length : 0;
    renderComboOptions(type);
  } else if (event.key === "Enter") {
    event.preventDefault();
    if (options[state.combo[type].activeIndex]) {
      selectComboOption(type, options[state.combo[type].activeIndex].value);
    }
  } else if (event.key === "Escape") {
    closeCombo(type);
    elements[`${type}ComboButton`].focus();
  }
}

function renderReplay() {
  if (!state.analysis) {
    $("chatTitle").textContent = "未选择群";
    $("chatSubtitle").textContent = "上传 Excel 或加载脱敏演示后开始还原";
    $("chatTranscript").innerHTML = "<div class=\"chat-empty\">上传 Excel 或加载脱敏演示后开始还原</div>";
    updateSummary([], []);
    return;
  }
  if (!state.selectedProjectId || !state.selectedGroupName) {
    $("chatTitle").textContent = "未选择群";
    $("chatSubtitle").textContent = "请选择项目ID和群名";
    $("chatTranscript").innerHTML = "<div class=\"chat-empty\">请选择项目ID和群名</div>";
    updateSummary(recordsForProject(state.selectedProjectId), []);
    return;
  }
  const projectRows = recordsForProject(state.selectedProjectId);
  const rows = recordsForCurrentGroup();
  const users = new Set(rows.map((record) => record["发送用户"]).filter(Boolean));
  $("chatTitle").textContent = state.selectedGroupName;
  $("chatSubtitle").textContent = `${state.selectedProjectId} · 当前群 ${A.formatNumber(rows.length)} 条 · 项目覆盖 ${A.formatNumber(groupCounts(projectRows).length)} 群 · ${A.formatNumber(users.size)} 用户`;
  $("chatTranscript").innerHTML = renderChatMessages(rows);
  updateSummary(projectRows, rows);
}

function updateSummary(projectRows, groupRows) {
  const users = new Set(groupRows.map((record) => record["发送用户"]).filter(Boolean));
  const dates = groupRows.map((record) => record.parsedDate).filter(Boolean).sort((a, b) => a - b);
  $("projectCount").textContent = projectRows.length ? A.formatNumber(projectRows.length) : "-";
  $("groupCount").textContent = groupRows.length ? A.formatNumber(groupRows.length) : "-";
  $("userCount").textContent = users.size ? A.formatNumber(users.size) : "-";
  $("dateRange").textContent = dates.length ? formatDateRange(dates[0], dates[dates.length - 1]) : "-";
}

function renderChatMessages(rows) {
  if (!rows.length) return "<div class=\"chat-empty\">当前群没有消息</div>";
  let previousDate = null;
  return rows.map((record, index) => {
    const separator = !previousDate || !sameDay(previousDate, record.parsedDate)
      ? `<div class="chat-time-separator">${safeText(formatChatDate(record.parsedDate, record["发送时间"]))}</div>`
      : "";
    previousDate = record.parsedDate || previousDate;
    const user = record["发送用户"] || "unknown";
    const initial = safeText(user.slice(0, 1).toLowerCase() || "?");
    const roleClass = record["用户角色"] === "IM_PM" ? " pm" : "";
    return `
      ${separator}
      <article class="chat-message${roleClass}" style="--avatar-bg:${avatarColor(user)}" data-row="${record.__rowNumber || index + 1}">
        <div class="chat-avatar">${initial}</div>
        <div class="chat-message-main">
          <div class="chat-meta">
            <strong>${safeText(user)}</strong>
            <span>${safeText(formatChatTime(record.parsedDate, record["发送时间"]))}</span>
          </div>
          ${renderChatBubble(record, index)}
        </div>
      </article>`;
  }).join("");
}

function renderChatBubble(record, index) {
  const content = shortText(record.text || record["消息内容"], 1200);
  const images = (record.imageUrls || []).slice(0, 4);
  const links = (record.urls || []).filter((url) => !(record.imageUrls || []).includes(url)).slice(0, 3);
  const imageHtml = images.map((url) => `
    <a class="chat-image" href="${safeText(url)}" target="_blank" rel="noreferrer">
      <span>图片附件 ${safeText(String(index + 1))}</span>
    </a>
  `).join("");
  const linkHtml = links.map((url) => `
    <a class="chat-link" href="${safeText(url)}" target="_blank" rel="noreferrer">${safeText(shortText(url, 90))}</a>
  `).join("");
  const fallback = images.length ? "图片附件" : record.contentFormat;
  return `
    <div class="chat-bubble">
      ${content ? `<p>${safeText(content)}</p>` : `<p class="chat-placeholder">${safeText(fallback)}</p>`}
      ${imageHtml ? `<div class="chat-images">${imageHtml}</div>` : ""}
      ${linkHtml ? `<div class="chat-links">${linkHtml}</div>` : ""}
    </div>`;
}

elements.fileInput.addEventListener("change", (event) => loadFile(event.target.files[0]));
elements.loadDemoBtn.addEventListener("click", loadDemo);
elements.restoreBtn.addEventListener("click", renderReplay);

["project", "group"].forEach((type) => {
  elements[`${type}ComboButton`].addEventListener("click", () => openCombo(type));
  elements[`${type}ComboSearch`].addEventListener("input", (event) => {
    state.combo[type].query = event.target.value;
    state.combo[type].activeIndex = 0;
    renderComboOptions(type);
  });
  elements[`${type}ComboSearch`].addEventListener("keydown", (event) => handleComboKeydown(type, event));
  elements[`${type}ComboOptions`].addEventListener("click", (event) => {
    const option = event.target.closest(".combo-option");
    if (!option) return;
    selectComboOption(option.dataset.type, option.dataset.value);
  });
});

document.addEventListener("click", (event) => {
  if (!elements.projectCombo.contains(event.target)) closeCombo("project");
  if (!elements.groupCombo.contains(event.target)) closeCombo("group");
});

["dragenter", "dragover"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("dragging");
  });
});

elements.dropzone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (file) loadFile(file);
});

if (new URLSearchParams(window.location.search).get("demo") === "1") {
  loadDemo();
}
