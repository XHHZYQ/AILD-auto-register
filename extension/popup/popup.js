const defaultState = {
  status: "idle",
  startRow: 2,
  pauseAfter: 1,
  completedCount: 0,
  currentExcelRow: null,
  currentStudent: null,
  message: "",
};

const fields = {
  statusText: document.querySelector("#statusText"),
  startRow: document.querySelector("#startRow"),
  pauseAfter: document.querySelector("#pauseAfter"),
  studentName: document.querySelector("#studentName"),
  teacherName: document.querySelector("#teacherName"),
  studentStatus: document.querySelector("#studentStatus"),
  message: document.querySelector("#message"),
  startButton: document.querySelector("#startButton"),
  pauseButton: document.querySelector("#pauseButton"),
  resumeButton: document.querySelector("#resumeButton"),
};

let state = { ...defaultState };

init();

async function init() {
  state = { ...defaultState, ...(await chrome.storage.local.get(defaultState)) };
  render();

  fields.startButton.addEventListener("click", () => sendCommand("start"));
  fields.pauseButton.addEventListener("click", () => sendCommand("pause"));
  fields.resumeButton.addEventListener("click", () => sendCommand("resume"));

  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== "local") return;
    for (const [key, change] of Object.entries(changes)) {
      state[key] = change.newValue;
    }
    render();
  });
}

async function sendCommand(command) {
  const startRow = normalizePositiveInteger(fields.startRow.value, 2);
  const pauseAfter = normalizePositiveInteger(fields.pauseAfter.value, 1);
  await chrome.storage.local.set({ startRow, pauseAfter });

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    await chrome.storage.local.set({ message: "未找到当前活动页面。" });
    return;
  }

  await chrome.tabs.sendMessage(tab.id, {
    source: "aild-auto-register-popup",
    command,
    payload: { startRow, pauseAfter },
  });
}

function normalizePositiveInteger(value, fallback) {
  const number = Number.parseInt(value, 10);
  return Number.isFinite(number) && number > 0 ? number : fallback;
}

function render() {
  fields.startRow.value = state.startRow;
  fields.pauseAfter.value = state.pauseAfter;
  fields.statusText.textContent = statusLabel(state.status);
  fields.studentName.textContent = state.currentStudent?.name || "-";
  fields.teacherName.textContent = state.currentStudent?.teacherName || "-";
  fields.studentStatus.textContent = state.currentStudent?.status || "等待执行";
  fields.message.textContent = state.message || "";

  fields.startButton.disabled = state.status === "running";
  fields.pauseButton.disabled = state.status !== "running";
  fields.resumeButton.disabled = state.status !== "paused";
}

function statusLabel(status) {
  const labels = {
    idle: "未开始",
    running: "执行中",
    paused: "已暂停",
    done: "已完成",
    error: "执行异常",
  };
  return labels[status] || status;
}
