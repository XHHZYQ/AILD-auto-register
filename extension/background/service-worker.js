chrome.runtime.onInstalled.addListener(async () => {
  await chrome.storage.local.set({
    status: "idle",
    startRow: 2,
    pauseAfter: 1,
    completedCount: 0,
    currentExcelRow: null,
    currentStudent: null,
    message: "插件已就绪，请在报名页面打开面板开始执行。",
  });
});
