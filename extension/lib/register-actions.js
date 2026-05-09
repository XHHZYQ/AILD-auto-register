(function exposeRegisterActions(global) {
  async function submitRegistrationPreview(options = {}) {
    const { signal } = options;

    const previewButton = await waitUntil(() => findButtonByText("预览确认"), {
      signal,
      timeout: 15000,
      errorMessage: "未找到预览确认按钮。",
      returnValue: true,
    });
    previewButton.click();

    await waitUntil(() => {
      const preview = [...document.querySelectorAll(".newPre, .form-container")]
        .filter(isVisibleElement)
        .find((node) => normalizeText(node.textContent).includes("初赛报名表"));
      const submitButton = findButtonByText("提交报名");
      return preview && submitButton;
    }, {
      signal,
      timeout: 20000,
      errorMessage: "等待初赛报名表预览页渲染超时。",
    });

    const submitButton = await waitUntil(() => findButtonByText("提交报名"), {
      signal,
      timeout: 10000,
      errorMessage: "未找到提交报名按钮。",
      returnValue: true,
    });
    submitButton.click();

    const successDialog = await waitUntil(() => (
      findMessageBoxByText("报名成功") || findMessageBoxByText("你已完成报名")
    ), {
      signal,
      timeout: 30000,
      errorMessage: "等待报名成功弹窗超时。",
      returnValue: true,
    });
    clickMessageBoxConfirm(successDialog);

    await waitUntil(() => !findMessageBoxByText("报名成功") && !findMessageBoxByText("你已完成报名"), {
      signal,
      timeout: 10000,
      errorMessage: "等待报名成功弹窗关闭超时。",
    });

    const backButton = await waitUntil(() => findButtonByText("返回") || findButtonByText("返回修改"), {
      signal,
      timeout: 15000,
      errorMessage: "未找到返回按钮。",
      returnValue: true,
    });
    backButton.click();

    await waitUntil(() => (
      normalizeText(document.body.textContent).includes("参赛信息")
      && normalizeText(document.body.textContent).includes("添加队员")
    ), {
      signal,
      timeout: 20000,
      errorMessage: "等待返回报名主页面超时。",
    });
  }

  async function acceptCommitmentDialog(options = {}) {
    const { signal, timeout = 15000 } = options;
    const dialog = await waitUntil(() => findCommitmentDialog(), {
      signal,
      timeout,
      errorMessage: "等待承诺书弹窗超时。",
      returnValue: true,
    });

    const checkboxInput = dialog.querySelector('input[type="checkbox"]');
    const checkbox = dialog.querySelector(".el-checkbox") || checkboxInput;
    const checked = checkboxInput?.checked
      || checkbox?.classList?.contains("is-checked")
      || checkbox?.getAttribute("aria-checked") === "true";

    if (!checked) {
      (checkbox || checkboxInput)?.click();
      await delay(120);
    }

    const confirmButton = [...dialog.querySelectorAll("button")]
      .find((button) => normalizeText(button.textContent).includes("确认"))
      || dialog.querySelector(".memOk")
      || dialog.querySelector("button");
    if (!confirmButton) throw new Error("承诺书弹窗中未找到确认按钮。");
    confirmButton.click();

    await waitUntil(() => !findCommitmentDialog(), {
      signal,
      timeout: 10000,
      errorMessage: "等待承诺书弹窗关闭超时。",
    });
  }

  function findButtonByText(text) {
    const normalizedText = normalizeText(text);
    return [...document.querySelectorAll("button")]
      .filter(isVisibleElement)
      .find((button) => normalizeText(button.textContent).includes(normalizedText)) || null;
  }

  function findMessageBoxByText(text) {
    const normalizedText = normalizeText(text);
    return [...document.querySelectorAll(".el-message-box__wrapper, .el-message-box")]
      .filter(isVisibleElement)
      .find((dialog) => normalizeText(dialog.textContent).includes(normalizedText)) || null;
  }

  function clickMessageBoxConfirm(dialog) {
    const button = [...dialog.querySelectorAll("button")]
      .find((candidate) => normalizeText(candidate.textContent).includes("确定"))
      || dialog.querySelector(".el-button--primary")
      || dialog.querySelector("button");
    button?.click();
  }

  function findCommitmentDialog() {
    return [...document.querySelectorAll(".el-dialog__wrapper, .el-dialog")]
      .filter(isVisibleElement)
      .find((dialog) => {
        const text = normalizeText(dialog.textContent);
        return text.includes("承诺书") || text.includes("我已阅读，继续报名");
      }) || null;
  }

  async function waitUntil(predicate, options = {}) {
    const { timeout = 15000, signal, errorMessage = "等待条件超时。", returnValue = false } = options;
    const start = Date.now();

    while (Date.now() - start < timeout) {
      if (signal?.aborted) throw new Error("执行已暂停");
      const result = predicate();
      if (result) return returnValue ? result : true;
      await delay(120);
    }

    throw new Error(errorMessage);
  }

  function isVisibleElement(element) {
    if (!element) return false;
    const style = window.getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") return false;
    const rect = element.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function normalizeText(value) {
    return String(value || "").replace(/\s+/g, "");
  }

  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  global.AILDRegisterActions = {
    acceptCommitmentDialog,
    submitRegistrationPreview,
  };
})(window);
