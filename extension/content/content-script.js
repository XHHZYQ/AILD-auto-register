const runtimeState = {
  status: "idle",
  controller: null,
  registrationData: null,
  currentExcelRow: null,
};

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.source !== "aild-auto-register-popup") return false;

  handleCommand(message.command, message.payload).then(
    () => sendResponse({ ok: true }),
    async (error) => {
      await updateUiState({
        status: "error",
        message: error?.message || "执行失败",
      });
      sendResponse({ ok: false, error: error?.message || String(error) });
    },
  );
  return true;
});

async function handleCommand(command, payload) {
  if (command === "pause") {
    runtimeState.status = "paused";
    runtimeState.controller?.abort();
    await updateUiState({ status: "paused", message: "已暂停，当前学生信息保留在页面。" });
    return;
  }

  if (command === "resume") {
    await startFlow(payload, true);
    return;
  }

  if (command === "start") {
    await startFlow(payload, false);
  }
}

async function startFlow(payload, isResume) {
  runtimeState.status = "running";
  runtimeState.controller = new AbortController();
  runtimeState.currentExcelRow = isResume && runtimeState.currentExcelRow ? runtimeState.currentExcelRow : payload.startRow;

  await updateUiState({
    status: "running",
    startRow: payload.startRow,
    pauseAfter: payload.pauseAfter,
    message: isResume ? "继续执行，正在准备读取报名数据。" : "开始执行，正在准备读取报名数据。",
  });

  await waitForElement(".form-container", { text: "参赛信息", signal: runtimeState.controller.signal });
  const data = await loadRegistrationData();
  const student = window.AILDExcelData.getStudentByExcelRow(data, runtimeState.currentExcelRow);

  if (!student) {
    await updateUiState({
      status: "done",
      currentStudent: null,
      message: `Excel 第 ${runtimeState.currentExcelRow} 行没有可报名学生，执行结束。`,
    });
    runtimeState.status = "done";
    return;
  }

  const teacher = window.AILDExcelData.findTeacher(data, student.teacherName);
  const quickTeacher = await findQuickTeacherButton(student.teacherName, {
    signal: runtimeState.controller.signal,
  });
  let teacherFormFilled = false;

  if (quickTeacher.exists) {
    quickTeacher.button.click();
    await delay(200);
  } else if (teacher) {
    await updateUiState({
      status: "running",
      currentStudent: {
        name: student.name,
        teacherName: student.teacherName || "-",
        status: "正在填写指导老师信息",
        excelRowNumber: student.excelRowNumber,
      },
      message: `页面未找到 ${student.teacherName}，正在填写指导老师信息表单。`,
    });
    await fillTeacherForm(teacher, { signal: runtimeState.controller.signal });
    teacherFormFilled = true;
  }

  await updateUiState({
    status: "running",
    currentStudent: {
      name: student.name,
      teacherName: student.teacherName || "-",
      status: "正在填写参赛信息",
      excelRowNumber: student.excelRowNumber,
    },
    message: `正在填写 ${student.name} 的参赛信息。`,
  });
  const competitionInfo = await fillCompetitionForm(student, {
    signal: runtimeState.controller.signal,
  });

  await updateUiState({
    status: "paused",
    currentStudent: {
      name: student.name,
      teacherName: student.teacherName || "-",
      status: `${buildTeacherDetectionStatus(teacher, quickTeacher.exists, teacherFormFilled)}；参赛信息已填写`,
      excelRowNumber: student.excelRowNumber,
      teamName: competitionInfo.teamName,
    },
    message: `${buildTeacherDetectionMessage(student, teacher, quickTeacher, teacherFormFilled)} 团队名称：${competitionInfo.teamName}。下一步将添加队员。`,
  });

  runtimeState.status = "paused";
}

async function fillCompetitionForm(student, options = {}) {
  const { signal } = options;
  const form = await findSectionForm(["参赛信息"], {
    signal,
    timeout: 15000,
  });

  await selectByLabel(form, ["项目名称"], "智能算法编程", { signal });
  await selectByLabel(form, ["场景名称"], "智械未来", {
    signal,
    filterText: "智械未来",
  });
  await selectByLabel(form, ["队员组别"], student.groupName, { signal });
  await selectByLabel(form, ["学校所在地区"], provinceCandidates(student.province), {
    signal,
    selectIndex: 0,
  });
  await selectByLabel(form, ["学校所在地区"], cityCandidates(student.province, student.city), {
    signal,
    selectIndex: 1,
  });

  const teamName = await fillUniqueTeamName(form, student, { signal });
  return { teamName };
}

async function fillTeacherForm(teacher, options = {}) {
  const { signal } = options;
  const form = await findSectionForm(["指导教师信息", "指导老师信息"], {
    signal,
    timeout: 15000,
  });

  await fillInputByLabel(form, ["姓名中文", "中文姓名", "姓名"], teacher.name, { signal });
  await selectByLabel(form, ["性别"], teacher.gender, { signal });
  await fillInputByLabel(form, ["民族"], teacher.nation, { signal });
  await fillInputByLabel(form, ["证件号码", "身份证号"], teacher.idNumber, {
    signal,
    inputSelector: 'input.am-form-field:not([style*="display: none"])',
  });
  await fillInputByLabel(form, ["手机号", "手机"], teacher.phone, { signal });
  await fillInputByLabel(form, ["邮箱"], teacher.email, { signal });
  await fillInputByLabel(form, ["学校全称", "学校", "单位"], teacher.workplace, { signal, optional: true });

  await uploadByLabel(form, ["一寸照片", "照片"], teacher.photo, {
    signal,
    optional: !teacher.photo?.hasImageFile,
  });
  await uploadByLabel(form, ["资格证书"], teacher.certificate, {
    signal,
    optional: true,
  });
}

async function findSectionForm(titleTexts, options = {}) {
  const { timeout = 15000, signal } = options;
  const titles = Array.isArray(titleTexts) ? titleTexts : [titleTexts];

  return waitUntil(() => {
    const title = [...document.querySelectorAll(".form-container .titleB, .form-container h3")]
      .find((node) => titles.some((text) => normalizeText(node.textContent).includes(normalizeText(text))));
    const section = title ? nextElementMatching(title, ".form-item") : null;
    return section?.querySelector("form") || null;
  }, {
    timeout,
    signal,
    errorMessage: `等待表单渲染超时：${titles.join(" / ")}`,
    returnValue: true,
  });
}

async function fillInputByLabel(scope, labels, value, options = {}) {
  const { signal, optional = false, inputSelector = "input.am-form-field, textarea" } = options;
  if (isBlank(value)) {
    if (optional) return null;
    throw new Error(`字段为空，无法填写：${labels}`);
  }

  const group = findFieldGroup(scope, labels);
  if (!group) {
    if (optional) return null;
    throw new Error(`未找到输入字段：${labels}`);
  }

  const input = [...group.querySelectorAll(inputSelector)].find(isVisibleElement);
  if (!input) {
    if (optional) return null;
    throw new Error(`未找到可见输入框：${labels}`);
  }

  setNativeValue(input, value);
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
  input.dispatchEvent(new Event("blur", { bubbles: true }));
  await delayWithAbort(80, signal);
  return input;
}

async function selectByLabel(scope, labels, value, options = {}) {
  const { signal, optional = false, selectIndex = 0, filterText = "" } = options;
  const values = Array.isArray(value) ? value.filter((item) => !isBlank(item)) : [value].filter((item) => !isBlank(item));
  if (!values.length) {
    if (optional) return null;
    throw new Error(`字段为空，无法选择：${labels}`);
  }

  const group = findFieldGroup(scope, labels);
  if (!group) {
    if (optional) return null;
    throw new Error(`未找到选择字段：${labels}`);
  }

  const selectInput = [...group.querySelectorAll(".el-select input.el-input__inner, input[readonly]")]
    .filter(isVisibleElement)[selectIndex];
  if (!selectInput) {
    if (optional) return null;
    throw new Error(`未找到选择框：${labels}`);
  }

  if (filterText) {
    setNativeValue(selectInput, filterText);
    selectInput.dispatchEvent(new Event("input", { bubbles: true }));
    await delayWithAbort(120, signal);
  }
  selectInput.click();
  await delayWithAbort(120, signal);

  const option = await waitUntil(() => findVisibleSelectOption(values), {
    signal,
    timeout: 10000,
    errorMessage: `等待选择项渲染超时：${values.join(" / ")}`,
    returnValue: true,
  });

  option.click();
  selectInput.dispatchEvent(new Event("change", { bubbles: true }));
  selectInput.dispatchEvent(new Event("blur", { bubbles: true }));
  await delayWithAbort(120, signal);
  return option;
}

async function fillUniqueTeamName(form, student, options = {}) {
  const { signal } = options;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const teamName = generateTeamName(student, attempt);
    await fillInputByLabel(form, ["团队名称"], teamName, { signal });
    const input = findFieldGroup(form, ["团队名称"])?.querySelector("input.am-form-field");
    input?.dispatchEvent(new Event("blur", { bubbles: true }));
    await delayWithAbort(500, signal);

    const duplicateDialog = findMessageBoxByText("团队名称已存在，请重新输入");
    if (!duplicateDialog) return teamName;

    clickMessageBoxConfirm(duplicateDialog);
    await waitUntil(() => !findMessageBoxByText("团队名称已存在，请重新输入"), {
      signal,
      timeout: 8000,
      errorMessage: "等待团队名称重复提示关闭超时。",
    });
  }

  throw new Error(`团队名称连续重复，无法为 ${student.name} 生成可用名称。`);
}

function generateTeamName(student, attempt = 0) {
  const source = `${student.idNumber || ""}${student.name || ""}${student.excelRowNumber || ""}${attempt}`;
  const chars = "智创星云启明远航新锐卓越未来灵动慧思";
  let hash = 0;
  for (const char of source) {
    hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  }

  let name = "";
  for (let index = 0; index < 4; index += 1) {
    name += chars[(hash + index * 7) % chars.length];
  }
  return name.replace(/[^\u4e00-\u9fa5]/g, "").slice(0, 5);
}

function provinceCandidates(province) {
  const normalized = normalizeText(province);
  const withoutSuffix = normalized.replace(/省$|市$|自治区$|壮族自治区$|回族自治区$|维吾尔自治区$/g, "");
  return [...new Set([normalized, withoutSuffix].filter(Boolean))];
}

function cityCandidates(province, city) {
  const normalizedProvince = normalizeText(province);
  const normalizedCity = normalizeText(city);
  const isDirectCity = ["北京市", "北京", "天津市", "天津", "上海市", "上海", "重庆市", "重庆"].includes(normalizedProvince);
  const directCity = normalizedProvince.replace(/市$/, "");
  return [...new Set([
    normalizedCity,
    normalizedCity.replace(/市$/, ""),
    isDirectCity ? `${directCity}市` : "",
    isDirectCity ? directCity : "",
    isDirectCity ? "市辖区" : "",
  ].filter(Boolean))];
}

async function uploadByLabel(scope, labels, imageInfo, options = {}) {
  const { signal, optional = false } = options;
  if (!imageInfo?.hasImageFile || !imageInfo.bytes) {
    if (optional) return null;
    throw new Error(`缺少可上传图片：${labels}`);
  }

  const group = findFieldGroup(scope, labels);
  if (!group) {
    if (optional) return null;
    throw new Error(`未找到上传字段：${labels}`);
  }

  const input = group.querySelector('input[type="file"]');
  if (!input) {
    if (optional) return null;
    throw new Error(`未找到上传控件：${labels}`);
  }

  const file = new File([imageInfo.bytes], imageInfo.fileName || "upload.png", {
    type: imageInfo.mimeType || "image/png",
  });
  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  input.files = dataTransfer.files;
  input.dispatchEvent(new Event("change", { bubbles: true }));

  await waitForUploadPreview(group, { signal });
  return input;
}

async function waitForUploadPreview(group, options = {}) {
  const { signal, timeout = 30000 } = options;
  return waitUntil(() => {
    const image = [...group.querySelectorAll(".el-upload img, img")]
      .find((node) => isVisibleElement(node) && node.getAttribute("src"));
    return image || null;
  }, {
    signal,
    timeout,
    errorMessage: "等待图片上传预览渲染超时。",
    returnValue: true,
  });
}

function findFieldGroup(scope, labels) {
  const normalizedLabels = (Array.isArray(labels) ? labels : [labels]).map(normalizeText);
  return [...scope.querySelectorAll(".am-input-group")]
    .find((group) => {
      const label = group.querySelector(".am-input-group-label, label, .xing")?.parentElement || group;
      const text = normalizeText(label.textContent);
      return normalizedLabels.some((candidate) => text.includes(candidate));
    }) || null;
}

function nextElementMatching(element, selector) {
  let cursor = element.nextElementSibling;
  while (cursor) {
    if (cursor.matches?.(selector)) return cursor;
    cursor = cursor.nextElementSibling;
  }
  return null;
}

function findVisibleSelectOption(values) {
  const normalizedValues = (Array.isArray(values) ? values : [values]).map(normalizeText).filter(Boolean);
  return [...document.querySelectorAll(".el-select-dropdown__item")]
    .filter(isVisibleElement)
    .find((option) => {
      const text = normalizeText(option.textContent);
      return normalizedValues.some((value) => text.includes(value) || value.includes(text));
    }) || null;
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

function setNativeValue(element, value) {
  const prototype = Object.getPrototypeOf(element);
  const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");
  if (descriptor?.set) {
    descriptor.set.call(element, value);
  } else {
    element.value = value;
  }
}

function isVisibleElement(element) {
  if (!element) return false;
  const style = window.getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") return false;
  const rect = element.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "" || String(value).trim() === "(空)";
}

async function findQuickTeacherButton(teacherName, options = {}) {
  const { signal } = options;
  const teacherList = await waitForElement(".form-container .teacherICon", {
    signal,
    timeout: 15000,
  });
  const normalizedTeacherName = normalizeText(teacherName);

  if (!normalizedTeacherName) {
    return {
      exists: false,
      teacherList,
      button: null,
      availableNames: getQuickTeacherNames(teacherList),
    };
  }

  await waitUntil(
    () => getQuickTeacherNames(teacherList).length > 0,
    {
      signal,
      timeout: 8000,
      errorMessage: "等待快捷指导老师列表渲染超时。",
    },
  ).catch(() => {});

  const buttons = [...teacherList.querySelectorAll("button, .teacherI")];
  const button = buttons.find((candidate) => {
    const candidateName = normalizeText(candidate.textContent);
    return candidateName.includes(normalizedTeacherName) || normalizedTeacherName.includes(candidateName);
  }) || null;

  return {
    exists: Boolean(button),
    teacherList,
    button,
    availableNames: getQuickTeacherNames(teacherList),
  };
}

function getQuickTeacherNames(container) {
  return [...container.querySelectorAll("button, .teacherI")]
    .map((node) => normalizeText(node.textContent))
    .filter(Boolean);
}

function buildTeacherDetectionStatus(teacher, quickTeacherExists, teacherFormFilled) {
  if (quickTeacherExists) return "页面已有该指导老师，已选择并跳过老师表单";
  if (teacherFormFilled) return "页面未找到该老师，已填写指导老师信息";
  if (teacher) return `页面未找到该老师，已匹配到${teacher.source === "internal" ? "内部" : "外部"}老师信息`;
  return "页面未找到该老师，Excel 老师表也未匹配到信息";
}

function buildTeacherDetectionMessage(student, teacher, quickTeacher, teacherFormFilled) {
  if (quickTeacher.exists) {
    return `第 ${student.excelRowNumber} 行 ${student.name}：页面快捷老师区已找到 ${student.teacherName}，已点击选择。下一步将填写参赛信息。`;
  }

  if (teacherFormFilled) {
    return `第 ${student.excelRowNumber} 行 ${student.name}：页面快捷老师区没有 ${student.teacherName}，已填写指导老师信息。下一步将填写参赛信息。`;
  }

  if (teacher) {
    const sourceLabel = teacher.source === "internal" ? "内部" : "外部";
    return `第 ${student.excelRowNumber} 行 ${student.name}：页面快捷老师区没有 ${student.teacherName}，已从${sourceLabel}老师表匹配到信息。下一步将填写指导老师表单。`;
  }

  const availableNames = quickTeacher.availableNames.length ? quickTeacher.availableNames.join("、") : "无";
  return `第 ${student.excelRowNumber} 行 ${student.name}：页面快捷老师区没有 ${student.teacherName || "空老师名"}，老师表也未匹配。当前快捷老师：${availableNames}。`;
}

async function loadRegistrationData() {
  if (!window.AILDExcelData) {
    throw new Error("Excel 解析模块未加载。");
  }

  if (!runtimeState.registrationData) {
    await updateUiState({ message: "正在读取三份 Excel，请稍候。" });
    runtimeState.registrationData = await window.AILDExcelData.loadRegistrationData();
    await updateUiState({
      message: `Excel 读取完成：学生 ${runtimeState.registrationData.students.length} 条，内部老师 ${runtimeState.registrationData.internalTeachers.length} 条，外部老师 ${runtimeState.registrationData.externalTeachers.length} 条。`,
    });
  }

  return runtimeState.registrationData;
}

async function updateUiState(patch) {
  await chrome.storage.local.set(patch);
}

async function waitForElement(selector, options = {}) {
  const { text, timeout = 15000, signal } = options;
  const start = Date.now();

  while (Date.now() - start < timeout) {
    if (signal?.aborted) throw new Error("执行已暂停");

    const element = [...document.querySelectorAll(selector)].find((node) => {
      if (!text) return true;
      return normalizeText(node.textContent).includes(normalizeText(text));
    });
    if (element) return element;

    await delay(120);
  }

  throw new Error(`等待页面元素超时：${selector}${text ? ` / ${text}` : ""}`);
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

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, "");
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function delayWithAbort(ms, signal) {
  if (signal?.aborted) return Promise.reject(new Error("执行已暂停"));
  return new Promise((resolve, reject) => {
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener("abort", () => {
      clearTimeout(timer);
      reject(new Error("执行已暂停"));
    }, { once: true });
  });
}
