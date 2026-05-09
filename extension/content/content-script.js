const runtimeState = {
  status: "idle",
  controller: null,
  registrationData: null,
  currentExcelRow: null,
  completedCount: 0,
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
  runtimeState.completedCount = isResume ? runtimeState.completedCount : 0;

  await updateUiState({
    status: "running",
    startRow: payload.startRow,
    pauseAfter: payload.pauseAfter,
    completedCount: runtimeState.completedCount,
    currentExcelRow: runtimeState.currentExcelRow,
    message: isResume ? "继续执行，正在准备读取报名数据。" : "开始执行，正在准备读取报名数据。",
  });

  await waitForElement(".form-container", { text: "参赛信息", signal: runtimeState.controller.signal });
  const data = await loadRegistrationData();
  let completedThisRun = 0;
  if (isResume) {
    await acceptCommitmentForNextStudent({
      signal: runtimeState.controller.signal,
      timeout: 2000,
    }).catch(() => {});
  }

  while (runtimeState.status === "running") {
    const student = window.AILDExcelData.getStudentByExcelRow(data, runtimeState.currentExcelRow);

    if (!student) {
      await updateUiState({
        status: "done",
        currentStudent: null,
        completedCount: runtimeState.completedCount,
        message: `Excel 第 ${runtimeState.currentExcelRow} 行没有可报名学生，执行结束。`,
      });
      runtimeState.status = "done";
      return;
    }

    await processStudentRegistration(student, data);
    runtimeState.completedCount += 1;
    completedThisRun += 1;
    runtimeState.currentExcelRow += 1;

    await updateUiState({
      completedCount: runtimeState.completedCount,
      currentExcelRow: runtimeState.currentExcelRow,
      currentStudent: {
        name: student.name,
        teacherName: student.teacherName || "-",
        status: "报名完成",
        excelRowNumber: student.excelRowNumber,
      },
    });

    if (completedThisRun >= payload.pauseAfter) {
      runtimeState.status = "paused";
      await updateUiState({
        status: "paused",
        completedCount: runtimeState.completedCount,
        currentExcelRow: runtimeState.currentExcelRow,
        currentStudent: {
          name: student.name,
          teacherName: student.teacherName || "-",
          status: `本次已完成 ${completedThisRun} 个学生，按设置自动暂停`,
          excelRowNumber: student.excelRowNumber,
        },
        message: `本次已完成 ${completedThisRun} 个学生，累计完成 ${runtimeState.completedCount} 个。点击“继续执行”将从 Excel 第 ${runtimeState.currentExcelRow} 行继续。`,
      });
      return;
    }

    await updateUiState({
      status: "running",
      message: `已完成 ${student.name}，准备第 ${runtimeState.currentExcelRow} 行。`,
    });
  }
}

async function processStudentRegistration(student, data) {
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
    status: "running",
    currentStudent: {
      name: student.name,
      teacherName: student.teacherName || "-",
      status: "正在添加队员",
      excelRowNumber: student.excelRowNumber,
      teamName: competitionInfo.teamName,
    },
    message: `正在添加队员：${student.name}。`,
  });
  await addStudentMember(student, { signal: runtimeState.controller.signal });
  await updateUiState({
    status: "running",
    currentStudent: {
      name: student.name,
      teacherName: student.teacherName || "-",
      status: "正在预览确认并提交报名",
      excelRowNumber: student.excelRowNumber,
      teamName: competitionInfo.teamName,
    },
    message: `正在提交 ${student.name} 的报名信息。`,
  });
  await submitCurrentRegistration({ signal: runtimeState.controller.signal });

  await updateUiState({
    status: "running",
    currentStudent: {
      name: student.name,
      teacherName: student.teacherName || "-",
      status: `${buildTeacherDetectionStatus(teacher, quickTeacher.exists, teacherFormFilled)}；报名已提交`,
      excelRowNumber: student.excelRowNumber,
      teamName: competitionInfo.teamName,
    },
      message: `${student.name} 报名成功并已返回主页面。团队名称：${competitionInfo.teamName}。`,
  });
}

async function acceptCommitmentForNextStudent(options = {}) {
  if (!window.AILDRegisterActions?.acceptCommitmentDialog) {
    throw new Error("承诺书处理模块未加载。");
  }
  await window.AILDRegisterActions.acceptCommitmentDialog(options);
}

async function submitCurrentRegistration(options = {}) {
  if (!window.AILDRegisterActions?.submitRegistrationPreview) {
    throw new Error("预览提交模块未加载。");
  }
  await window.AILDRegisterActions.submitRegistrationPreview(options);
}

async function addStudentMember(student, options = {}) {
  const { signal } = options;
  const addButton = await waitUntil(() => findButtonByText("添加队员"), {
    signal,
    timeout: 15000,
    errorMessage: "未找到添加队员按钮。",
    returnValue: true,
  });
  addButton.click();

  const modal = await waitUntil(() => findMemberModal(), {
    signal,
    timeout: 15000,
    errorMessage: "等待添加队员弹窗超时。",
    returnValue: true,
  });
  const form = modal.querySelector("form");
  if (!form) throw new Error("添加队员弹窗中未找到表单。");

  await fillInputByLabel(form, ["姓名中文", "中文姓名", "姓名"], student.name, { signal });
  await selectByLabel(form, ["性别"], student.gender, { signal });
  await fillInputByLabel(form, ["民族"], student.nation, { signal });
  await fillInputByLabel(form, ["学校全称", "学校"], student.school, { signal });
  await selectByLabel(form, ["年级"], gradeCandidates(student), { signal });
  await fillStudentIdNumber(form, student.idNumber, { signal });
  await fillInputByLabel(form, ["监护人邮箱", "邮箱"], student.guardianEmail, { signal });
  await fillInputByLabel(form, ["监护人手机", "手机"], student.guardianPhone, { signal });
  await uploadByLabel(form, ["一寸照片", "照片"], student.studentPhoto, { signal });

  const submitButton = modal.querySelector(".memOk") || [...modal.querySelectorAll("button")].find((button) => normalizeText(button.textContent).includes("提交"));
  if (!submitButton) throw new Error("添加队员弹窗中未找到提交按钮。");
  submitButton.click();

  await waitUntil(() => !findMemberModal(), {
    signal,
    timeout: 20000,
    errorMessage: "等待添加队员弹窗关闭超时。",
  });
}

async function fillStudentIdNumber(form, idNumber, options = {}) {
  const { signal } = options;
  if (isBlank(idNumber)) throw new Error("学生身份证号为空。");

  const groups = [...form.querySelectorAll(".am-input-group")].filter((group) => {
    const text = normalizeText(group.textContent);
    return text.includes("证件号码") || text.includes("身份证号");
  });
  const group = groups.find((candidate) => candidate.querySelector('input.am-form-field:not([style*="display: none"])'))
    || groups[0];
  if (!group) throw new Error("未找到学生身份证号输入框。");

  const input = [...group.querySelectorAll('input.am-form-field:not([style*="display: none"])')]
    .find(isVisibleElement) || [...group.querySelectorAll("input.am-form-field")].find(isVisibleElement);
  if (!input) throw new Error("未找到可见学生身份证号输入框。");

  setNativeValue(input, idNumber);
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
  input.dispatchEvent(new Event("blur", { bubbles: true }));
  await delayWithAbort(80, signal);
}

function findMemberModal() {
  return [...document.querySelectorAll(".am-modal.am-modal-active, .am-modal")]
    .filter(isVisibleElement)
    .find((modal) => {
      const text = normalizeText(modal.textContent);
      return text.includes("请输入队员信息") || (text.includes("姓名中文") && text.includes("监护人邮箱") && text.includes("提交"));
    }) || null;
}

function findButtonByText(text) {
  const normalizedText = normalizeText(text);
  return [...document.querySelectorAll("button")]
    .filter(isVisibleElement)
    .find((button) => normalizeText(button.textContent).includes(normalizedText)) || null;
}

function gradeCandidates(student) {
  const gradeText = student.grade;
  const normalizedGrade = normalizeText(gradeText);
  const candidates = [];

  // 添加原始年级文本作为兜底
  if (gradeText) {
    candidates.push(gradeText);
  }

  if (normalizedGrade) {
    // 初高中年级映射表：中文年级关键词 -> 页面数字年级 + 别名
    const gradeMap = [
      { keywords: ['高一', '高中一', '高中1', '10年级'], page: '10年级', aliases: ['高一', '高中一年级'] },
      { keywords: ['高二', '高中二', '高中2', '11年级'], page: '11年级', aliases: ['高二', '高中二年级'] },
      { keywords: ['高三', '高中三', '高中3', '12年级'], page: '12年级', aliases: ['高三', '高中三年级'] },
      { keywords: ['初一', '初中一', '初中1', '7年级'],  page: '7年级',  aliases: ['初一', '初中一年级'] },
      { keywords: ['初二', '初中二', '初中2', '8年级'],  page: '8年级',  aliases: ['初二', '初中二年级'] },
      { keywords: ['初三', '初中三', '初中3', '9年级'],  page: '9年级',  aliases: ['初三', '初中三年级'] },
      { keywords: ['六年级', '小学六', '6年级'],         page: '6年级',  aliases: ['六年级'] },
      { keywords: ['五年级', '小学五', '5年级'],         page: '5年级',  aliases: ['五年级'] },
      { keywords: ['四年级', '小学四', '4年级'],         page: '4年级',  aliases: ['四年级'] },
      { keywords: ['三年级', '小学三', '3年级'],         page: '3年级',  aliases: ['三年级'] },
      { keywords: ['二年级', '小学二', '2年级'],         page: '2年级',  aliases: ['二年级'] },
      { keywords: ['一年级', '小学一', '1年级'],         page: '1年级',  aliases: ['一年级'] },
    ];

    // 按顺序匹配（初高中优先于小学，避免"高一年级"误匹配"一年级"）
    const matched = gradeMap.find(({ keywords }) =>
      keywords.some((kw) => normalizedGrade.includes(kw))
    );

    if (matched) {
      candidates.push(matched.page, ...matched.aliases);
    }
  }

  const result = [...new Set(candidates.filter(Boolean))];
  console.log('年级处理结果', result);
  return result;
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
    remoteSearchText: "智械未来",
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
      .find((node) => isSectionTitleMatch(node, titles));
    const section = title ? nextElementMatching(title, ".form-item") : null;
    return section?.querySelector("form") || null;
  }, {
    timeout,
    signal,
    errorMessage: `等待表单渲染超时：${titles.join(" / ")}`,
    returnValue: true,
  });
}

function isSectionTitleMatch(node, titles) {
  const nodeText = normalizeText(node.textContent).replace(/[：:]/g, "");
  if (!nodeText || nodeText.includes("快速注册")) return false;

  return titles.some((title) => {
    const expected = normalizeText(title).replace(/[：:]/g, "");
    return nodeText === expected || nodeText.startsWith(expected);
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
    throw new Error(`未找到输入字段：${labels}。当前表单字段：${listFieldHints(scope)}`);
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
  const {
    signal,
    optional = false,
    selectIndex = 0,
    filterText = "",
    remoteSearchText = "",
  } = options;
  const values = Array.isArray(value) ? value.filter((item) => !isBlank(item)) : [value].filter((item) => !isBlank(item));
  if (!values.length) {
    if (optional) return null;
    throw new Error(`字段为空，无法选择：${labels}`);
  }

  const group = findFieldGroup(scope, labels);
  if (!group) {
    if (optional) return null;
    throw new Error(`未找到选择字段：${labels}。当前表单字段：${listFieldHints(scope)}`);
  }

  const selectInput = [...group.querySelectorAll(".el-select input.el-input__inner, input[readonly]")]
    .filter(isVisibleElement)[selectIndex];
  if (!selectInput) {
    if (optional) return null;
    throw new Error(`未找到选择框：${labels}`);
  }

  selectInput.click();
  await delayWithAbort(120, signal);

  const searchText = remoteSearchText || filterText;
  if (searchText) {
    setNativeValue(selectInput, "");
    dispatchTextInputEvents(selectInput);
    await delayWithAbort(80, signal);

    setNativeValue(selectInput, searchText);
    dispatchTextInputEvents(selectInput);
    await delayWithAbort(remoteSearchText ? 500 : 120, signal);
  }

  const option = await waitUntil(() => findVisibleSelectOption(values), {
    signal,
    timeout: remoteSearchText ? 20000 : 10000,
    errorMessage: `等待选择项渲染超时：${values.join(" / ")}`,
    returnValue: true,
  });

  option.click();
  selectInput.dispatchEvent(new Event("change", { bubbles: true }));
  selectInput.dispatchEvent(new Event("blur", { bubbles: true }));
  await delayWithAbort(120, signal);
  return option;
}

function dispatchTextInputEvents(input) {
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, key: "Process" }));
  input.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, key: "Process" }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

async function fillUniqueTeamName(form, student, options = {}) {
  const { signal } = options;
  for (let attempt = 0; attempt < 15; attempt += 1) {
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
  const CHARS =
    "智慧思维悟启迪析辩探钻研" +
    "创新锐进取革拓越攀超卓绝" +
    "明星辉耀曙晨曦灿烁炯炳朗" +
    "航远征探索跋涉渡越闯踏行" +
    "云霞峰岳渊澜涌潮浪川岚霖" +
    "龙凤麟翔腾跃奋勇毅韧志恒" +
    "未来望梦憬程途域境界纪元" +
    "学思问行知合致良敏勤笃博" +
    "青春少年壮志凌云奔赴山海";

  // 用学生固定信息生成确定性基础 seed（保证同一学生首次生成稳定）
  const stableSource = `${student.idNumber || ""}|${student.name || ""}|${student.excelRowNumber || ""}`;
  let stableHash = 2166136261; // FNV-1a 32bit offset basis
  for (const char of stableSource) {
    stableHash ^= char.charCodeAt(0);
    stableHash = (stableHash * 16777619) >>> 0;
  }

  // 引入真随机 + attempt，保证每次重试都不同
  const randomBytes = new Uint32Array(2);
  crypto.getRandomValues(randomBytes);
  const entropy = (randomBytes[0] ^ (attempt * 2654435761)) >>> 0;
  const entropy2 = (randomBytes[1] ^ stableHash) >>> 0;

  // 用多个独立 seed 各自取一个字符，避免步长固定导致的碰撞
  const seeds = [
    (stableHash ^ entropy) >>> 0,
    (entropy + 0x9e3779b9) >>> 0,
    (entropy2 ^ (attempt * 0x517cc1b7)) >>> 0,
    (stableHash + entropy2 + attempt) >>> 0,
  ];

  const name = seeds.map((seed) => CHARS[seed % CHARS.length]).join("");
  return name.replace(/[^\u4e00-\u9fa5]/g, "").slice(0, 4);
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

  if (isDirectCity) {
    // 直辖市：页面"市"下拉选项是不带"市"字的名称（如"北京"），优先匹配
    const directCityName = normalizedProvince.replace(/市$/, "");
    console.log('第二次格式化 市', isDirectCity, directCityName);
    return [directCityName];
  }

  // 普通省份：优先不带"市"后缀，再兜底带"市"
  return [...new Set([
    normalizedCity.replace(/市$/, ""),
    normalizedCity,
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
    throw new Error(`未找到上传字段：${labels}。当前表单字段：${listFieldHints(scope)}`);
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
      const fieldText = getFieldHintText(group);
      return normalizedLabels.some((candidate) => fieldText.includes(candidate));
    }) || null;
}

function getFieldHintText(group) {
  const label = group.querySelector(".am-input-group-label, label, .disTtl");
  const placeholders = [...group.querySelectorAll("input, textarea")]
    .map((input) => input.getAttribute("placeholder") || "")
    .join("");
  const visibleText = group.textContent || "";
  return normalizeText(`${label?.textContent || ""}${placeholders}${visibleText}`);
}

function listFieldHints(scope) {
  return [...scope.querySelectorAll(".am-input-group")]
    .map(getFieldHintText)
    .filter(Boolean)
    .slice(0, 12)
    .join(" | ") || "无";
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
