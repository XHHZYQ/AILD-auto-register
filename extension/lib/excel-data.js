(function exposeExcelData(global) {
  const FILES = {
    students: "assets/registration/students.xlsx",
    internalTeachers: "assets/registration/internal-teachers.xlsx",
    externalTeachers: "assets/registration/external-teachers.xlsx",
  };

  const FIELD_ALIASES = {
    serialNumber: ["序号"],
    studentName: ["参赛姓名", "姓名"],
    teacherName: ["指导老师姓名"],
    gender: ["性别"],
    nation: ["民族"],
    school: ["学校全称", "学校"],
    province: ["省"],
    city: ["市"],
    grade: ["年级"],
    idNumber: ["身份证号码", "身份证号", "证件号码"],
    guardianEmail: ["监护人邮箱", "邮箱"],
    guardianPhone: ["监护人手机", "手机"],
    studentPhoto: ["参赛选手1寸照片", "一寸照片", "照片"],
    teacherPhoto: ["指导老师1寸照片", "一寸照片", "照片"],
    teacherCertificate: ["指导老师资格证书", "资格证书"],
    teacherPhone: ["手机"],
    teacherEmail: ["邮箱"],
    teacherWorkplace: ["学校（单位）", "学校", "单位"],
    teacherAccount: ["劳智赛官网注册账号", "注册账号"],
    teacherPassword: ["劳智赛官网注册密码", "注册密码"],
    guidedStudentName: ["您所指导的学员姓名", "学员姓名"],
  };

  const DIRECT_CITIES = new Set(["北京", "北京市", "天津", "天津市", "上海", "上海市", "重庆", "重庆市"]);

  async function loadRegistrationData() {
    const [studentWorkbook, internalWorkbook, externalWorkbook] = await Promise.all([
      readWorkbook(FILES.students),
      readWorkbook(FILES.internalTeachers),
      readWorkbook(FILES.externalTeachers),
    ]);

    const studentImages = extractWorkbookImages(studentWorkbook);
    const internalTeacherImages = extractWorkbookImages(internalWorkbook);
    const externalTeacherImages = extractWorkbookImages(externalWorkbook);

    const students = parseRows(studentWorkbook, {
      preferredAliases: [
        FIELD_ALIASES.studentName,
        FIELD_ALIASES.province,
        FIELD_ALIASES.city,
        FIELD_ALIASES.teacherName,
      ],
    }).map((row) => normalizeStudent(row, studentImages)).filter((row) => row.name);
    const internalTeachers = parseRows(internalWorkbook, {
      preferredAliases: [
        FIELD_ALIASES.teacherName,
        FIELD_ALIASES.idNumber,
        FIELD_ALIASES.teacherPhone,
        FIELD_ALIASES.teacherPhoto,
        FIELD_ALIASES.teacherAccount,
      ],
    })
      .map((row) => normalizeTeacher(row, "internal", internalTeacherImages))
      .filter((row) => row.name);
    const externalTeachers = parseRows(externalWorkbook, {
      preferredAliases: [
        FIELD_ALIASES.teacherName,
        FIELD_ALIASES.teacherWorkplace,
        FIELD_ALIASES.guidedStudentName,
        FIELD_ALIASES.teacherPhoto,
        FIELD_ALIASES.teacherCertificate,
      ],
    })
      .map((row) => normalizeTeacher(row, "external", externalTeacherImages))
      .filter((row) => row.name);

    return {
      students,
      internalTeachers,
      externalTeachers,
      teachersByName: buildTeachersByName(internalTeachers, externalTeachers),
    };
  }

  function getStudentByExcelRow(data, excelRowNumber) {
    const rowNumber = Number.parseInt(excelRowNumber, 10);
    if (!Number.isFinite(rowNumber) || rowNumber < 2) {
      throw new Error("Excel 起始行必须大于等于 2。");
    }
    return data.students[rowNumber - 2] || null;
  }

  function findTeacher(data, teacherName) {
    const key = normalizeName(teacherName);
    if (!key) return null;
    return data.teachersByName.get(key) || findTeacherByContains(data.internalTeachers, key) || findTeacherByContains(data.externalTeachers, key);
  }

  function parseRows(workbook, options = {}) {
    const sheetName = chooseSheetName(workbook, options.preferredAliases || []);
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    });

    if (!matrix.length) return [];

    const headers = matrix[0].map((value) => String(value || "").trim());
    return matrix.slice(1).map((row, index) => ({
      excelRowNumber: index + 2,
      sheetName,
      raw: row,
      get: (aliases) => getCellByAliases(headers, row, aliases),
    }));
  }

  function chooseSheetName(workbook, preferredAliases) {
    const adjustedSheetName = workbook.SheetNames.find((sheetName) => (
      normalizeTextForSheetName(sheetName) === "调整后数据"
    ));
    if (adjustedSheetName) return adjustedSheetName;

    throw new Error(`Excel 缺少“调整后数据”sheet，当前 sheets：${workbook.SheetNames.join("、")}`);
  }

  function getCellByAliases(headers, row, aliases) {
    const keys = Array.isArray(aliases) ? aliases : [aliases];
    const normalizedKeys = keys.map(normalizeHeader).filter(Boolean);
    const index = headers.findIndex((header) => {
      const normalizedHeader = normalizeHeader(header);
      return normalizedKeys.some((key) => normalizedHeader.includes(key) || key.includes(normalizedHeader));
    });
    return index >= 0 ? cleanCell(row[index]) : "";
  }

  function normalizeStudent(row, images) {
    const grade = row.get(FIELD_ALIASES.grade);
    const province = row.get(FIELD_ALIASES.province);
    const city = row.get(FIELD_ALIASES.city);

    return {
      excelRowNumber: row.excelRowNumber,
      sheetName: row.sheetName,
      serialNumber: row.get(FIELD_ALIASES.serialNumber),
      name: row.get(FIELD_ALIASES.studentName),
      teacherName: row.get(FIELD_ALIASES.teacherName),
      gender: row.get(FIELD_ALIASES.gender),
      nation: row.get(FIELD_ALIASES.nation),
      school: row.get(FIELD_ALIASES.school),
      province,
      city: normalizeCityForProvince(province, city),
      grade,
      gradeNumber: getGradeNumber(grade),
      groupName: getGroupNameByGrade(grade),
      idNumber: row.get(FIELD_ALIASES.idNumber),
      guardianEmail: row.get(FIELD_ALIASES.guardianEmail),
      guardianPhone: row.get(FIELD_ALIASES.guardianPhone),
      studentPhoto: normalizeImageCell(row.get(FIELD_ALIASES.studentPhoto), images),
    };
  }

  function normalizeTeacher(row, source, images) {
    return {
      source,
      excelRowNumber: row.excelRowNumber,
      sheetName: row.sheetName,
      serialNumber: row.get(FIELD_ALIASES.serialNumber),
      name: row.get(FIELD_ALIASES.teacherName),
      gender: row.get(FIELD_ALIASES.gender),
      nation: row.get(FIELD_ALIASES.nation),
      idNumber: row.get(FIELD_ALIASES.idNumber),
      phone: row.get(FIELD_ALIASES.teacherPhone),
      email: row.get(FIELD_ALIASES.teacherEmail),
      workplace: row.get(FIELD_ALIASES.teacherWorkplace),
      account: row.get(FIELD_ALIASES.teacherAccount),
      password: row.get(FIELD_ALIASES.teacherPassword),
      guidedStudentName: row.get(FIELD_ALIASES.guidedStudentName),
      photo: normalizeImageCell(row.get(FIELD_ALIASES.teacherPhoto), images),
      certificate: normalizeImageCell(row.get(FIELD_ALIASES.teacherCertificate), images),
    };
  }

  function buildTeachersByName(internalTeachers, externalTeachers) {
    const map = new Map();
    for (const teacher of [...internalTeachers, ...externalTeachers]) {
      const key = normalizeName(teacher.name);
      if (key && !map.has(key)) map.set(key, teacher);
    }
    return map;
  }

  function findTeacherByContains(teachers, normalizedName) {
    return teachers.find((teacher) => {
      const candidate = normalizeName(teacher.name);
      return candidate.includes(normalizedName) || normalizedName.includes(candidate);
    }) || null;
  }

  async function readWorkbook(path) {
    const url = chrome.runtime.getURL(path);
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`读取 Excel 失败：${path}`);
    }
    const buffer = await response.arrayBuffer();
    return XLSX.read(buffer, {
      type: "array",
      bookFiles: true,
      cellDates: true,
      cellFormula: true,
      cellNF: false,
      cellText: true,
    });
  }

  function extractWorkbookImages(workbook) {
    const files = workbook.files || {};
    const cellImagesXml = getWorkbookFileText(files, "xl/cellimages.xml");
    const relationshipsXml = getWorkbookFileText(files, "xl/_rels/cellimages.xml.rels");
    if (!cellImagesXml || !relationshipsXml) return new Map();

    const relationshipTargets = new Map();
    for (const match of relationshipsXml.matchAll(/<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g)) {
      relationshipTargets.set(match[1], match[2]);
    }

    const imageMap = new Map();
    for (const match of cellImagesXml.matchAll(/<xdr:cNvPr\b[^>]*\bname="([^"]+)"[\s\S]*?<a:blip\b[^>]*\br:embed="([^"]+)"/g)) {
      const [, imageId, relationshipId] = match;
      const target = relationshipTargets.get(relationshipId);
      if (!target) continue;

      const normalizedTarget = target.startsWith("../") ? target.replace(/^\.\.\//, "xl/") : `xl/${target}`;
      const file = files[normalizedTarget];
      const bytes = getWorkbookFileBytes(file);
      if (!bytes) continue;

      const fileName = normalizedTarget.split("/").pop() || `${imageId}.png`;
      imageMap.set(imageId, {
        imageId,
        fileName,
        mimeType: getMimeType(fileName),
        bytes,
      });
    }

    return imageMap;
  }

  function getWorkbookFileText(files, path) {
    const file = files[path];
    if (!file) return "";
    if (typeof file.asNodeBuffer === "function") return file.asNodeBuffer().toString("utf8");
    if (file.content instanceof Uint8Array) return new TextDecoder().decode(file.content);
    if (Array.isArray(file.content)) return new TextDecoder().decode(new Uint8Array(file.content));
    return String(file.content || "");
  }

  function getWorkbookFileBytes(file) {
    if (!file) return null;
    if (typeof file.asNodeBuffer === "function") return new Uint8Array(file.asNodeBuffer());
    if (file.content instanceof Uint8Array) return new Uint8Array(file.content);
    if (Array.isArray(file.content)) return new Uint8Array(file.content);
    return null;
  }

  function getMimeType(fileName) {
    const extension = String(fileName).split(".").pop().toLowerCase();
    const mimeTypes = {
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      png: "image/png",
      gif: "image/gif",
      webp: "image/webp",
      bmp: "image/bmp",
    };
    return mimeTypes[extension] || "application/octet-stream";
  }

function normalizeCityForProvince(province, city) {
  if (!province) return city;
  // 直辖市时，city 字段统一返回不带"市"后缀的省名（与页面选项保持一致）
  console.log('第一次格式化 市', DIRECT_CITIES.has(province) ? province.replace(/市$/, "") : city);
  return DIRECT_CITIES.has(province) ? province.replace(/市$/, "") : city;
}
function getGradeNumber(grade) {
  const text = String(grade || "").trim().replace(/\s+/g, "");

  // 1. 优先匹配初高中年级（必须在纯数字和中文数字之前判断）
  const middleHighMap = [
    { keywords: ["高一", "高中一", "高中1"], number: 10 },
    { keywords: ["高二", "高中二", "高中2"], number: 11 },
    { keywords: ["高三", "高中三", "高中3"], number: 12 },
    { keywords: ["初一", "初中一", "初中1"], number: 7 },
    { keywords: ["初二", "初中二", "初中2"], number: 8 },
    { keywords: ["初三", "初中三", "初中3"], number: 9 },
  ];
  const matched = middleHighMap.find(({ keywords }) =>
    keywords.some((kw) => text.includes(kw))
  );
  if (matched) return matched.number;

  // 2. 纯阿拉伯数字
  const digit = text.match(/\d+/);
  if (digit) return Number.parseInt(digit[0], 10);

  // 3. 中文数字（小学场景，此时已排除初高中干扰）
  const chineseDigits = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
  if (text.includes("十一")) return 11;
  if (text.includes("十二")) return 12;
  const match = text.match(/[一二三四五六七八九十]/);
  return match ? chineseDigits[match[0]] : null;
}

  function getGroupNameByGrade(grade) {
    const number = getGradeNumber(grade);
    if (number >= 1 && number <= 3) return "小学低年级组（1-3）";
    if (number >= 4 && number <= 6) return "小学高年级组（4-6）";
    if (number >= 7 && number <= 9) return "初中组（7-9）";
    if (number >= 10 && number <= 12) return "高中组（10-12）";
    return "";
  }

  function normalizeImageCell(value, images) {
    const text = cleanCell(value);
    const imageId = text.match(/DISPIMG\("([^"]+)"/i)?.[1] || "";
    const image = imageId ? images.get(imageId) : null;
    return {
      value: text,
      imageId,
      hasImageReference: Boolean(imageId),
      hasImageFile: Boolean(image),
      fileName: image?.fileName || "",
      mimeType: image?.mimeType || "",
      bytes: image?.bytes || null,
    };
  }

  function cleanCell(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim();
  }

  function normalizeHeader(value) {
    return String(value || "")
      .replace(/[：:]/g, "")
      .replace(/^\s*\d+[、.．]\s*/g, "")
      .replace(/[（）()]/g, "")
      .replace(/\s+/g, "")
      .trim();
  }

  function normalizeTextForSheetName(value) {
    return String(value || "").replace(/\s+/g, "").trim();
  }

  function normalizeName(value) {
    return String(value || "").replace(/\s+/g, "").trim();
  }

  global.AILDExcelData = {
    loadRegistrationData,
    getStudentByExcelRow,
    findTeacher,
    getGradeNumber,
    getGroupNameByGrade,
    normalizeCityForProvince,
  };
})(window);
