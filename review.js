// =========================================================
// 문서 읽기(PDF·엑셀 → 글자)와 검토 로직
// ---------------------------------------------------------
// 파일은 브라우저 안에서만 읽는다(어디에도 전송하지 않음).
// 결과 항목: { section, doc, title, status: ok|warn|fail|info, detail, basis }
// =========================================================

/* global pdfjsLib, XLSX */

const DOC_KINDS = {
  permit: "작업허가서",
  plan: "작업계획서",
  risk: "위험성평가",
  education: "특별안전교육",
};

// ── 문서 읽기 ──
async function readPdf(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const lines = [];
  for (let p = 1; p <= pdf.numPages; p += 1) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    // 같은 높이(y)의 글자 조각을 한 줄로 모은다
    const rows = new Map();
    content.items.forEach((item) => {
      if (!item.str || !item.str.trim()) return;
      const y = Math.round(item.transform[5] / 3) * 3;
      if (!rows.has(y)) rows.set(y, []);
      rows.get(y).push({ x: item.transform[4], text: item.str });
    });
    [...rows.entries()]
      .sort((a, b) => b[0] - a[0])
      .forEach(([, parts]) => {
        lines.push(parts.sort((a, b) => a.x - b.x).map((part) => part.text).join(" ").replace(/\s+/g, " ").trim());
      });
  }
  return { lines, tables: [] };
}

function readXlsx(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const lines = [];
  const tables = [];
  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: "" });
    const cleaned = rows.map((row) => row.map((cell) => String(cell).replace(/\s+/g, " ").trim()));
    tables.push(cleaned);
    cleaned.forEach((row) => {
      const text = row.filter(Boolean).join(" ");
      if (text) lines.push(text);
    });
  });
  return { lines, tables };
}

async function readDocument(file) {
  const buffer = await file.arrayBuffer();
  const name = file.name.toLowerCase();
  if (name.endsWith(".pdf")) return readPdf(buffer);
  if (/\.(xlsx|xlsm|xls)$/.test(name)) return readXlsx(buffer);
  throw new Error("PDF 또는 엑셀(xlsx) 파일만 읽을 수 있습니다.");
}

// ── 글자에서 정보 뽑기 ──
const DATE_PATTERNS = [
  /(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})\s*일?/g,
  /(?<!\d)(\d{2})\s*[./]\s*(\d{1,2})\s*[./]\s*(\d{1,2})(?!\d)/g,
];

function findDates(text) {
  const found = new Set();
  DATE_PATTERNS.forEach((pattern, index) => {
    for (const m of text.matchAll(pattern)) {
      const year = index === 0 ? Number(m[1]) : 2000 + Number(m[1]);
      const month = Number(m[2]);
      const day = Number(m[3]);
      if (year < 2000 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) continue;
      found.add(`${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`);
    }
  });
  return [...found].sort();
}

function findHours(text) {
  const hours = [...text.matchAll(/(\d{1,2}(?:\.\d)?)\s*(?:시간|hr|H\b)/gi)]
    .map((m) => Number(m[1]))
    .filter((h) => h > 0 && h <= 40);
  if (hours.length) return Math.max(...hours);
  // '09:00~13:00' 같은 시각 범위를 더한다
  let minutes = 0;
  for (const m of text.matchAll(/(\d{1,2}):(\d{2})\s*[~∼-]\s*(\d{1,2}):(\d{2})/g)) {
    const span = Number(m[3]) * 60 + Number(m[4]) - (Number(m[1]) * 60 + Number(m[2]));
    if (span > 0 && span <= 12 * 60) minutes += span;
  }
  return minutes ? Math.round((minutes / 60) * 10) / 10 : null;
}

const NAME_HEADER = /^(성\s*명|이\s*름|작업자\s*명?|근로자\s*명|교육생|참석자|이수자|참여자)$/;
const NAME_LIKE = /^[가-힣]{2,4}$/;

function isName(word) {
  const rules = window.PERMIT_RULES;
  return NAME_LIKE.test(word) && !rules.nameStopwords.includes(word)
    && !(word.length >= 3 && new RegExp(rules.nameStopSuffix).test(word));
}

// '투입 인원 4명 (김민수, 이준호)'처럼 명단이 한 줄 안에 있는 경우
const INLINE_NAME_LIST = /(투입\s*인원|작업\s*인원|작업자\s*명단|참여\s*자|참석\s*자|이수\s*자)/;

function findNames(doc) {
  const names = new Set();
  // 엑셀: '성명' 같은 머리글 열 아래의 값
  doc.tables.forEach((rows) => {
    rows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (!NAME_HEADER.test(cell)) return;
        for (let k = r + 1; k < rows.length; k += 1) {
          const value = (rows[k][c] || "").replace(/\s/g, "");
          if (NAME_HEADER.test(value)) break;
          if (isName(value)) names.add(value);
        }
      });
    });
  });
  // PDF(또는 머리글을 못 찾은 엑셀): '성명'이 나온 줄 아래의 2~4글자 한글 단어
  if (!names.size) {
    let inList = false;
    doc.lines.forEach((line) => {
      if (INLINE_NAME_LIST.test(line)) {
        (line.match(/\(([^)]*)\)/g) || []).forEach((group) =>
          group.slice(1, -1).split(/[\s,·ㆍ]+/).forEach((word) => {
            if (isName(word)) names.add(word);
          })
        );
      }
      if (/성\s*명|이\s*름|참석자|이수자|명\s*단/.test(line)) {
        inList = true;
        return;
      }
      if (!inList) return;
      if (/합\s*계|총\s*원|이상\s*\d|위와\s*같이|확인\s*합니다/.test(line)) {
        inList = false;
        return;
      }
      line.split(/[\s,()]+/).forEach((word) => {
        if (isName(word)) names.add(word);
      });
    });
  }
  return [...names];
}

function matchKeywords(lines, keywords) {
  const patterns = keywords.map((k) => new RegExp(k, "i"));
  for (const line of lines) {
    if (patterns.some((p) => p.test(line))) return line;
  }
  return null;
}

function detectWorkTypes(lines) {
  return window.PERMIT_RULES.workTypes.filter((type) => matchKeywords(lines, type.keywords));
}

function analyzeDocument(kind, parts) {
  const lines = parts.flatMap((p) => p.lines);
  const tables = parts.flatMap((p) => p.tables);
  const doc = { kind, lines, tables, text: lines.join("\n") };
  doc.dates = findDates(doc.text);
  doc.names = findNames(doc);
  doc.hours = kind === "education" ? findHours(doc.text) : null;
  doc.workTypes = detectWorkTypes(lines);
  return doc;
}

// ── 검토 ──
function snippet(line) {
  if (!line) return "";
  return line.length > 80 ? `${line.slice(0, 80)}…` : line;
}

function checkItems(results, section, doc, items, basis, missingStatus = "fail") {
  items.forEach((item) => {
    const line = matchKeywords(doc.lines, item.keywords);
    results.push({
      section,
      doc: DOC_KINDS[doc.kind],
      title: item.name,
      status: line ? "ok" : missingStatus,
      detail: line ? `“${snippet(line)}”` : "해당 내용을 찾지 못했습니다.",
      basis,
    });
  });
}

function formatList(values, limit = 8) {
  if (!values.length) return "없음";
  return values.length > limit ? `${values.slice(0, limit).join(", ")} 외 ${values.length - limit}` : values.join(", ");
}

function reviewDocuments(docs) {
  const rules = window.PERMIT_RULES;
  const results = [];
  const { permit, plan, risk, education } = docs;
  const workSource = permit || plan;
  const workTypes = workSource ? workSource.workTypes : [];

  // 1) 작업 정보
  const workDates = permit ? permit.dates : plan ? plan.dates : [];
  const workStart = workDates[0] || null;
  const workers = [...new Set([...(permit ? permit.names : []), ...(plan ? plan.names : [])])];

  // 2) 서류 제출 여부
  Object.entries(DOC_KINDS).forEach(([kind, label]) => {
    results.push({
      section: "제출 서류",
      doc: label,
      title: `${label} 제출`,
      status: docs[kind] ? "ok" : kind === "education" && !workTypes.some((t) => t.education !== "none") ? "info" : "fail",
      detail: docs[kind] ? "파일을 읽었습니다." : "파일이 없습니다.",
      basis: "",
    });
  });

  results.push({
    section: "작업 정보",
    doc: workSource ? DOC_KINDS[workSource.kind] : "-",
    title: "작업 종류",
    status: workTypes.length ? "info" : "warn",
    detail: workTypes.length ? workTypes.map((t) => t.name).join(", ") : "작업 종류를 알아내지 못했습니다. 작업허가서에 작업 내용이 있는지 확인하세요.",
    basis: "",
  });

  // 3) 작업계획서
  if (plan) {
    checkItems(results, "작업계획서", plan, rules.planCommon, rules.planBasis);
    workTypes.forEach((type) => {
      checkItems(results, "작업계획서", plan,
        type.planItems.map((item) => ({ ...item, name: `[${type.name}] ${item.name}` })),
        type.basis, "warn");
    });
  }

  // 4) 위험성평가
  if (risk) {
    checkItems(results, "위험성평가", risk, rules.riskItems, rules.riskBasis);
    results.push({
      section: "위험성평가",
      doc: DOC_KINDS.risk,
      title: "평가 일자",
      status: risk.dates.length ? "ok" : "fail",
      detail: risk.dates.length ? `문서의 날짜: ${formatList(risk.dates, 4)}` : "날짜를 찾지 못했습니다.",
      basis: rules.riskBasis,
    });
    workTypes.forEach((type) => {
      const line = matchKeywords(risk.lines, type.keywords);
      results.push({
        section: "위험성평가",
        doc: DOC_KINDS.risk,
        title: `허가 작업 평가 여부: ${type.name}`,
        status: line ? "ok" : "fail",
        detail: line ? `“${snippet(line)}”` : "작업허가서의 이 작업이 위험성평가에 없습니다.",
        basis: rules.riskBasis,
      });
    });
  }

  // 5) 특별안전교육
  const educationTypes = workTypes.filter((t) => t.education !== "none");
  if (education) {
    checkItems(results, "특별안전교육", education, rules.educationItems, rules.educationBasis);
    const hours = education.hours;
    let status = "fail";
    let detail = "교육시간을 찾지 못했습니다.";
    if (hours !== null) {
      if (hours >= rules.educationHours.full) {
        status = "ok";
        detail = `${hours}시간 (16시간 이상)`;
      } else if (hours >= rules.educationHours.short) {
        status = "warn";
        detail = `${hours}시간 — 일용근로자·1주일 이하 기간제·단기간 작업이면 충족, 그 외 근로자는 16시간 필요`;
      } else {
        detail = `${hours}시간 — 최소 2시간 미만`;
      }
    }
    results.push({ section: "특별안전교육", doc: DOC_KINDS.education, title: "교육시간", status, detail: `${detail}. 기준: ${rules.educationHours.note}`, basis: rules.educationBasis });
    results.push({
      section: "특별안전교육",
      doc: DOC_KINDS.education,
      title: "교육일",
      status: education.dates.length ? "ok" : "fail",
      detail: education.dates.length ? `문서의 날짜: ${formatList(education.dates, 4)}` : "교육일을 찾지 못했습니다.",
      basis: rules.educationBasis,
    });
    educationTypes.forEach((type) => {
      const line = matchKeywords(education.lines, type.keywords);
      results.push({
        section: "특별안전교육",
        doc: DOC_KINDS.education,
        title: `교육 대상 작업 일치: ${type.name}`,
        status: line ? "ok" : type.education === "required" ? "fail" : "warn",
        detail: line ? `“${snippet(line)}”` : `교육 내용에 이 작업이 없습니다. ${type.educationNote}`,
        basis: rules.educationBasis,
      });
    });
  } else if (educationTypes.length) {
    results.push({
      section: "특별안전교육",
      doc: DOC_KINDS.education,
      title: "특별교육 필요 여부",
      status: educationTypes.some((t) => t.education === "required") ? "fail" : "warn",
      detail: educationTypes.map((t) => t.educationNote).join(" "),
      basis: rules.educationBasis,
    });
  }

  // 6) 서류 간 일치
  const cross = "서류 간 일치";
  if (workStart) {
    if (plan && plan.dates.length && workSource !== plan) {
      const inRange = workDates.every((d) => d >= plan.dates[0] && d <= plan.dates[plan.dates.length - 1]);
      results.push({
        section: cross, doc: "작업허가서 ↔ 작업계획서", title: "작업일이 작업계획 기간 안에 있음",
        status: inRange ? "ok" : "warn",
        detail: `작업일 ${formatList(workDates, 3)} / 계획서 날짜 ${plan.dates[0]} ~ ${plan.dates[plan.dates.length - 1]}`,
        basis: "",
      });
    }
    if (risk && risk.dates.length) {
      const before = risk.dates[0] <= workStart;
      results.push({
        section: cross, doc: "작업허가서 ↔ 위험성평가", title: "위험성평가를 작업 전에 실시",
        status: before ? "ok" : "fail",
        detail: `평가일(가장 이른 날짜) ${risk.dates[0]} / 작업 시작 ${workStart}`,
        basis: rules.riskBasis,
      });
    }
    if (education && education.dates.length) {
      const before = education.dates[0] <= workStart;
      results.push({
        section: cross, doc: "작업허가서 ↔ 특별안전교육", title: "특별교육을 작업 전에 실시",
        status: before ? "ok" : "fail",
        detail: `교육일(가장 이른 날짜) ${education.dates[0]} / 작업 시작 ${workStart} — 최초 작업 전 4시간 이상은 작업 전에 해야 합니다.`,
        basis: rules.educationBasis,
      });
    }
  } else if (workSource) {
    results.push({ section: cross, doc: DOC_KINDS[workSource.kind], title: "작업일", status: "warn", detail: "작업일을 찾지 못해 날짜 전후 관계를 확인하지 못했습니다.", basis: "" });
  }

  if (plan && workSource !== plan && workTypes.length) {
    const planTypes = plan.workTypes.map((t) => t.id);
    const missing = workTypes.filter((t) => !planTypes.includes(t.id));
    results.push({
      section: cross, doc: "작업허가서 ↔ 작업계획서", title: "작업 종류 일치",
      status: missing.length ? "warn" : "ok",
      detail: missing.length ? `계획서에 없는 작업: ${missing.map((t) => t.name).join(", ")}` : "허가 작업이 계획서에 모두 있습니다.",
      basis: "",
    });
  }

  // 계획서·위험성평가에는 있는데 허가서에 없는 작업(허가 범위 밖 작업일 수 있음)
  if (permit) {
    const permitIds = permit.workTypes.map((t) => t.id);
    const extra = [plan, risk].filter(Boolean)
      .flatMap((doc) => doc.workTypes.filter((t) => !permitIds.includes(t.id)).map((t) => `${t.name}(${DOC_KINDS[doc.kind]})`));
    if (extra.length) {
      results.push({
        section: cross, doc: "작업허가서 ↔ 계획서·위험성평가", title: "허가서에 없는 작업이 다른 서류에 있음",
        status: "warn",
        detail: `${[...new Set(extra)].join(", ")} — 위험요인 언급일 수도 있으니, 실제로 하는 작업이면 허가 구분에 넣었는지 확인하세요.`,
        basis: "",
      });
    }
  }

  if (education && educationTypes.length) {
    if (workers.length && education.names.length) {
      const missing = workers.filter((name) => !education.names.includes(name));
      results.push({
        section: cross, doc: "작업자 명단 ↔ 교육 이수자", title: "투입 작업자 전원 특별교육 이수",
        status: missing.length ? "fail" : "ok",
        detail: missing.length
          ? `교육 명단에 없는 작업자 ${missing.length}명: ${formatList(missing)}`
          : `작업자 ${workers.length}명이 모두 교육 명단에 있습니다.`,
        basis: rules.educationBasis,
      });
    } else {
      results.push({
        section: cross, doc: "작업자 명단 ↔ 교육 이수자", title: "투입 작업자 전원 특별교육 이수",
        status: "warn",
        detail: `이름을 읽지 못해 대조하지 못했습니다(작업자 ${workers.length}명, 교육 명단 ${education.names.length}명). 직접 확인하세요.`,
        basis: rules.educationBasis,
      });
    }
  }
  if (risk && workers.length && risk.names.length) {
    const missing = workers.filter((name) => !risk.names.includes(name));
    results.push({
      section: cross, doc: "작업자 명단 ↔ 위험성평가 참여자", title: "작업자의 위험성평가 공유·참여",
      status: missing.length ? "warn" : "ok",
      detail: missing.length ? `참여 명단에 없는 작업자: ${formatList(missing)}` : "작업자 모두 참여 명단에 있습니다.",
      basis: rules.riskBasis,
    });
  }

  return {
    results,
    summary: {
      workTypes: workTypes.map((t) => t.name),
      workDates,
      workers,
    },
  };
}

window.PermitReview = { DOC_KINDS, readDocument, analyzeDocument, reviewDocuments };
