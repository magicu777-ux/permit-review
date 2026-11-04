// =========================================================
// 문서 읽기(PDF·엑셀·사진 → 글자)와 검토 로직
// ---------------------------------------------------------
// 파일은 브라우저 안에서만 읽는다(어디에도 전송하지 않음).
// 결과 항목: { section, doc, title, status: fail|warn|manual|ok|info, detail, basis }
// =========================================================

/* global pdfjsLib, XLSX */

const DOC_KINDS = {
  permit: "작업허가서",
  plan: "작업계획서",
  risk: "위험성평가",
  education: "특별안전교육",
  attachments: "기타 첨부",
};

// ── 문서 읽기 ──
const IMAGE_OPS = () => [pdfjsLib.OPS.paintImageXObject, pdfjsLib.OPS.paintInlineImageXObject,
  pdfjsLib.OPS.paintImageXObjectRepeat, pdfjsLib.OPS.paintJpegXObject].filter((op) => op !== undefined);

async function readPdf(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pages = [];
  let images = 0;
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
    const lines = [...rows.entries()]
      .sort((a, b) => b[0] - a[0])
      .map(([, parts]) => parts.sort((a, b) => a.x - b.x).map((part) => part.text).join(" ").replace(/\s+/g, " ").trim());
    // 표의 좁은 칸에서 줄바꿈된 글자('투입 인' + '원')를 한 칸으로 다시 이어 붙인 줄도 더한다(찾기용)
    // 가운데 정렬 칸은 둘째 줄의 x가 달라지므로, 바로 위 조각의 가로 범위 안에서 시작하면 같은 칸으로 본다
    const pieces = content.items
      .filter((item) => item.str && item.str.trim())
      .map((item) => ({ x: item.transform[4], y: item.transform[5], w: item.width || 0, h: item.height || 10, text: item.str.trim() }))
      .sort((a, b) => b.y - a.y || a.x - b.x);
    const groups = [];
    pieces.forEach((piece) => {
      const group = groups.find((g) => {
        const last = g[g.length - 1];
        const gap = last.y - piece.y;
        return gap > 0 && gap <= last.h * 1.8 && piece.x >= last.x - 2 && piece.x <= last.x + last.w + 2;
      });
      if (group) group.push(piece);
      else groups.push([piece]);
    });
    groups.filter((g) => g.length > 1).forEach((g) => lines.push(g.map((p) => p.text).join("")));
    // 사진(이미지) 개수: 교육 증빙사진 확인용
    const ops = await page.getOperatorList();
    const imageOps = IMAGE_OPS();
    const pageImages = ops.fnArray.filter((fn) => imageOps.includes(fn)).length;
    images += pageImages;
    pages.push({ lines, images: pageImages });
  }
  return { lines: pages.flatMap((page) => page.lines), pages, tables: [], images };
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
  return { lines, pages: [], tables, images: 0 };
}

// ── Word(docx)·PowerPoint(pptx) ──
// 둘 다 zip 안의 XML이다. zip은 SheetJS에 들어 있는 CFB로 풀고, XML은 브라우저 DOMParser로 읽는다.
function zipEntries(arrayBuffer) {
  const container = XLSX.CFB.read(new Uint8Array(arrayBuffer), { type: "array" });
  const entries = new Map();
  container.FullPaths.forEach((path, index) => {
    const file = container.FileIndex[index];
    if (file && file.type === 2 && file.content) entries.set(path.replace(/^[^/]*\//, ""), file.content);
  });
  return entries;
}

function zipXml(entries, path) {
  const content = entries.get(path);
  if (!content) return null;
  return new DOMParser().parseFromString(new TextDecoder("utf-8").decode(content), "application/xml");
}

function textOf(node, textTag) {
  return [...node.getElementsByTagName(textTag)].map((t) => t.textContent).join("").replace(/\s+/g, " ").trim();
}

function mediaCount(entries, prefix) {
  return [...entries.keys()].filter((path) => path.startsWith(prefix) && /\.(png|jpe?g|gif|bmp|emf|wmf|tiff?)$/i.test(path)).length;
}

function readDocx(arrayBuffer) {
  const entries = zipEntries(arrayBuffer);
  const doc = zipXml(entries, "word/document.xml");
  if (!doc) throw new Error("Word 문서 내용을 찾지 못했습니다.");
  const body = doc.getElementsByTagName("w:body")[0];
  const pages = [{ lines: [], images: 0 }];
  const tables = [];
  const pushLine = (line) => line && pages[pages.length - 1].lines.push(line);
  // 본문 순서대로: 문단은 한 줄, 표는 행마다 한 줄(칸은 공백으로 잇고, 이름 찾기용 표로도 둔다)
  [...body.children].forEach((node) => {
    if (node.tagName === "w:p") {
      const breaksPage = node.getElementsByTagName("w:lastRenderedPageBreak").length
        || [...node.getElementsByTagName("w:br")].some((br) => br.getAttribute("w:type") === "page");
      if (breaksPage && pages[pages.length - 1].lines.length) pages.push({ lines: [], images: 0 });
      pushLine(textOf(node, "w:t"));
    } else if (node.tagName === "w:tbl") {
      const rows = [...node.getElementsByTagName("w:tr")].map((tr) =>
        [...tr.getElementsByTagName("w:tc")].map((tc) => textOf(tc, "w:t")));
      tables.push(rows);
      rows.forEach((row) => pushLine(row.filter(Boolean).join(" ")));
    }
  });
  const images = mediaCount(entries, "word/media/");
  return { lines: pages.flatMap((p) => p.lines), pages, tables, images };
}

function readPptx(arrayBuffer) {
  const entries = zipEntries(arrayBuffer);
  const slidePaths = [...entries.keys()]
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort((a, b) => Number(a.match(/(\d+)\.xml$/)[1]) - Number(b.match(/(\d+)\.xml$/)[1]));
  if (!slidePaths.length) throw new Error("슬라이드를 찾지 못했습니다.");
  const pages = [];
  const tables = [];
  let images = 0;
  slidePaths.forEach((path) => {
    const slide = zipXml(entries, path);
    const lines = [];
    // 표(a:tbl)는 행 단위로, 나머지 글상자는 문단(a:p) 단위로 읽는다
    const tableNodes = [...slide.getElementsByTagName("a:tbl")];
    tableNodes.forEach((tbl) => {
      const rows = [...tbl.getElementsByTagName("a:tr")].map((tr) =>
        [...tr.getElementsByTagName("a:tc")].map((tc) => textOf(tc, "a:t")));
      tables.push(rows);
      rows.forEach((row) => row.some(Boolean) && lines.push(row.filter(Boolean).join(" ")));
    });
    [...slide.getElementsByTagName("a:p")].forEach((p) => {
      if (tableNodes.some((tbl) => tbl.contains(p))) return;
      const line = textOf(p, "a:t");
      if (line) lines.push(line);
    });
    // 슬라이드에 놓인 그림 수(교육 사진 확인용)
    const pictures = slide.getElementsByTagName("p:pic").length;
    images += pictures;
    pages.push({ lines, images: pictures });
  });
  return { lines: pages.flatMap((p) => p.lines), pages, tables, images: images || mediaCount(entries, "ppt/media/") };
}

const IMAGE_FILE = /\.(jpe?g|png|gif|bmp|webp|heic)$/i;

// ── 그림 속 글자 인식(OCR): Safe_b 화면 캡처, 자격증 사진 등 ──
// tesseract.js를 lib/tesseract에 두고 브라우저 안에서만 인식한다(그림을 밖으로 보내지 않음).
// 웹 주소(http/https)로 열었을 때 동작한다. 파일을 더블클릭해 연 화면(file://)에서는 브라우저가 막는다.
let ocrWorkerPromise = null;
let ocrProgress = null;

function getOcrWorker() {
  if (!ocrWorkerPromise) {
    const base = new URL("lib/tesseract/", location.href).href;
    ocrWorkerPromise = Tesseract.createWorker(["kor", "eng"], 1, {
      workerPath: `${base}worker.min.js`,
      corePath: `${base}core`,
      langPath: `${base}lang`,
      gzip: true,
      logger: (message) => ocrProgress && ocrProgress(message),
    }).catch((error) => {
      ocrWorkerPromise = null;
      throw error;
    });
  }
  return ocrWorkerPromise;
}

// OCR은 '김 민 수'처럼 한 글자씩 띄우는 경우가 있어, 한 글자짜리가 이어지면 붙인다
function joinSpacedSyllables(line) {
  return line.replace(/(?:^|(?<=\s))([가-힣])(?:\s([가-힣]))+(?=\s|$)/g, (match) => match.replace(/\s/g, ""));
}

// 화면 캡처는 글자가 작아(14px 안팎) 그대로는 한글을 거의 못 읽는다 → 2~3배 키우고 흑백으로 바꿔 인식한다
const OCR_TARGET_WIDTH = 2400;

async function prepareForOcr(file) {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(3, Math.max(1, OCR_TARGET_WIDTH / bitmap.width));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  const context = canvas.getContext("2d");
  context.imageSmoothingQuality = "high";
  context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  const image = context.getImageData(0, 0, canvas.width, canvas.height);
  const px = image.data;
  const { width, height } = canvas;
  // 흑백 두 값으로 나눈다(true = 검은 점)
  const dark = new Uint8Array(width * height);
  for (let i = 0, p = 0; i < px.length; i += 4, p += 1) {
    dark[p] = 0.299 * px[i] + 0.587 * px[i + 1] + 0.114 * px[i + 2] < 165 ? 1 : 0;
  }
  // 표 테두리처럼 가로·세로로 길게 이어진 선은 지운다(남겨 두면 표 전체를 글자로 잘못 읽는다)
  const maxRun = Math.round(40 * scale);
  const erase = new Uint8Array(width * height);
  for (let y = 0; y < height; y += 1) {
    let start = -1;
    for (let x = 0; x <= width; x += 1) {
      const on = x < width && dark[y * width + x];
      if (on && start < 0) start = x;
      if (!on && start >= 0) {
        if (x - start > maxRun) for (let k = start; k < x; k += 1) erase[y * width + k] = 1;
        start = -1;
      }
    }
  }
  for (let x = 0; x < width; x += 1) {
    let start = -1;
    for (let y = 0; y <= height; y += 1) {
      const on = y < height && dark[y * width + x];
      if (on && start < 0) start = y;
      if (!on && start >= 0) {
        if (y - start > maxRun) for (let k = start; k < y; k += 1) erase[k * width + x] = 1;
        start = -1;
      }
    }
  }
  for (let p = 0; p < dark.length; p += 1) {
    const value = dark[p] && !erase[p] ? 0 : 255;
    px[p * 4] = px[p * 4 + 1] = px[p * 4 + 2] = value;
    px[p * 4 + 3] = 255;
  }
  context.putImageData(image, 0, 0);
  return canvas;
}

async function ocrImage(file, onProgress) {
  ocrProgress = onProgress;
  try {
    const worker = await getOcrWorker();
    const { data } = await worker.recognize(await prepareForOcr(file));
    return data.text.split("\n").map((line) => joinSpacedSyllables(line.replace(/\s+/g, " ").trim())).filter(Boolean);
  } finally {
    ocrProgress = null;
  }
}

async function readDocument(file, options = {}) {
  const name = file.name;
  let part;
  if (IMAGE_FILE.test(name)) {
    part = { lines: [], pages: [], tables: [], images: 1 };
    if (options.ocr) {
      try {
        part.lines = await ocrImage(file, options.onProgress);
        part.ocr = true;
      } catch (error) {
        throw new Error(location.protocol === "file:"
          ? "그림 글자 인식은 웹 주소로 열었을 때만 됩니다(파일로 연 화면에서는 브라우저가 막음)."
          : `그림 글자 인식에 실패했습니다(${error.message || error}).`);
      }
    }
  } else {
    const buffer = await file.arrayBuffer();
    if (/\.pdf$/i.test(name)) part = await readPdf(buffer);
    else if (/\.(xlsx|xlsm|xls)$/i.test(name)) part = readXlsx(buffer);
    else if (/\.docx$/i.test(name)) part = readDocx(buffer);
    else if (/\.pptx$/i.test(name)) part = readPptx(buffer);
    else if (/\.(doc|ppt)$/i.test(name)) throw new Error("옛 형식(.doc·.ppt)은 읽을 수 없습니다. Word·PowerPoint에서 docx·pptx 또는 PDF로 다시 저장해 주세요.");
    else throw new Error("PDF, 엑셀, Word(docx), PowerPoint(pptx), 사진 파일만 읽을 수 있습니다.");
  }
  // 파일 이름도 단서로 쓴다(예: '지게차운전기능사_홍길동.jpg', '화기감시자_지정서.pdf')
  return { ...part, fileName: name, lines: [`[파일] ${name.replace(/[_\-.]+/g, " ")}`, ...part.lines] };
}

// ── 글자에서 정보 뽑기 ──
const DATE_PATTERNS = [
  /(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})\s*일?/g,
  /(?<![\d.])(\d{2})\s*[./]\s*(\d{1,2})\s*[./]\s*(\d{1,2})(?![\d.])/g,
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
// '투입 인원 4명 (김민수, 이준호)'처럼 명단이 한 줄 안에 있는 경우
const INLINE_NAME_LIST = /(투입\s*인원|작업\s*인원|작업자\s*명단|참여\s*자|참석\s*자|이수\s*자)/;

function isName(word) {
  const rules = window.PERMIT_RULES;
  return NAME_LIKE.test(word) && !rules.nameStopwords.includes(word)
    && !(word.length >= 3 && new RegExp(rules.nameStopSuffix).test(word));
}

function namesInText(text) {
  return text.split(/[\s,()·ㆍ:：/]+/).filter(isName);
}

function findNames(lines, tables) {
  const names = new Set();
  // 엑셀: '성명' 같은 머리글 열 아래의 값
  tables.forEach((rows) => {
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
  // PDF(또는 머리글을 못 찾은 엑셀): 명단 줄 아래·괄호 안의 2~4글자 한글 단어
  if (!names.size) {
    let inList = false;
    lines.forEach((line) => {
      // 괄호 안 명단 '(김민수, 이준호)': 쉼표로 나눈 항목이 모두 이름 한 단어일 때만 명단으로 본다
      // ('(안전대 착용, 안전고리 체결)' 같은 문구는 제외)
      (line.match(/\(([^)]*)\)/g) || []).forEach((group) => {
        const items = group.slice(1, -1).split(/[,，·ㆍ]/).map((item) => item.trim()).filter(Boolean);
        const allNames = items.length && items.every((item) => isName(item));
        if (allNames && (items.length >= 2 || INLINE_NAME_LIST.test(line))) items.forEach((n) => names.add(n));
      });
      if (/성\s*명|이\s*름|참석자|이수자|명\s*단/.test(line)) {
        inList = true;
        return;
      }
      if (!inList) return;
      if (/합\s*계|총\s*원|위와\s*같이|확인\s*합니다/.test(line)) {
        inList = false;
        return;
      }
      namesInText(line).forEach((n) => names.add(n));
    });
  }
  return [...names];
}

function regexes(keywords) {
  return keywords.map((k) => new RegExp(k, "i"));
}

function findLine(lines, keywords) {
  const patterns = regexes(keywords);
  // 본문 줄을 먼저 보고, 없으면 파일 이름 줄
  const isFileLine = (line) => line.startsWith("[파일]");
  const hit = lines.filter((l) => !isFileLine(l)).find((line) => patterns.some((p) => p.test(line)))
    || lines.filter(isFileLine).find((line) => patterns.some((p) => p.test(line)));
  if (hit) return hit;
  // PDF 표의 좁은 칸은 '투입 인 / 원'처럼 줄이 나뉘므로, 공백·줄바꿈을 지운 글자에서 한 번 더 찾는다
  const compact = lines.filter((l) => !isFileLine(l)).join("").replace(/\s+/g, "");
  for (const pattern of patterns) {
    const m = compact.match(pattern);
    if (m) return compact.slice(Math.max(0, m.index - 10), m.index + m[0].length + 40);
  }
  return null;
}

function hasAny(text, keywords) {
  const compact = text.replace(/\s+/g, "");
  return regexes(keywords).some((p) => p.test(text) || p.test(compact));
}

function detectWorkTypes(lines) {
  return window.PERMIT_RULES.workTypes.filter((type) => findLine(lines, type.keywords));
}

function analyzePart(part) {
  const text = part.lines.join("\n");
  const rules = window.PERMIT_RULES;
  return {
    ...part,
    text,
    dates: findDates(text),
    names: findNames(part.lines, part.tables),
    hours: findHours(text),
    workTypes: detectWorkTypes(part.lines),
    category: {
      daily: hasAny(text, rules.workerCategory.daily),
      newHire: hasAny(text, rules.workerCategory.newHire),
      shortTerm: hasAny(text, rules.workerCategory.shortTerm),
    },
  };
}

function analyzeDocument(kind, parts) {
  const files = parts.map(analyzePart);
  const lines = files.flatMap((f) => f.lines);
  const doc = analyzePart({ lines, pages: files.flatMap((f) => f.pages), tables: files.flatMap((f) => f.tables),
    images: files.reduce((sum, f) => sum + f.images, 0) });
  return { ...doc, kind, files };
}

// ── 작업계획서 작성 대상 판단(회사 기준) ──
function planTargetOf(type, text) {
  if (type.planTarget === true) return { target: true, note: "" };
  if (type.planTarget === "depth") {
    const depths = [...text.matchAll(/깊이\s*[:：]?\s*(\d+(?:\.\d+)?)\s*(?:m|M|미터)|(\d+(?:\.\d+)?)\s*(?:m|M|미터)\s*(?:이상\s*)?(?:깊이|굴착|터파기)/g)]
      .map((m) => Number(m[1] || m[2]));
    if (!depths.length) return { target: true, note: "굴착 깊이를 못 읽어 2m 이상으로 봅니다." };
    const max = Math.max(...depths);
    return { target: max >= 2, note: `굴착 깊이 ${max}m` };
  }
  if (type.planTarget === "voltage") {
    const volts = [...text.matchAll(/(\d+(?:\.\d+)?)\s*(kV|KV|kv|V|볼트)(?![a-zA-Z])/g)]
      .map((m) => Number(m[1]) * (/k/i.test(m[2]) ? 1000 : 1));
    if (!volts.length) return { target: true, note: "전압을 못 읽어 50V 이상으로 봅니다." };
    const max = Math.max(...volts);
    return { target: max >= 50, note: `전압 ${max >= 1000 ? `${max / 1000}kV` : `${max}V`}` };
  }
  return { target: false, note: "" };
}

// ── 검토 ──
function snippet(line) {
  if (!line) return "";
  const clean = line.replace(/^\[파일\]\s*/, "파일: ");
  return clean.length > 80 ? `${clean.slice(0, 80)}…` : clean;
}

function formatList(values, limit = 8) {
  if (!values.length) return "없음";
  return values.length > limit ? `${values.slice(0, limit).join(", ")} 외 ${values.length - limit}` : values.join(", ");
}

function reviewDocuments(docs) {
  const rules = window.PERMIT_RULES;
  const results = [];
  const add = (section, doc, title, status, detail, basis = "") => results.push({ section, doc, title, status, detail, basis });
  const checkItems = (section, doc, items, basis, missingStatus = "fail") => items.forEach((item) => {
    const line = findLine(doc.lines, item.keywords);
    add(section, DOC_KINDS[doc.kind], item.name, line ? "ok" : missingStatus,
      line ? `“${snippet(line)}”` : "해당 내용을 찾지 못했습니다.", basis);
  });

  const { permit, plan, risk, education, attachments } = docs;
  const workDocs = [permit, plan].filter(Boolean);
  const workText = workDocs.map((d) => d.text).join("\n");
  const workLines = workDocs.flatMap((d) => d.lines);
  // 작업 종류는 허가서 기준(없으면 계획서). 계획서의 안전대책 문장 속 단어로 작업을 늘리지 않는다.
  const workTypes = detectWorkTypes((permit || plan || { lines: [] }).lines);
  const has = (id) => workTypes.some((t) => t.id === id);
  // 자격서류·지정서·도면은 계획서뿐 아니라 허가서·기타 첨부에 따로 붙어 올 수 있다
  const evidenceDocs = [attachments, plan, permit].filter(Boolean);
  const evidenceLines = evidenceDocs.flatMap((d) => d.lines);

  const planTargets = workTypes
    .map((type) => ({ type, ...planTargetOf(type, workText) }))
    .filter((t) => t.target);
  const educationTypes = workTypes.filter((t) => t.education);

  const workDates = permit && permit.dates.length ? permit.dates : plan ? plan.dates : [];
  const workStart = workDates[0] || null;
  const workers = [...new Set(workDocs.flatMap((d) => d.names))];

  // 1) 제출 서류
  const S1 = "제출 서류";
  add(S1, DOC_KINDS.permit, "작업허가서", permit ? "ok" : "fail", permit ? "파일을 읽었습니다." : "파일이 없습니다.");
  if (planTargets.length) {
    add(S1, DOC_KINDS.plan, "작업계획서(작성 대상)", plan ? "ok" : "fail",
      `${plan ? "첨부됨" : "첨부 필요"} — 작성 대상: ${planTargets.map((t) => t.type.name + (t.note ? `(${t.note})` : "")).join(", ")}`,
      rules.planBasis);
  } else {
    add(S1, DOC_KINDS.plan, "작업계획서", plan ? "ok" : "info", plan ? "파일을 읽었습니다." : "회사 기준 작성 대상 작업이 확인되지 않았습니다.");
  }
  add(S1, DOC_KINDS.risk, "위험성평가", risk ? "ok" : "fail", risk ? "파일을 읽었습니다." : "파일이 없습니다.", rules.riskBasis);
  if (educationTypes.length) {
    add(S1, DOC_KINDS.education, "특별안전교육(교육 대상 작업)", education ? "ok" : "fail",
      `${education ? `${education.files.length}개 파일` : "실시 결과 첨부 필요"} — 대상: ${educationTypes.map((t) => `${t.education.no}. ${t.education.name}`).join(" / ")}`,
      rules.educationBasis);
  } else {
    add(S1, DOC_KINDS.education, "특별안전교육", education ? "ok" : "info", education ? "파일을 읽었습니다." : "특별교육 대상 작업이 확인되지 않았습니다.");
  }
  if (attachments) add(S1, DOC_KINDS.attachments, "기타 첨부", "info", formatList(attachments.files.map((f) => f.fileName)));

  // 2) 작업 정보
  const S2 = "작업 정보";
  add(S2, workDocs.map((d) => DOC_KINDS[d.kind]).join("·") || "-", "작업 종류", workTypes.length ? "info" : "warn",
    workTypes.length ? workTypes.map((t) => t.name).join(", ") : "작업 종류를 알아내지 못했습니다. 작업허가서에 작업 내용이 있는지 확인하세요.");

  // 3) 작업계획서
  const S3 = "작업계획서";
  if (plan) {
    checkItems(S3, plan, rules.planCommon, rules.planBasis);
    rules.planSigners.forEach((signer) => {
      const line = findLine(plan.lines, signer.keywords);
      add(S3, DOC_KINDS.plan, signer.name, line ? "ok" : "fail",
        line ? `“${snippet(line)}”` : "서명란을 찾지 못했습니다. 도급사 안전관리총괄책임자·협력사 안전보건관리책임자 서명은 필수입니다.");
    });
    add(S3, DOC_KINDS.plan, "모든 서명란 서명 여부", "manual", "서명은 글자로 읽을 수 없습니다. 원본에서 빈 서명란이 없는지 확인하세요.");

    // 작업지휘자: 지정 + 한 페이지 + 서명
    const supervisorLine = findLine(plan.lines, rules.supervisor.keywords);
    add(S3, DOC_KINDS.plan, "작업지휘자 지정", supervisorLine ? "ok" : "fail",
      supervisorLine ? `“${snippet(supervisorLine)}”` : "작업지휘자가 지정되어 있지 않습니다(필수).", rules.planBasis);
    if (supervisorLine && plan.pages.length) {
      const page = plan.pages.findIndex((p) => hasAny(p.lines.join(" "), rules.supervisor.keywords) && hasAny(p.lines.join(" "), rules.supervisor.pageKeywords));
      add(S3, DOC_KINDS.plan, "작업지휘자 지정 페이지", page >= 0 ? "ok" : "warn",
        page >= 0 ? `${page + 1}쪽에 작업지휘자 지정 내용이 있습니다. 서명은 원본에서 확인하세요.`
          : "작업지휘자 '지정' 페이지를 찾지 못했습니다. 계획서 안에 한 페이지로 포함되고 서명되었는지 확인하세요.");
    }

    // 차량계 건설기계·하역운반기계 사용 작업
    const machineTypes = planTargets.filter((t) => t.type.machine).map((t) => t.type);
    const handling = machineTypes.find((t) => t.id === "handlingMachine");
    const manualOnly = handling && !findLine(workLines, handling.machineKeywords) && !machineTypes.some((t) => t.id === "constructionMachine");
    if (machineTypes.length && !manualOnly) {
      const signal = findLine(plan.lines, rules.signalman);
      add(S3, DOC_KINDS.plan, "[장비 사용] 신호수 지정", signal ? "ok" : "fail",
        signal ? `“${snippet(signal)}”` : "신호수가 지정되어 있지 않습니다(인력운반만 하는 경우는 불필요).", rules.planBasis);
      const drawing = findLine(evidenceLines, rules.workDrawing);
      add(S3, DOC_KINDS.plan, "[장비 사용] 작업계획도(동선도)", drawing ? "ok" : "fail",
        drawing ? `“${snippet(drawing)}”` : "중량물 이동동선이 담긴 작업계획도를 찾지 못했습니다.", rules.planBasis);
      add(S3, DOC_KINDS.plan, "[장비 사용] 작업계획도 표기 내용", "manual",
        "작업계획도에 작업지휘자, 안전시설물, 신호수, 진입경로가 표기되어 있는지 확인하세요.");
      rules.qualifications.forEach((qual) => {
        const equipment = findLine(workLines, qual.detect);
        if (!equipment) return;
        const document = findLine(evidenceLines, qual.documents);
        add(S3, "작업계획서·기타 첨부", `[자격서류] ${qual.equipment}`, document ? "ok" : "fail",
          document ? `“${snippet(document)}”` : `${qual.label} 첨부가 필요합니다.`, rules.planBasis);
      });
    } else if (manualOnly) {
      add(S3, DOC_KINDS.plan, "[중량물] 인력운반", "info", "장비 없이 인력운반으로 보여 신호수·자격서류는 확인하지 않았습니다.");
    }
    if (handling) {
      const binding = findLine(plan.lines, rules.binding);
      add(S3, DOC_KINDS.plan, "[중량물] 취급·결속방법", binding ? "ok" : "warn",
        binding ? `“${snippet(binding)}”` : "중량물 결속방법(줄걸이·슬링 등)을 찾지 못했습니다.", rules.planBasis);
      add(S3, DOC_KINDS.plan, "[중량물] 결속방법 적절성", "manual", "결속방법이 중량물의 형태·무게에 적절한지 확인하세요.");
    }

    // 전기 작업: 도면
    if (planTargets.some((t) => t.type.id === "electric")) {
      const drawing = findLine(evidenceLines, rules.electricDrawing);
      add(S3, DOC_KINDS.plan, "[전기] 전기도면(단선도·결선도)", drawing ? "ok" : "fail",
        drawing ? `“${snippet(drawing)}”` : "전기도면·단선도·결선도를 찾지 못했습니다.", rules.planBasis);
      add(S3, DOC_KINDS.plan, "[전기] 도면 형태", "manual",
        "분전반·차단기를 기호와 선으로 표시한 설계 도면인지 확인하세요. 실물 사진만 첨부한 경우 반려입니다.");
    }

    // 작업 종류별 세부 항목
    workTypes.filter((t) => t.planItems).forEach((type) => {
      checkItems(S3, plan, type.planItems.map((item) => ({ ...item, name: `[${type.name}] ${item.name}` })), type.basis, "warn");
    });
  }

  // 4) 화기 작업: 화기감시자 지정서
  if (has("hot")) {
    const S4 = "화기 작업";
    const watcherFile = [...evidenceDocs.flatMap((d) => d.files)]
      .find((f) => hasAny(f.text, rules.fireWatcher.keywords) && hasAny(f.text, rules.fireWatcher.appointment));
    add(S4, "기타 첨부", "화기감시자 지정서", watcherFile ? "ok" : "fail",
      watcherFile ? `${watcherFile.fileName}` : "화기감시자 지정서를 찾지 못했습니다(별도 파일로 첨부).",
      "「산업안전보건기준에 관한 규칙」 제241조의2(화재감시자)");
    if (watcherFile) {
      // '화기감시자 성명: 정우진' / '화재감시자 정우진' 처럼 바로 뒤의 이름, 없으면 지정서의 성명 칸 명단
      const fromLabel = [...watcherFile.text.matchAll(/화[기재]\s*감시자\s*(?:성\s*명|이\s*름)?\s*[:：]?\s*([가-힣]{2,4})(?![가-힣])/g)]
        .map((m) => m[1]).filter(isName);
      const watchers = [...new Set(fromLabel.length ? fromLabel : watcherFile.names)];
      const hotEducationNames = (education ? education.files : [])
        .filter((f) => hasAny(f.text, rules.workTypes.find((t) => t.id === "hot").educationKeywords))
        .flatMap((f) => f.names);
      if (!watchers.length) {
        add(S4, "화기감시자 ↔ 화기 교육 이수자", "화기감시자가 화기 특별교육 이수자인지", "warn",
          "지정서에서 화기감시자 이름을 읽지 못했습니다. 화기(화재위험) 특별교육 이수자 중에서 지정했는지 확인하세요.");
      } else {
        const missing = watchers.filter((n) => !hotEducationNames.includes(n));
        add(S4, "화기감시자 ↔ 화기 교육 이수자", "화기감시자가 화기 특별교육 이수자인지", missing.length ? "fail" : "ok",
          missing.length ? `화기 특별교육 이수자 명단에 없음: ${missing.join(", ")}` : `${watchers.join(", ")} — 화기 교육 이수자입니다.`);
      }
    }
  }

  // 5) 위험성평가
  const S5 = "위험성평가";
  if (risk) {
    checkItems(S5, risk, rules.riskItems, rules.riskBasis);
    const levelCells = risk.tables.flat().filter((c) => /^(상|중|하)$/.test(c)).length
      + risk.lines.filter((l) => /(^|\s)(상|중|하)(\s|$)/.test(l)).length;
    add(S5, DOC_KINDS.risk, "위험도 상·중·하 표기", levelCells ? "ok" : "fail",
      levelCells ? `상·중·하 표기 ${levelCells}곳` : "위험도가 상·중·하로 표기되어 있지 않습니다.", rules.riskBasis);
    add(S5, DOC_KINDS.risk, "평가 일자", risk.dates.length ? "ok" : "fail",
      risk.dates.length ? `문서의 날짜: ${formatList(risk.dates, 4)}` : "날짜를 찾지 못했습니다.", rules.riskBasis);
    // 계획서·교육 기준 작업이 위험성평가에 들어 있는지
    const coverTypes = [...new Map([...workTypes, ...(education ? education.workTypes : [])].map((t) => [t.id, t])).values()];
    coverTypes.forEach((type) => {
      const line = findLine(risk.lines.filter((l) => !/^\[파일\]/.test(l)), type.keywords);
      add(S5, DOC_KINDS.risk, `작업 포함: ${type.name}`, line ? "ok" : "fail",
        line ? `“${snippet(line)}”` : "작업계획서·특별교육의 이 작업이 위험성평가에 없습니다.", rules.riskBasis);
    });
    const titleLine = plan ? findLine(plan.lines, ["작업\\s*명", "공사\\s*명"]) : null;
    if (titleLine) {
      const words = titleLine.replace(/작업\s*명|공사\s*명/, "").split(/[\s,()·ㆍ]+/).filter((w) => w.length >= 2 && !/작업|공사|및/.test(w));
      const shared = words.filter((w) => risk.text.includes(w));
      add(S5, "작업계획서 ↔ 위험성평가", "작업 내용 반영", shared.length ? "ok" : "warn",
        shared.length ? `계획서 작업명의 '${formatList(shared, 5)}'이(가) 평가표에 있습니다.` : "계획서 작업명의 내용이 위험성평가에 보이지 않습니다. 대략적인 작업 내용이 반영됐는지 확인하세요.");
    }
  }

  // 6) 특별안전교육
  const S6 = "특별안전교육";
  if (educationTypes.length && education) {
    // 작업별로 교육일지가 따로 있어야 한다(화기+전기 → 2개)
    const matchedFiles = new Map();
    educationTypes.forEach((type) => {
      const files = education.files.filter((f) => hasAny(f.lines.join(" "), type.educationKeywords));
      matchedFiles.set(type.id, files);
      add(S6, DOC_KINDS.education, `교육일지: ${type.education.no}. ${type.education.name}`, files.length ? "ok" : "fail",
        files.length ? formatList(files.map((f) => f.fileName)) : "이 작업의 특별교육 실시 결과가 없습니다.", rules.educationBasis);
    });
    if (educationTypes.length > 1) {
      const used = new Set();
      let shared = false;
      educationTypes.forEach((type) => {
        const files = matchedFiles.get(type.id).filter((f) => !used.has(f.fileName));
        if (files.length) used.add(files[0].fileName);
        else if (matchedFiles.get(type.id).length) shared = true;
      });
      add(S6, DOC_KINDS.education, "작업별 교육일지 따로 등록", shared ? "fail" : "ok",
        shared ? "여러 작업의 교육이 한 일지로만 확인됩니다. 작업별(예: 화기·전기) 교육일지를 각각 등록해야 합니다."
          : `작업 ${educationTypes.length}종 각각 교육일지가 있습니다.`, rules.educationBasis);
    }
  } else if (educationTypes.length) {
    add(S6, DOC_KINDS.education, "특별교육 실시 결과", "fail",
      educationTypes.map((t) => `${t.education.no}. ${t.education.name}`).join(" / "), rules.educationBasis);
  }

  if (education) {
    const hireDocs = [education, attachments].filter(Boolean).flatMap((d) => d.files);
    const hasHireEducation = (hours) => hireDocs.some((f) => hasAny(f.text, rules.hireEducation)
      && (!hours || (f.hours !== null && f.hours >= hours)));
    education.files.forEach((file) => {
      if (!file.lines.length || file.lines.length <= 1) {
        if (file.images) add(S6, file.fileName, "교육 사진", "info", "사진 파일입니다.");
        return;
      }
      const label = `[${file.fileName}]`;
      rules.educationItems.forEach((item) => {
        const line = findLine(file.lines, item.keywords);
        add(S6, file.fileName, `${label} ${item.name}`, line ? "ok" : "fail", line ? `“${snippet(line)}”` : "해당 내용을 찾지 못했습니다.", rules.educationBasis);
      });
      add(S6, file.fileName, `${label} 교육일`, file.dates.length ? "ok" : "fail",
        file.dates.length ? `문서의 날짜: ${formatList(file.dates, 4)}` : "교육일을 찾지 못했습니다.", rules.educationBasis);

      // 교육시간: 정규 16h / 일용 2h + 채용 시 교육 1h / 신규 정규 4h + 채용 시 교육 8h(3개월 내 12h) / 단기간·간헐 2h
      const h = file.hours;
      const hrs = rules.educationHours;
      let status = "fail";
      let detail = "교육시간을 찾지 못했습니다.";
      if (h !== null) {
        if (h >= hrs.full) {
          status = "ok";
          detail = `${h}시간 (16시간 이상: 공통 8시간 + 작업 8시간)`;
        } else if (file.category.daily && h >= hrs.daily) {
          status = hasHireEducation(1) ? "ok" : "fail";
          detail = `${h}시간 (일용근로자 2시간) — 일용직 신규채용자 교육 1시간 이수 서류 ${status === "ok" ? "있음" : "필요"}`;
        } else if (file.category.newHire && h >= hrs.newHire) {
          status = hasHireEducation(8) ? "warn" : "fail";
          detail = `${h}시간 (신규 정규직 4시간) — 신규채용자 교육 8시간 이수 서류 ${status === "warn" ? "있음" : "필요"}, 3개월 이내 12시간 특별교육 추가 필요`;
        } else if (file.category.shortTerm && h >= hrs.daily) {
          status = "manual";
          detail = `${h}시간 (단기간·간헐적 작업 2시간) — 근로자가 아닌 회사 기준으로 단기간(2개월 이내 일회성)·간헐적(연 합계 60일 이하) 작업인지 확인`;
        } else if (h >= hrs.daily) {
          status = "fail";
          detail = `${h}시간 — 16시간 미만인데 일용직·신규 정규직·단기간 작업 여부가 적혀 있지 않습니다.`;
        } else {
          detail = `${h}시간 — 최소 2시간 미만`;
        }
      }
      add(S6, file.fileName, `${label} 교육시간`, status, detail, rules.educationBasis);

      // 증빙사진
      add(S6, file.fileName, `${label} 교육 증빙사진`, file.images ? "ok" : "fail",
        file.images ? `사진(이미지) ${file.images}장` : "교육 사진을 찾지 못했습니다.", rules.educationBasis);
      add(S6, file.fileName, `${label} 사진 확인`, "manual",
        `교육 대상 ${file.names.length ? `${file.names.length}명(강사 제외)` : "인원 미확인"} — ${rules.photoChecks.join(" / ")}`);
    });
  }

  // 7) 서류 간 일치
  const S7 = "서류 간 일치";
  if (workStart) {
    if (plan && permit && permit.dates.length && plan.dates.length) {
      const inRange = workDates.every((d) => d >= plan.dates[0] && d <= plan.dates[plan.dates.length - 1]);
      add(S7, "작업허가서 ↔ 작업계획서", "작업일이 작업계획 기간 안에 있음", inRange ? "ok" : "warn",
        `작업일 ${formatList(workDates, 3)} / 계획서 날짜 ${plan.dates[0]} ~ ${plan.dates[plan.dates.length - 1]}`);
    }
    if (risk && risk.dates.length) {
      add(S7, "작업일 ↔ 위험성평가", "위험성평가를 작업 전에 실시", risk.dates[0] <= workStart ? "ok" : "fail",
        `평가일(가장 이른 날짜) ${risk.dates[0]} / 작업 시작 ${workStart}`, rules.riskBasis);
    }
    if (education) {
      education.files.filter((f) => f.dates.length).forEach((file) => {
        add(S7, "작업일 ↔ 특별안전교육", `[${file.fileName}] 작업 전 교육`, file.dates[0] <= workStart ? "ok" : "fail",
          `교육일 ${file.dates[0]} / 작업 시작 ${workStart} — 최초 작업 전에 해야 합니다.`, rules.educationBasis);
      });
    }
  } else if (workDocs.length) {
    add(S7, "작업허가서", "작업일", "warn", "작업일을 찾지 못해 날짜 전후 관계를 확인하지 못했습니다.");
  }

  if (permit) {
    const permitIds = permit.workTypes.map((t) => t.id);
    const extra = [plan, risk].filter(Boolean)
      .flatMap((doc) => doc.workTypes.filter((t) => !permitIds.includes(t.id)).map((t) => `${t.name}(${DOC_KINDS[doc.kind]})`));
    if (extra.length) {
      add(S7, "작업허가서 ↔ 계획서·위험성평가", "허가서에 없는 작업이 다른 서류에 있음", "warn",
        `${[...new Set(extra)].join(", ")} — 위험요인 언급일 수도 있으니, 실제로 하는 작업이면 허가 구분에 넣었는지 확인하세요.`);
    }
  }

  if (education && educationTypes.length) {
    const educated = education.names;
    if (workers.length && educated.length) {
      const missing = workers.filter((name) => !educated.includes(name));
      add(S7, "작업자 명단 ↔ 교육 이수자", "투입 작업자 전원 특별교육 이수", missing.length ? "fail" : "ok",
        missing.length ? `교육 명단에 없는 작업자 ${missing.length}명: ${formatList(missing)}` : `작업자 ${workers.length}명이 모두 교육 명단에 있습니다.`,
        rules.educationBasis);
    } else {
      add(S7, "작업자 명단 ↔ 교육 이수자", "투입 작업자 전원 특별교육 이수", "warn",
        `이름을 읽지 못해 대조하지 못했습니다(작업자 ${workers.length}명, 교육 명단 ${educated.length}명). 직접 확인하세요.`, rules.educationBasis);
    }
  }
  if (risk && workers.length && risk.names.length) {
    const missing = workers.filter((name) => !risk.names.includes(name));
    add(S7, "작업자 명단 ↔ 위험성평가 참여자", "작업자의 위험성평가 공유·참여", missing.length ? "warn" : "ok",
      missing.length ? `참여 명단에 없는 작업자: ${formatList(missing)}` : "작업자 모두 참여 명단에 있습니다.", rules.riskBasis);
  }

  return {
    results,
    summary: {
      workTypes: workTypes.map((t) => t.name),
      planTargets: planTargets.map((t) => t.type.name),
      educationTypes: educationTypes.map((t) => `${t.education.no}. ${t.education.name}`),
      workDates,
      workers,
    },
  };
}

window.PermitReview = { DOC_KINDS, readDocument, analyzeDocument, reviewDocuments };
