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
    // 좁은 머리글 칸('투입 인 / 원')만 잇는다: 긴 본문이나 ☑·☐ 표시가 섞이면 엉뚱한 글자가 붙으므로 제외
    groups
      .filter((g) => g.length > 1 && g.every((p) => p.text.length <= 12 && !/[☑☒■▣●◉✔✓✅☐□○◯]/.test(p.text)))
      .forEach((g) => lines.push(g.map((p) => p.text).join("")));
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

// ── 압축 파일(zip) 풀기 ──
// 한국어 Windows에서 만든 zip은 파일 이름이 CP949라 그대로 풀면 한글이 깨진다 → zip 목록을 직접 읽어 이름을 판별한다.
// 압축 해제는 브라우저 내장 DecompressionStream(deflate-raw)을 쓴다.
const ZIP_SKIP = /(^|\/)(__MACOSX\/|\.DS_Store$|Thumbs\.db$|desktop\.ini$)/i;
const MAX_ZIP_DEPTH = 2;

function decodeZipName(bytes, utf8Flag) {
  if (utf8Flag) return new TextDecoder("utf-8").decode(bytes);
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch (error) {
    return new TextDecoder("euc-kr").decode(bytes);
  }
}

async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

// 결과: { files: [File(.zipPath, .origin)], skipped: ['이름(이유)'] }
async function unzipFile(file, depth = 0) {
  const buffer = new Uint8Array(await file.arrayBuffer());
  const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  let end = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65557); i -= 1) {
    if (view.getUint32(i, true) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) throw new Error("압축 파일을 읽지 못했습니다(zip 형식만 지원).");

  const count = view.getUint16(end + 10, true);
  let pos = view.getUint32(end + 16, true);
  const result = { files: [], skipped: [] };
  const origin = file.origin ? `${file.origin} › ${file.zipPath || file.name}` : file.name;
  for (let n = 0; n < count && view.getUint32(pos, true) === 0x02014b50; n += 1) {
    const flags = view.getUint16(pos + 8, true);
    const method = view.getUint16(pos + 10, true);
    const compressedSize = view.getUint32(pos + 20, true);
    const nameLength = view.getUint16(pos + 28, true);
    const extraLength = view.getUint16(pos + 30, true);
    const commentLength = view.getUint16(pos + 32, true);
    const localOffset = view.getUint32(pos + 42, true);
    const path = decodeZipName(buffer.subarray(pos + 46, pos + 46 + nameLength), flags & 0x800);
    pos += 46 + nameLength + extraLength + commentLength;

    if (path.endsWith("/") || ZIP_SKIP.test(path)) continue;
    const baseName = path.split("/").pop();
    if (flags & 0x1) {
      result.skipped.push(`${baseName}(암호)`);
      continue;
    }
    if (method !== 0 && method !== 8) {
      result.skipped.push(`${baseName}(지원하지 않는 압축 방식)`);
      continue;
    }
    const start = localOffset + 30 + view.getUint16(localOffset + 26, true) + view.getUint16(localOffset + 28, true);
    const raw = buffer.subarray(start, start + compressedSize);
    const bytes = method === 8 ? await inflateRaw(raw) : raw;
    const inner = new File([bytes], baseName, { type: guessType(baseName) });
    inner.zipPath = path;
    inner.origin = origin;
    if (/\.zip$/i.test(baseName) && depth < MAX_ZIP_DEPTH) {
      const nested = await unzipFile(inner, depth + 1);
      result.files.push(...nested.files);
      result.skipped.push(...nested.skipped);
    } else {
      result.files.push(inner);
    }
  }
  return result;
}

function guessType(name) {
  const ext = (name.split(".").pop() || "").toLowerCase();
  return { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp",
    pdf: "application/pdf" }[ext] || "";
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

// 화면 캡처는 글자가 작아(11~14px) 그대로는 한글을 거의 못 읽는다 → 2~3배 키우고 흑백으로 바꿔 인식한다
const OCR_TARGET_WIDTH = 5600;

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
  // Safe_B는 선택된 버튼을 파란 글자·테두리로 보여 주므로 흑백으로 바꾸기 전의 색을 남겨 둔다
  const color = new Uint8ClampedArray(px);
  const { width, height } = canvas;
  // 흑백 두 값으로 나눈다(1 = 검은 점). 입력 칸의 옅은 회색 글자·파란 글자도 살리도록,
  // 고정 기준 대신 주변(창 크기 win) 평균 밝기보다 충분히 어두우면 글자로 본다(국소 문턱값)
  // 밝기는 가장 어두운 색 성분으로 잰다: 파란 글자(선택된 버튼)도 검은 글자처럼 진하게 남는다
  const lum = new Float32Array(width * height);
  for (let i = 0, p = 0; i < px.length; i += 4, p += 1) lum[p] = Math.min(px[i], px[i + 1], px[i + 2]);
  const integral = new Float64Array((width + 1) * (height + 1));
  for (let y = 0; y < height; y += 1) {
    let rowSum = 0;
    for (let x = 0; x < width; x += 1) {
      rowSum += lum[y * width + x];
      integral[(y + 1) * (width + 1) + x + 1] = integral[y * (width + 1) + x + 1] + rowSum;
    }
  }
  const win = Math.max(8, Math.round(12 * scale));
  const dark = new Uint8Array(width * height);
  for (let y = 0; y < height; y += 1) {
    const ya = Math.max(0, y - win);
    const yb = Math.min(height, y + win + 1);
    for (let x = 0; x < width; x += 1) {
      const xa = Math.max(0, x - win);
      const xb = Math.min(width, x + win + 1);
      const sum = integral[yb * (width + 1) + xb] - integral[ya * (width + 1) + xb] - integral[yb * (width + 1) + xa] + integral[ya * (width + 1) + xa];
      const mean = sum / ((yb - ya) * (xb - xa));
      const v = lum[y * width + x];
      dark[y * width + x] = v < 235 && v < mean * 0.86 ? 1 : 0;
    }
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
  // 글자 인식에는 흑백 대신 부드러운 회색 그림을 준다(작고 굵은 글자가 흑백으로는 뭉개짐). 표 테두리 자리만 하얗게 지운다.
  for (let p = 0; p < dark.length; p += 1) {
    const erased = erase[p] && dark[p];
    dark[p] = dark[p] && !erase[p] ? 1 : 0;
    const value = erased ? 255 : Math.min(255, lum[p] * 1.08);
    px[p * 4] = px[p * 4 + 1] = px[p * 4 + 2] = value;
    px[p * 4 + 3] = 255;
  }
  context.putImageData(image, 0, 0);
  return { canvas, dark, color, width, height, sourceWidth: bitmap.width };
}

// ── Safe_B 선택 버튼(칩) 찾기 ──
// Safe_B의 예/아니오·작업구분은 '✓ 이름' 모양의 둥근 버튼이고, 고른 버튼만 파란 테두리·파란 글자다.
// 안 고른 버튼은 테두리가 거의 보이지 않아, 넓은 빈칸으로 버튼을 나누고 글자가 파란지로 선택 여부를 정한다.
// 고른 버튼은 테두리·글자가 파랗다(0.6 이상). 안 고른 버튼은 글자 색 번짐 정도(0.25 이하)라 0.4로 가른다.
const CHIP_SELECTED_BLUE = 0.4;
// 파란 점: 파랑이 빨강·초록보다 18 이상 높음(선택 버튼의 옅은 파란 글자·테두리까지 센다).
// 검은 글자 가장자리의 색 번짐도 조금 섞이지만(0.2 안팎), 선택 버튼(0.6 이상)과는 CHIP_SELECTED_BLUE로 갈린다.
const isBluish = (r, g, b) => b - Math.max(r, g) > 18;

// 버튼 테두리 상자 찾기: 안 고른 버튼 테두리는 아주 옅은 회색(밝기 237~244)이라 247 미만을 '그려진 점'으로 본다.
// 한 줄 띠에서 열마다 그려진 점이 있는지 보면 버튼 하나는 테두리 때문에 끊기지 않고 이어지고, 버튼 사이는 빈 열로 끊긴다.
// 위·아래 테두리 줄이 있고 안이 대부분 비어 있는(바탕색 칸이 아닌) 구간만 버튼으로 본다.
// centerY: 그 줄 글자의 세로 가운데, size: 화면 전체의 보통 한글 글자 높이
// 버튼들의 위 테두리는 한 높이에 나란히 놓인 가로선이다. 글자 위쪽(글자 높이 0.55~1.6배)에서 가장 많이 그려진 줄을
// 위 테두리 줄로 잡고, 그 줄을 따라 이어진 구간 하나하나를 버튼으로 본 뒤, 아래 테두리가 있는지 확인한다.
function detectChipBoxes(image, centerY, size) {
  const { color, width, height } = image;
  const drawn = (x, y) => {
    const i = (y * width + x) * 4;
    return Math.min(color[i], color[i + 1], color[i + 2]) < 247;
  };
  // 그려진 점 수. 표의 칸 선(글자 높이 40배보다 길게 이어진 줄)은 버튼 테두리가 아니므로 0으로 본다
  const rowCount = (y) => {
    let n = 0;
    let run = 0;
    for (let x = 0; x < width; x += 1) {
      if (drawn(x, y)) {
        n += 1;
        run += 1;
        if (run > size * 40) return 0;
      } else {
        run = 0;
      }
    }
    return n;
  };
  const bestRow = (from, to) => {
    let best = -1;
    let bestCount = 0;
    for (let y = Math.max(0, Math.round(from)); y <= Math.min(height - 1, Math.round(to)); y += 1) {
      const n = rowCount(y);
      if (n > bestCount) {
        best = y;
        bestCount = n;
      }
    }
    return best;
  };
  const top = bestRow(centerY - size * 1.6, centerY - size * 0.55);
  const bottom = bestRow(centerY + size * 0.55, centerY + size * 1.6);
  if (top < 0 || bottom < 0 || bottom - top < size * 1.2 || bottom - top > size * 2.8) return [];
  // 위 테두리 줄(두께를 감안해 ±2줄)을 따라 이어진 구간 = 버튼 하나(작은 틈 2px까지는 이어진 것으로)
  const onTop = (x) => {
    for (let y = top - 2; y <= top + 2; y += 1) if (y >= 0 && y < height && drawn(x, y)) return true;
    return false;
  };
  const segments = [];
  let start = -1;
  let lastOn = -1;
  for (let x = 0; x <= width; x += 1) {
    if (x < width && onTop(x)) {
      if (start < 0) start = x;
      lastOn = x;
    } else if (start >= 0 && x - lastOn > 2) {
      segments.push([start, lastOn]);
      start = -1;
    }
  }
  const coverage = (y, xa, xb) => {
    let best = 0;
    for (let yy = y - 2; yy <= y + 2; yy += 1) {
      if (yy < 0 || yy >= height) continue;
      let n = 0;
      for (let x = xa; x <= xb; x += 1) n += drawn(x, yy) ? 1 : 0;
      best = Math.max(best, n / (xb - xa + 1));
    }
    return best;
  };
  const boxes = [];
  segments.forEach(([x0, x1]) => {
    // 버튼 너비: 글자 높이 1.5배~40배(표 칸 선처럼 아주 긴 줄은 버튼이 아니다)
    if (x1 - x0 < size * 1.5 || x1 - x0 > size * 40) return;
    const ia = x0 + Math.round(size * 0.2);
    const ib = x1 - Math.round(size * 0.2);
    if (ib <= ia || coverage(bottom, ia, ib) < 0.8) return;
    let inside = 0;
    let insideTotal = 0;
    let ink = 0;
    let blue = 0;
    for (let y = top; y <= bottom; y += 1) {
      for (let x = x0; x <= x1; x += 1) {
        const i = (y * width + x) * 4;
        const r = color[i];
        const g = color[i + 1];
        const b = color[i + 2];
        const isDrawn = Math.min(r, g, b) < 247;
        if (y > top + 2 && y < bottom - 2 && x > x0 + 2 && x < x1 - 2) {
          insideTotal += 1;
          if (isDrawn) inside += 1;
        }
        if (isDrawn) {
          ink += 1;
          if (isBluish(r, g, b)) blue += 1;
        }
      }
    }
    if (insideTotal && inside / insideTotal > 0.5) return; // 바탕색이 칠해진 칸(항목 이름 칸 등)
    boxes.push({ x0, x1, y0: top, y1: bottom, blue: ink ? blue / ink : 0 });
  });
  return boxes;
}

// 단어 글자 획 중 파란색 비율(0~1): Safe_B에서 선택된 버튼 글자는 파랗고, 선택 안 된 것은 회색·검정이다
function blueRatio(image, bbox) {
  const { color, width, height } = image;
  const x0 = Math.max(0, Math.floor(bbox.x0));
  const x1 = Math.min(width - 1, Math.ceil(bbox.x1));
  const y0 = Math.max(0, Math.floor(bbox.y0));
  const y1 = Math.min(height - 1, Math.ceil(bbox.y1));
  let ink = 0;
  let blue = 0;
  for (let y = y0; y <= y1; y += 1) {
    for (let x = x0; x <= x1; x += 1) {
      const i = (y * width + x) * 4;
      const r = color[i];
      const g = color[i + 1];
      const b = color[i + 2];
      if (0.299 * r + 0.587 * g + 0.114 * b > 200) continue; // 바탕(흰색·옅은 파랑)
      ink += 1;
      if (isBluish(r, g, b)) blue += 1;
    }
  }
  return ink ? blue / ink : 0;
}

// ── 그림 속 체크박스·라디오 버튼 ──
// 글자 인식은 ☑/☐, ●/○를 제대로 읽지 못하고 'ㅁ', '00', '10', '@' 같은 짧은 엉뚱한 단어로 읽는다.
// 그런 단어를 '선택 표시'로 보고, 그 상자 안이 얼마나 검게 채워졌는지(채움 비율)로 선택 여부를 판단한다.
// 테두리는 둘 다 있으므로 표시의 '가운데'만 본다: ☐·○는 가운데가 비어 있고, ☑·●·(체크된) 파란 체크박스는 차 있다.
// 판단은 같은 묶음(작업구분 칸들, 예/아니오 한 쌍)끼리 비교해서 한다.
const MARKER_WORD = /^(10|[ㅁㅇ口0Oo@©®◎●○☐☑☒■□▣◉✓✔vVxX[\]]{1,3})$/;
const MARKER_CENTER = 0.3; // 상자 가장자리에서 30%씩 안쪽

// 글자 인식이 준 상자는 표시보다 길쭉하거나 어긋나므로, 그 근처에서 글자 크기(size)만 한 네모·동그라미 덩어리 중
// 인식 상자 가운데에 가장 가까운 것을 표시의 실제 테두리로 본다
// 영역 안의 '이어진 검은 점 덩어리'마다 테두리 상자를 돌려준다
function inkComponents(image, region) {
  const { dark, width, height } = image;
  const x0 = Math.max(0, Math.floor(region.x0));
  const x1 = Math.min(width - 1, Math.ceil(region.x1));
  const y0 = Math.max(0, Math.floor(region.y0));
  const y1 = Math.min(height - 1, Math.ceil(region.y1));
  const regionWidth = x1 - x0 + 1;
  const seen = new Uint8Array(regionWidth * (y1 - y0 + 1));
  const components = [];
  for (let y = y0; y <= y1; y += 1) {
    for (let x = x0; x <= x1; x += 1) {
      const local = (y - y0) * regionWidth + (x - x0);
      if (seen[local] || !dark[y * width + x]) continue;
      let bx0 = x, bx1 = x, by0 = y, by1 = y;
      const stack = [[x, y]];
      seen[local] = 1;
      while (stack.length) {
        const [cx, cy] = stack.pop();
        if (cx < bx0) bx0 = cx;
        if (cx > bx1) bx1 = cx;
        if (cy < by0) by0 = cy;
        if (cy > by1) by1 = cy;
        for (const [nx, ny] of [[cx + 1, cy], [cx - 1, cy], [cx, cy + 1], [cx, cy - 1]]) {
          if (nx < x0 || nx > x1 || ny < y0 || ny > y1) continue;
          const nLocal = (ny - y0) * regionWidth + (nx - x0);
          if (!seen[nLocal] && dark[ny * width + nx]) {
            seen[nLocal] = 1;
            stack.push([nx, ny]);
          }
        }
      }
      components.push({ x0: bx0, x1: bx1, y0: by0, y1: by1 });
    }
  }
  return components;
}

function markerShaped(box, size) {
  const w = box.x1 - box.x0 + 1;
  const h = box.y1 - box.y0 + 1;
  return w >= size * 0.5 && h >= size * 0.5 && w <= size * 1.4 && h <= size * 1.4 && w / h >= 0.7 && w / h <= 1.45;
}

function markerOutline(image, bbox, size) {
  const pad = size * 0.6;
  const cx = (bbox.x0 + bbox.x1) / 2;
  const cy = (bbox.y0 + bbox.y1) / 2;
  const candidates = inkComponents(image, { x0: bbox.x0 - pad, x1: bbox.x1 + pad, y0: bbox.y0 - pad, y1: bbox.y1 + pad })
    .filter((box) => markerShaped(box, size));
  let best = null;
  let bestDistance = Infinity;
  candidates.forEach((box) => {
    const distance = Math.hypot((box.x0 + box.x1) / 2 - cx, (box.y0 + box.y1) / 2 - cy);
    if (distance < bestDistance) {
      best = box;
      bestDistance = distance;
    }
  });
  return best || bbox;
}

function boxDensity(image, box) {
  const { dark, width } = image;
  let ink = 0;
  for (let y = box.y0; y <= box.y1; y += 1) for (let x = box.x0; x <= box.x1; x += 1) ink += dark[y * width + x];
  return ink / ((box.x1 - box.x0 + 1) * (box.y1 - box.y0 + 1));
}

// 네 변이 모두 검은 네모 테두리인지(체크박스): 안에 체크가 있든 없든 테두리로 알아본다
function isSquareFrame(image, box) {
  const { dark, width } = image;
  const band = Math.max(1, Math.round((box.x1 - box.x0) * 0.12));
  const edgeInk = (x0, x1, y0, y1) => {
    let ink = 0;
    let total = 0;
    for (let y = y0; y <= y1; y += 1) {
      for (let x = x0; x <= x1; x += 1) {
        total += 1;
        ink += dark[y * width + x];
      }
    }
    return ink / total;
  };
  // 각 변 띠(두께 band)에서, 그 줄에 검은 점이 하나라도 있는 비율
  const rowCovered = (y0, y1) => {
    let covered = 0;
    for (let x = box.x0; x <= box.x1; x += 1) {
      let any = 0;
      for (let y = y0; y <= y1; y += 1) any |= dark[y * width + x];
      covered += any;
    }
    return covered / (box.x1 - box.x0 + 1);
  };
  const colCovered = (x0, x1) => {
    let covered = 0;
    for (let y = box.y0; y <= box.y1; y += 1) {
      let any = 0;
      for (let x = x0; x <= x1; x += 1) any |= dark[y * width + x];
      covered += any;
    }
    return covered / (box.y1 - box.y0 + 1);
  };
  return rowCovered(box.y0, box.y0 + band) > 0.85 && rowCovered(box.y1 - band, box.y1) > 0.85
    && colCovered(box.x0, box.x0 + band) > 0.85 && colCovered(box.x1 - band, box.x1) > 0.85
    && edgeInk(box.x0 + band + 1, box.x1 - band - 1, box.y0 + band + 1, box.y1 - band - 1) < 0.6;
}

// 글자 인식이 아예 단어로 잡지 못한 표시도 찾는다: 글자 크기만 한 체크박스 테두리, 빈 동그라미, 꽉 찬 점
function scanMarkers(image, lineBox, size) {
  return inkComponents(image, { x0: lineBox.x0 - size * 2, x1: lineBox.x1 + size, y0: lineBox.y0 - size * 0.3, y1: lineBox.y1 + size * 0.3 })
    .filter((box) => {
      const w = box.x1 - box.x0 + 1;
      const h = box.y1 - box.y0 + 1;
      if (w < size * 0.7 || h < size * 0.7 || w > size * 1.3 || h > size * 1.3 || w / h < 0.8 || w / h > 1.25) return false;
      const whole = boxDensity(image, box);
      const center = inkDensity(image, box, size, true);
      return whole > 0.75 || (center < 0.03 && whole > 0.12 && whole < 0.6) || isSquareFrame(image, box);
    });
}

function inkDensity(image, wordBox, size, exact = false) {
  const { dark, width, height } = image;
  const bbox = exact ? wordBox : markerOutline(image, wordBox, size);
  const insetX = (bbox.x1 - bbox.x0) * MARKER_CENTER;
  const insetY = (bbox.y1 - bbox.y0) * MARKER_CENTER;
  const x0 = Math.max(0, Math.floor(bbox.x0 + insetX));
  const x1 = Math.min(width - 1, Math.ceil(bbox.x1 - insetX));
  const y0 = Math.max(0, Math.floor(bbox.y0 + insetY));
  const y1 = Math.min(height - 1, Math.ceil(bbox.y1 - insetY));
  let ink = 0;
  for (let y = y0; y <= y1; y += 1) for (let x = x0; x <= x1; x += 1) ink += dark[y * width + x];
  return ink / Math.max(1, (x1 - x0 + 1) * (y1 - y0 + 1));
}

// ── Safe_B 항목 이름 칸(옅은 파란 바탕) 찾기 ──
// 화면 전체를 한 번에 읽으면 작고 굵은 항목 이름(관리팀·작업인원 등)을 자주 놓친다.
// 항목 이름 칸과 그 오른쪽 값 칸을 따로 한 줄씩 읽으면 훨씬 정확하다.
function findLabelCells(image, scale) {
  const { color, width, height } = image;
  // 칸 사이 1px 경계선을 건너뛰지 않도록 원래 화면의 1px 간격으로 본다(건너뛰면 위아래 칸이 한 칸으로 붙는다)
  const step = Math.max(1, Math.floor(scale));
  const gw = Math.floor(width / step);
  const gh = Math.floor(height / step);
  const isLabelBg = (x, y) => {
    const i = (y * width + x) * 4;
    const r = color[i];
    const g = color[i + 1];
    const b = color[i + 2];
    // 일반 항목 칸 (239,242,255), 필수(*) 항목 칸 (221,228,254)
    return b >= 248 && r >= 212 && r <= 246 && g >= 220 && g <= 249 && b - r >= 9;
  };
  const mask = new Uint8Array(gw * gh);
  for (let gy = 0; gy < gh; gy += 1) for (let gx = 0; gx < gw; gx += 1) mask[gy * gw + gx] = isLabelBg(gx * step, gy * step) ? 1 : 0;
  const seen = new Uint8Array(gw * gh);
  const cells = [];
  for (let start = 0; start < mask.length; start += 1) {
    if (!mask[start] || seen[start]) continue;
    let x0 = gw, x1 = 0, y0 = gh, y1 = 0, count = 0;
    const stack = [start];
    seen[start] = 1;
    while (stack.length) {
      const p = stack.pop();
      const gx = p % gw;
      const gy = (p - gx) / gw;
      count += 1;
      if (gx < x0) x0 = gx;
      if (gx > x1) x1 = gx;
      if (gy < y0) y0 = gy;
      if (gy > y1) y1 = gy;
      for (const q of [p - 1, p + 1, p - gw, p + gw]) {
        if (q < 0 || q >= mask.length || seen[q] || !mask[q]) continue;
        if (Math.abs((q % gw) - gx) > 1) continue;
        seen[q] = 1;
        stack.push(q);
      }
    }
    const w = (x1 - x0 + 1) * step;
    const h = (y1 - y0 + 1) * step;
    // 항목 이름 칸 크기(원래 화면 기준 가로 60px·세로 18px 이상), 대부분이 바탕색
    if (w < 60 * scale || h < 18 * scale || w > width * 0.4 || count / ((x1 - x0 + 1) * (y1 - y0 + 1)) < 0.6) continue;
    cells.push({ x0: x0 * step, x1: (x1 + 1) * step, y0: y0 * step, y1: (y1 + 1) * step });
  }
  return cells.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
}

async function readLabelCells(worker, image, scale) {
  const cells = findLabelCells(image, scale);
  const pad = Math.round(4 * scale);
  const result = [];
  const read = async (box, psm) => {
    await worker.setParameters({ tessedit_pageseg_mode: psm });
    const { data } = await worker.recognize(image.canvas, {
      rectangle: { left: Math.round(box.x0), top: Math.round(box.y0), width: Math.round(box.x1 - box.x0), height: Math.round(box.y1 - box.y0) },
    });
    return data.text.replace(/\s+/g, " ").trim();
  };
  try {
    for (const cell of cells) {
      // 같은 줄(세로로 절반 이상 겹침)에서 오른쪽으로 가장 가까운 다음 항목 이름 칸 전까지가 값 칸
      const next = cells
        .filter((c) => c !== cell && c.x0 > cell.x1 && Math.min(c.y1, cell.y1) - Math.max(c.y0, cell.y0) > (cell.y1 - cell.y0) * 0.5)
        .sort((a, b) => a.x0 - b.x0)[0];
      const valueBox = { x0: cell.x1 + pad, x1: (next ? next.x0 : image.width) - pad, y0: cell.y0 + pad, y1: cell.y1 - pad };
      if (valueBox.x1 - valueBox.x0 < 10 * scale) continue;
      const labelBox = { x0: cell.x0 + pad, x1: cell.x1 - pad, y0: cell.y0 + pad, y1: cell.y1 - pad };
      // 한 줄로 못 읽으면(빈 글자) 덩어리 글로 한 번 더
      let label = await read(labelBox, "7");
      if (!/[가-힣A-Za-z]{2,}/.test(label)) label = await read(labelBox, "6");
      const value = await read(valueBox, "6");
      result.push({ label, value });
    }
  } finally {
    await worker.setParameters({ tessedit_pageseg_mode: "6" }); // tesseract.js 기본값(한 덩어리 글)으로 되돌린다
  }
  return result;
}

async function ocrImage(file, onProgress) {
  ocrProgress = onProgress;
  try {
    const worker = await getOcrWorker();
    const image = await prepareForOcr(file);
    const { data } = await worker.recognize(image.canvas, {}, { blocks: true, text: true });
    const lines = [];
    const options = [];
    const wordColors = [];
    const records = [];
    (data.blocks || []).forEach((block) => block.paragraphs.forEach((paragraph) => paragraph.lines.forEach((ocrLine) => {
      const words = ocrLine.words.filter((w) => w.text.trim());
      if (!words.length) return;
      // 줄 글자를 만들면서 단어마다 글자 위치를 기억한다(선택지가 어느 항목 뒤에 있는지 찾는 데 씀)
      // 한 글자씩 띄워 읽힌 한글('긴 급 작 업')은 붙여서 만든다(위치가 어긋나지 않도록 여기서 바로)
      let text = "";
      const single = (w) => /^[가-힣]$/.test(w.text.trim());
      words.forEach((word, i) => {
        const separator = !text ? "" : (single(word) && single(words[i - 1]) ? "" : " ");
        word.offset = text.length + separator.length;
        text = `${text}${separator}${word.text.trim()}`;
      });
      // 선택 표시로 읽힌 짧은 단어 → 그 뒤 ~ 다음 표시 앞까지의 단어가 선택지 글자
      const markerIndexes = words.map((w, i) => (MARKER_WORD.test(w.text.trim()) ? i : -1)).filter((i) => i >= 0);
      // 이 줄의 보통 글자 높이(표시 단어는 상자가 길쭉하게 나와 빼고 잰다)
      const textHeights = words.filter((w) => !MARKER_WORD.test(w.text.trim())).map((w) => w.bbox.y1 - w.bbox.y0).sort((a, b) => a - b);
      const size = textHeights.length ? textHeights[Math.floor(textHeights.length / 2)] : 30;
      words.forEach((word) => { word.blue = blueRatio(image, word.bbox); });
      records.push({ text, words, size, ocrLine });
    })));

    // Safe_B 버튼(칩): 테두리 상자로 찾는다(글자 인식이 ✓·버튼 이름을 엉뚱하게 읽어도 상자는 정확하다)
    // 글자 높이는 화면 전체의 한글 단어(두 글자 이상)로 잰다(줄마다 재면 들쭉날쭉하다)
    const median = (values) => values.sort((a, b) => a - b)[Math.floor(values.length / 2)];
    const hangulWords = records.flatMap((r) => r.words.filter((w) => /[가-힣]{2,}/.test(w.text)));
    const pageTextHeight = hangulWords.length ? median(hangulWords.map((w) => w.bbox.y1 - w.bbox.y0)) : 30;
    records.forEach((record) => {
      const centers = record.words.filter((w) => /[가-힣A-Za-z0-9]{2,}/.test(w.text)).map((w) => (w.bbox.y0 + w.bbox.y1) / 2);
      const centerY = centers.length ? median(centers) : (record.ocrLine.bbox.y0 + record.ocrLine.bbox.y1) / 2;
      record.chipBoxes = detectChipBoxes(image, centerY, pageTextHeight);
    });

    // 한 줄 안에 파란 버튼과 파랗지 않은 버튼이 같이 있으면 Safe_B 버튼 화면으로 본다(선택 = 파란 테두리·글자)
    const chipStyle = records.some((r) => r.chipBoxes.some((b) => b.blue >= CHIP_SELECTED_BLUE)
      && r.chipBoxes.some((b) => b.blue < 0.25));

    records.forEach(({ text, words, size, ocrLine, chipBoxes }) => {
      words.forEach((word) => wordColors.push({
        lineText: text, start: word.offset, end: word.offset + word.text.trim().length, blue: Math.round(word.blue * 100) / 100,
      }));
      lines.push(text);
      if (chipStyle && chipBoxes.length) {
        chipBoxes.forEach((box) => {
          const inside = words.filter((w) => {
            const cx = (w.bbox.x0 + w.bbox.x1) / 2;
            return cx > box.x0 && cx < box.x1;
          }).sort((a, b) => a.offset - b.offset);
          // 버튼 안 글자를 못 읽었으면, 버튼 바로 오른쪽 단어 위치를 기준으로 삼는다(예/아니오 구간 판단용)
          const anchor = inside[0] || [...words].sort((a, b) => a.bbox.x0 - b.bbox.x0).find((w) => w.bbox.x0 >= box.x0);
          const last = inside[inside.length - 1];
          options.push({
            text: inside.length ? text.slice(inside[0].offset, last.offset + last.text.trim().length).replace(/^[vV✓✔+=|ㅜㅠ~＊*<«,"'。ㆍ[`^\s]+/, "") : "",
            textIndex: anchor ? anchor.offset : text.length,
            x: box.x0,
            density: null,
            filled: box.blue >= CHIP_SELECTED_BLUE,
            blue: Math.round(box.blue * 100) / 100,
            chip: true,
            check: true,
            lineText: text,
            page: 0,
          });
        });
        return;
      }
      const markerIndexes = words.map((w, i) => (MARKER_WORD.test(w.text.trim()) ? i : -1)).filter((i) => i >= 0);

      // 표시 위치: 표시로 읽힌 단어의 실제 테두리 + 단어로 잡히지 않은 표시(겹치면 하나로)
      const overlaps = (a, b) => a.x0 < b.x1 && a.x1 > b.x0 && a.y0 < b.y1 && a.y1 > b.y0;
      const markers = markerIndexes.map((i) => markerOutline(image, words[i].bbox, size));
      scanMarkers(image, ocrLine.bbox, size).forEach((box) => {
        if (!markers.some((m) => overlaps(m, box))) markers.push(box);
      });
      markers.sort((a, b) => a.x0 - b.x0);
      const isMarkerWord = (w) => MARKER_WORD.test(w.text.trim()) || markers.some((m) => overlaps(m, w.bbox));
      markers.forEach((marker, n) => {
        const nextX = n + 1 < markers.length ? markers[n + 1].x0 : Infinity;
        const following = words.filter((w) => !isMarkerWord(w) && w.bbox.x0 >= marker.x1 - size * 0.3 && w.bbox.x0 < nextX);
        // 글자 인식이 표시와 옆 글자를 한 단어('[cd')로 묶어 버리면 선택지 글자가 없다 → 위치만 남긴다(예/아니오는 순서로 판단)
        const anchor = following[0] || words.find((w) => overlaps(marker, w.bbox) || w.bbox.x0 >= marker.x0);
        if (!anchor) return;
        options.push({
          text: text.slice(anchor.offset, following.length ? following[following.length - 1].offset + following[following.length - 1].text.trim().length : anchor.offset),
          textIndex: anchor.offset,
          density: Math.round(inkDensity(image, marker, size, true) * 100) / 100,
          filled: null,
          lineText: text,
          page: 0,
        });
      });
    });
    // Safe_B 화면이면 항목 이름 칸·값 칸을 한 칸씩 다시 읽는다
    const formCells = chipStyle ? await readLabelCells(worker, image, image.width / (image.sourceWidth || image.width)) : [];
    return { lines: lines.filter(Boolean), options, wordColors, chipStyle, formCells };
  } finally {
    ocrProgress = null;
  }
}

// PDF·Word·엑셀처럼 글자로 된 서류의 ☑/☐, ●/○ 선택지
const FILLED_MARKS = "☑☒■▣●◉✔✓✅";
const EMPTY_MARKS = "☐□○◯";

function textOptions(lines, page = 0) {
  const marks = FILLED_MARKS + EMPTY_MARKS;
  const pattern = new RegExp(`([${marks}])\\s*([^${marks}]+)`, "g");
  const options = [];
  lines.forEach((line, index) => {
    for (const m of line.matchAll(pattern)) {
      options.push({ text: m[2].trim(), textIndex: m.index + m[0].indexOf(m[2]), density: null,
        filled: FILLED_MARKS.includes(m[1]), lineText: line, page });
    }
    // 줄 끝에 표시만 있고 글자는 다음 줄로 넘어간 경우('… ☐' / '지게차 작업 ☐ 덤프트럭')
    const trailing = line.match(new RegExp(`([${marks}])\\s*$`));
    if (trailing && index + 1 < lines.length) {
      const next = lines[index + 1].split(new RegExp(`[${marks}]`))[0].trim();
      if (next) options.push({ text: next, textIndex: trailing.index, density: null, filled: FILLED_MARKS.includes(trailing[1]), lineText: line, page });
    }
  });
  return options;
}

// 체크하지 않은 선택지(☐ 지게차 작업 등)의 글자를 뺀 줄: 장비·작업 찾기에 쓴다
function activeLines(doc) {
  const unchecked = (doc.options || []).filter((o) => o.filled === false && o.text);
  if (!unchecked.length) return doc.lines;
  return doc.lines.map((line) => unchecked.filter((o) => line.includes(o.text)).reduce((l, o) => l.replace(o.text, " "), line));
}

// 문서보안(DRM) 암호화 파일 알아보기: 파일 형식에 맞는 머리 글자(PDF '%PDF', docx·pptx·xlsx 'PK', 옛 Office D0CF)가
// 아니면 암호화(또는 손상)된 것이다. 'SCDSA'로 시작하면 SoftCamp 문서보안.
function detectDrm(name, buffer) {
  const head = new Uint8Array(buffer.slice(0, 1024));
  const ascii = String.fromCharCode(...head.slice(0, 8));
  if (ascii.startsWith("SCDSA")) return "문서보안(SoftCamp)";
  const text = String.fromCharCode(...head);
  if (/\.pdf$/i.test(name)) return text.includes("%PDF") ? null : "암호화 또는 손상";
  if (/\.(docx|pptx|xlsx|xlsm)$/i.test(name)) return ascii.startsWith("PK") ? null : "암호화 또는 손상";
  if (/\.(doc|ppt|xls)$/i.test(name)) return head[0] === 0xd0 && head[1] === 0xcf ? null : "암호화 또는 손상";
  return null;
}

async function readDocument(file, options = {}) {
  const name = file.name;
  let part;
  if (IMAGE_FILE.test(name)) {
    part = { lines: [], pages: [], tables: [], images: 1, options: [] };
    if (options.ocr) {
      try {
        const ocr = await ocrImage(file, options.onProgress);
        part.lines = ocr.lines;
        part.options = ocr.options;
        part.wordColors = ocr.wordColors;
        part.formCells = ocr.formCells;
        part.ocr = true;
      } catch (error) {
        throw new Error(location.protocol === "file:"
          ? "그림 글자 인식은 웹 주소로 열었을 때만 됩니다(파일로 연 화면에서는 브라우저가 막음)."
          : `그림 글자 인식에 실패했습니다(${error.message || error}).`);
      }
    }
  } else {
    const buffer = await file.arrayBuffer();
    const drm = detectDrm(name, buffer);
    if (drm) {
      // 회사 문서보안(DRM)으로 암호화된 파일: 내용은 읽을 수 없고 파일 이름만 단서로 쓴다
      return {
        lines: [`[파일] ${(file.zipPath || name).replace(/[_\-./]+/g, " ")}`], pages: [], tables: [], images: 0, options: [],
        drm, fileName: file.origin ? `${file.origin} › ${file.zipPath || name}` : name, origin: file.origin || null,
      };
    }
    if (/\.pdf$/i.test(name)) part = await readPdf(buffer);
    else if (/\.(xlsx|xlsm|xls)$/i.test(name)) part = readXlsx(buffer);
    else if (/\.docx$/i.test(name)) part = readDocx(buffer);
    else if (/\.pptx$/i.test(name)) part = readPptx(buffer);
    else if (/\.(doc|ppt)$/i.test(name)) throw new Error("옛 형식(.doc·.ppt)은 읽을 수 없습니다. Word·PowerPoint에서 docx·pptx 또는 PDF로 다시 저장해 주세요.");
    else throw new Error("PDF, 엑셀, Word(docx), PowerPoint(pptx), 사진 파일만 읽을 수 있습니다.");
    part.options = part.pages.length
      ? part.pages.flatMap((page, index) => textOptions(page.lines, index))
      : textOptions(part.lines);
  }
  // 파일 이름(압축 안 폴더 포함)도 단서로 쓴다(예: '지게차운전기능사_홍길동.jpg', '화기감시자_지정서.pdf')
  const displayName = file.origin ? `${file.origin} › ${file.zipPath || name}` : name;
  return {
    ...part,
    fileName: displayName,
    origin: file.origin || null,
    lines: [`[파일] ${(file.zipPath || name).replace(/[_\-./]+/g, " ")}`, ...part.lines],
  };
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
    // 파일 이름 줄('[파일] 이수자 확인서 pdf')은 명단의 시작으로 보지 않는다
    lines.filter((line) => !line.startsWith("[파일]")).forEach((line) => {
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

// 창 선택 캡처로 쪽마다 찍은 그림들(readDocument 결과)을 문서 하나로 합친다.
// 쪽 순서를 지켜 pages[]를 만들고, 표시(options)의 page를 쪽 번호로 바꾼다(예: 하역운반기계 계획서 '1쪽' 확인).
// 찍은 화면 자체는 교육 사진이 아니므로 images는 0으로 두고 captured 표시를 남긴다.
function mergeCapturedParts(name, parts) {
  const pages = parts.map((part) => ({ lines: part.lines.slice(1), images: 0 })); // 첫 줄은 '[파일] 캡처 이름'
  return {
    lines: [`[파일] ${name.replace(/[_\-./]+/g, " ")}`, ...pages.flatMap((page) => page.lines)],
    pages,
    tables: [],
    images: 0,
    options: parts.flatMap((part, index) => (part.options || []).map((option) => ({ ...option, page: index }))),
    wordColors: parts.flatMap((part) => part.wordColors || []),
    formCells: parts.flatMap((part) => part.formCells || []),
    ocr: true,
    captured: true,
    capturePages: parts.length,
    fileName: `📷 ${name} (캡처 ${parts.length}쪽)`,
    origin: null,
  };
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
    images: files.reduce((sum, f) => sum + f.images, 0), options: files.flatMap((f) => f.options || []),
    wordColors: files.flatMap((f) => f.wordColors || []), formCells: files.flatMap((f) => f.formCells || []) });
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
  // 작업 종류는 허가서 기준(없으면 계획서). 계획서의 안전대책 문장 속 단어로 작업을 늘리지 않는다.
  // Safe_B 작업구분 칸은 고르지 않은 작업도 화면에 다 적혀 있으므로, 그 줄은 빼고 '선택된' 작업구분만 더한다.
  // 계획서의 체크하지 않은 장비(☐ 지게차 작업 등)도 장비 찾기에서 뺀다.
  const form = permit ? readPermitForm(permit) : null;
  const permitLines = permit ? permit.lines.filter((line) => !form.categoryLines.has(line)) : [];
  const workLines = [...permitLines, ...(plan ? activeLines(plan) : [])];
  const workTypes = (() => {
    if (!form) return detectWorkTypes(plan ? activeLines(plan) : []);
    const detected = detectWorkTypes(permitLines);
    const ids = new Set([...detected.map((t) => t.id), ...permitWorkTypeIds(form)]);
    return rules.workTypes.filter((t) => ids.has(t.id));
  })();
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

  // 작업허가서(Safe_B) 칸 점검
  if (form) reviewPermitForm(form, docs, add);

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
    // 일지와 따로 넣은 사진 파일(같은 압축 안 또는 같은 칸)
    const photoFiles = education.files.filter((f) => f.lines.length <= 1 && f.images);
    const logFiles = education.files.filter((f) => f.lines.length > 1);
    education.files.forEach((file) => {
      if (!file.lines.length || file.lines.length <= 1) return;
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

      // 증빙사진: 일지 안의 사진 → 같은 압축 안의 사진 파일 → (일지가 하나뿐이면) 칸에 따로 넣은 사진 파일
      const samePack = file.origin ? photoFiles.filter((p) => p.origin === file.origin) : [];
      const loose = logFiles.length === 1 ? photoFiles : photoFiles.filter((p) => !p.origin);
      let photoStatus = "fail";
      let photoDetail = "교육 사진을 찾지 못했습니다.";
      if (file.captured) {
        // 화면을 찍어 넣은 일지: 일지 안 사진은 글자 인식으로 셀 수 없다
        photoStatus = "manual";
        photoDetail = "화면 캡처로 넣은 일지 — 일지 안의 교육 사진은 원본 문서에서 직접 확인하세요";
      } else if (file.images) {
        photoStatus = "ok";
        photoDetail = `일지 안의 사진(이미지) ${file.images}장`;
      } else if (samePack.length) {
        photoStatus = "ok";
        photoDetail = `같은 압축 파일 안의 사진 파일 ${samePack.length}장`;
      } else if (loose.length) {
        photoStatus = logFiles.length === 1 ? "ok" : "warn";
        photoDetail = `따로 넣은 사진 파일 ${loose.length}장${logFiles.length === 1 ? "" : " — 어느 교육일지의 사진인지 확인하세요"}`;
      }
      add(S6, file.fileName, `${label} 교육 증빙사진`, photoStatus, photoDetail, rules.educationBasis);
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

  applyDrmLimits(docs, results, add);

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

// 문서보안(DRM)으로 암호화된 파일이 있으면: 그 파일들을 알리고, 그 서류 내용으로 판단하는 '못 찾음' 결과는
// '미흡'이 아니라 '직접 확인'으로 바꾼다(내용을 못 읽었을 뿐, 없다고 단정할 수 없으므로). 파일 이름으로 확인된 항목은 그대로.
const DRM_SLOT_WORDS = {
  plan: ["작업계획서", "화기 작업", "화기감시자", "신호수", "작업지휘자", "자격서류", "[장비 사용]", "[중량물]", "[전기]"],
  education: ["특별안전교육", "교육"],
  risk: ["위험성평가"],
  attachments: ["기타 첨부", "자격서류", "화기감시자"],
};

function applyDrmLimits(docs, results, add) {
  const drmFiles = Object.entries(docs).flatMap(([kind, doc]) => (doc ? doc.files : [])
    .filter((f) => f.drm).map((f) => ({ kind, name: f.fileName, drm: f.drm })));
  if (!drmFiles.length) return;
  const slots = [...new Set(drmFiles.map((f) => f.kind))];
  results.forEach((r) => {
    if (!["fail", "warn"].includes(r.status)) return;
    if (!/찾지 못|없습니다|읽지 못|확인하지 못|첨부 필요|첨부되지|없음|있지 않습니다/.test(r.detail)) return;
    const hay = `${r.section} ${r.doc} ${r.title}`;
    if (!slots.some((kind) => (DRM_SLOT_WORDS[kind] || []).some((word) => hay.includes(word)))) return;
    r.status = "manual";
    r.detail += " — 암호화(DRM) 파일이 있어 내용을 확인하지 못했습니다. 문서를 열어 직접 확인하거나, 해당 쪽을 '창 선택 캡처'로 넣어 주세요.";
  });
  add("제출 서류", "암호화 파일", `암호화(DRM) 파일 ${drmFiles.length}개 — 내용 검토 불가`, "warn",
    `${drmFiles.map((f) => `${f.name}(${f.drm})`).join(", ")}. 파일 이름으로 제출 여부만 확인했습니다. `
    + "내용까지 검토하려면 ① 보안 해제(복호화)한 파일을 넣거나 ② 문서를 열고 각 쪽을 '🖥 창 선택 캡처'로 넣어 주세요.");
}

window.PermitReview = { DOC_KINDS, readDocument, analyzeDocument, reviewDocuments };
