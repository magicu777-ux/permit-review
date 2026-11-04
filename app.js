// 화면: 서류 칸(드래그·선택) → 검토 → 결과 표·엑셀 저장
/* global XLSX, DOC_KINDS, readDocument, analyzeDocument, reviewDocuments */
// DOC_KINDS·readDocument 등은 review.js에서 전역으로 선언된다.

const SLOT_HINTS = {
  permit: "Safe_b 화면 캡처(여러 장 가능)",
  plan: "작업지휘자·서명·도면",
  risk: "위험요인·위험도(상중하)",
  education: "작업별 교육일지·사진",
  attachments: "자격증·화기감시자 지정서·채용시 교육",
};
// 사진을 넣을 수 있는 칸, 그중 그림 속 글자를 인식(OCR)할 칸(교육 사진은 사람 사진이라 인식하지 않음)
// 계획서·위험성평가는 압축 안에 도면·스캔 그림이 같이 오는 경우가 있어 그림도 받는다
const IMAGE_SLOTS = ["permit", "plan", "risk", "education", "attachments"];
const OCR_SLOTS = ["permit", "plan", "risk", "attachments"];
const IMAGE_NAME = /\.(jpe?g|png|gif|bmp|webp|heic)$/i;
const STATUS_LABEL = { ok: "적합", warn: "확인 필요", fail: "미흡", manual: "직접 확인", info: "참고" };
const SECTION_ORDER = ["제출 서류", "작업 정보", "작업계획서", "화기 작업", "위험성평가", "특별안전교육", "서류 간 일치"];

// 샘플(가상 서류) — samples 폴더
const SAMPLE_FILES = {
  permit: ["samples/작업허가서_캡처_샘플.png"],
  plan: ["samples/작업계획서_샘플.pdf"],
  risk: ["samples/위험성평가_샘플.xlsx"],
  education: ["samples/특별안전교육_밀폐공간_샘플.pdf"],
  attachments: ["samples/고소작업대_조종교육_이수증_샘플.pdf"],
};

const files = { permit: [], plan: [], risk: [], education: [], attachments: [] };
let lastReview = null;

const slotsEl = document.getElementById("slots");
const reviewButton = document.getElementById("reviewButton");
const statusEl = document.getElementById("status");
const resultPanel = document.getElementById("resultPanel");

function showStatus(text, isError) {
  statusEl.textContent = text;
  statusEl.classList.toggle("error", Boolean(isError));
  statusEl.classList.remove("hidden");
}

function acceptable(kind, file) {
  // .doc·.ppt(옛 형식)도 받아 두고, 읽을 때 '다시 저장' 안내를 띄운다
  return /\.(pdf|xlsx|xlsm|xls|docx|pptx|doc|ppt)$/i.test(file.name)
    || (IMAGE_SLOTS.includes(kind) && /\.(jpe?g|png|gif|bmp|webp|heic)$/i.test(file.name));
}

function renderSlots() {
  slotsEl.innerHTML = "";
  Object.entries(DOC_KINDS).forEach(([kind, label]) => {
    const slot = document.createElement("div");
    slot.className = files[kind].length ? "slot filled" : "slot";

    const title = document.createElement("div");
    title.className = "slot-title";
    title.innerHTML = `${label} <small>· ${SLOT_HINTS[kind]}</small>`;

    const input = document.createElement("input");
    input.type = "file";
    input.multiple = true;
    const documentTypes = ".pdf,.xlsx,.xlsm,.xls,.docx,.pptx,.doc,.ppt,.zip";
    input.accept = IMAGE_SLOTS.includes(kind) ? `${documentTypes},image/*` : documentTypes;
    input.id = `file-${kind}`;
    input.addEventListener("change", () => addFiles(kind, input.files));

    const pick = document.createElement("label");
    pick.className = "pick";
    pick.htmlFor = input.id;
    pick.textContent = "파일 선택 또는 끌어 놓기";

    // 작업허가서: Safe_b 화면을 캡처 도구로 찍어 바로 넣는다
    let captureButton = null;
    let windowCaptureButton = null;
    if (kind === "permit") {
      captureButton = document.createElement("button");
      captureButton.type = "button";
      captureButton.className = "capture";
      captureButton.textContent = "📷 Safe_b 화면 캡처";
      captureButton.addEventListener("click", () => startCapture(kind));
      // 캡처 도구가 막힌 PC용: 브라우저가 창을 직접 찍고, 필요한 영역을 잘라 넣는다
      windowCaptureButton = document.createElement("button");
      windowCaptureButton.type = "button";
      windowCaptureButton.className = "capture secondary";
      windowCaptureButton.textContent = "🖥 창 선택 캡처";
      windowCaptureButton.addEventListener("click", () => startWindowCapture(kind));
    }

    const list = document.createElement("ul");
    list.className = "file-list";
    files[kind].forEach((file, index) => {
      const item = document.createElement("li");
      const name = document.createElement("span");
      name.textContent = file.origin ? `📦 ${file.origin} › ${file.zipPath || file.name}` : file.name;
      if (IMAGE_NAME.test(file.name)) {
        const thumb = document.createElement("img");
        thumb.className = "thumb";
        thumb.alt = "";
        thumb.src = URL.createObjectURL(file);
        thumb.addEventListener("click", () => window.open(thumb.src, "_blank"));
        name.prepend(thumb);
      }
      const remove = document.createElement("button");
      remove.type = "button";
      remove.title = "빼기";
      remove.textContent = "✕";
      remove.addEventListener("click", () => {
        files[kind].splice(index, 1);
        renderSlots();
      });
      item.append(name, remove);
      list.appendChild(item);
    });

    slot.append(...[title, captureButton, windowCaptureButton, pick, input, list].filter(Boolean));
    ["dragenter", "dragover"].forEach((type) =>
      slot.addEventListener(type, (event) => {
        event.preventDefault();
        slot.classList.add("dragover");
      })
    );
    ["dragleave", "drop"].forEach((type) =>
      slot.addEventListener(type, () => slot.classList.remove("dragover"))
    );
    slot.addEventListener("drop", (event) => {
      event.preventDefault();
      addFiles(kind, event.dataTransfer.files);
    });
    slotsEl.appendChild(slot);
  });
  reviewButton.disabled = !Object.values(files).some((list) => list.length);
}

// 압축 파일(zip)은 넣는 즉시 풀어서 안의 파일을 하나씩 칸에 넣는다(읽을 수 없는 파일은 건너뛰고 알려 줌)
async function addFiles(kind, fileList) {
  const incoming = [...fileList];
  const added = [];
  const skipped = [];
  for (const file of incoming) {
    if (/\.zip$/i.test(file.name)) {
      showStatus(`${file.name} 압축을 푸는 중입니다…`);
      try {
        const unpacked = await unzipFile(file);
        unpacked.files.forEach((inner) => (acceptable(kind, inner) ? added.push(inner) : skipped.push(inner.name)));
        skipped.push(...unpacked.skipped);
        if (!unpacked.files.length) skipped.push(`${file.name}(빈 압축)`);
      } catch (error) {
        skipped.push(`${file.name}(${error.message || "압축을 풀지 못함"})`);
      }
    } else if (acceptable(kind, file)) {
      added.push(file);
    } else {
      skipped.push(file.name);
    }
  }
  files[kind].push(...added);
  renderSlots();
  if (skipped.length) {
    showStatus(`읽을 수 없는 파일은 건너뛰었습니다: ${skipped.join(", ")} — PDF, 엑셀, Word(docx), PowerPoint(pptx), 사진, zip만 넣을 수 있습니다.`, true);
  } else if (incoming.some((f) => /\.zip$/i.test(f.name))) {
    showStatus(`압축을 풀어 ${added.length}개 파일을 ${DOC_KINDS[kind]} 칸에 넣었습니다.`);
  }
}

// ── Safe_b 화면 캡처 ──
// 버튼 → Windows 캡처 도구(ms-screenclip:) 실행 → 사용자가 영역을 드래그 → 클립보드에 들어온 그림을
// 이 화면으로 돌아왔을 때 자동으로 가져온다. 클립보드 읽기를 막았으면 Ctrl+V로 붙여 넣으면 된다.
const CAPTURE_WAIT_MS = 3 * 60 * 1000;
const TOOL_OPEN_TIMEOUT_MS = 4000;
let capture = null;
let lastImportedSignature = null;

async function clipboardImage() {
  const items = await navigator.clipboard.read();
  for (const item of items) {
    const type = item.types.find((t) => t.startsWith("image/"));
    if (type) return item.getType(type);
  }
  return null;
}

async function blobSignature(blob) {
  if (!blob) return null;
  const head = new Uint8Array(await blob.slice(0, 2048).arrayBuffer());
  let hash = blob.size;
  head.forEach((byte) => { hash = (hash * 31 + byte) >>> 0; });
  return `${blob.size}:${hash}`;
}

function addCapture(kind, blob) {
  const now = new Date();
  const stamp = [now.getHours(), now.getMinutes(), now.getSeconds()].map((n) => String(n).padStart(2, "0")).join("");
  const extension = (blob.type.split("/")[1] || "png").replace("jpeg", "jpg");
  files[kind].push(new File([blob], `Safe_b_작업허가서_캡처_${stamp}.${extension}`, { type: blob.type }));
  capture = null;
  renderSlots();
  showStatus(`캡처한 화면을 작업허가서 칸에 넣었습니다(${files[kind].length}장). 더 캡처하거나 '검토하기'를 누르세요.`);
}

function startCapture(kind) {
  // 버튼을 누른 그 순간에 바로 실행해야 브라우저가 막지 않는다(먼저 다른 확인 창을 띄우면 막힘)
  // Windows 10·11 캡처 도구(Win+Shift+S와 같음). 브라우저가 '열기'를 물으면 허용한다.
  window.location.href = "ms-screenclip:";
  const clipboardBlocked = !navigator.clipboard || !navigator.clipboard.read;
  capture = { kind, until: Date.now() + CAPTURE_WAIT_MS, clipboardBlocked, toolOpened: false, startedAt: Date.now() };
  showStatus(clipboardBlocked
    ? "캡처 도구에서 Safe_b의 작업허가서 영역을 드래그한 뒤, 이 화면에서 Ctrl+V를 누르세요."
    : "캡처 도구에서 Safe_b의 작업허가서 영역을 드래그하세요. 이 화면으로 돌아오면 자동으로 들어옵니다.");
  if (!clipboardBlocked) waitForClipboard();
}

// 캡처 도구가 열리면 이 화면이 잠시 포커스를 잃는다 → 그 뒤에 들어온 클립보드 그림만 가져온다
window.addEventListener("blur", () => {
  if (capture) capture.toolOpened = true;
});

async function waitForClipboard() {
  while (capture && Date.now() < capture.until) {
    await new Promise((resolve) => setTimeout(resolve, 700));
    if (!capture) return;
    if (!capture.toolOpened && Date.now() - capture.startedAt > TOOL_OPEN_TIMEOUT_MS) {
      showStatus("캡처 도구가 열리지 않았습니다(회사 PC에서 막혀 있을 수 있음). Win+Shift+S로 캡처한 뒤 Ctrl+V를 누르거나, '🖥 창 선택 캡처'를 사용하세요.", true);
      capture.clipboardBlocked = true;
      return;
    }
    if (!capture.toolOpened || !document.hasFocus()) continue; // 클립보드는 이 화면이 앞에 있을 때만 읽힌다
    try {
      const blob = await clipboardImage();
      const signature = await blobSignature(blob);
      if (blob && signature !== lastImportedSignature) {
        lastImportedSignature = signature;
        addCapture(capture.kind, blob);
        return;
      }
    } catch (error) {
      if (error && error.name === "NotAllowedError") {
        capture.clipboardBlocked = true;
        showStatus("클립보드 읽기가 허용되지 않았습니다. 캡처한 뒤 이 화면에서 Ctrl+V를 누르세요.", true);
        return;
      }
    }
  }
}

// ── 창 선택 캡처(브라우저 화면 공유 기능) → 영역 자르기 ──
async function startWindowCapture(kind) {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getDisplayMedia) {
    showStatus("이 브라우저는 창 캡처를 지원하지 않습니다. 크롬·엣지 최신 버전에서 웹 주소로 열어 주세요.", true);
    return;
  }
  let stream;
  try {
    stream = await navigator.mediaDevices.getDisplayMedia({ video: { cursor: "never" }, audio: false });
  } catch (error) {
    showStatus("창 선택을 취소했거나 브라우저가 허용하지 않았습니다.", true);
    return;
  }
  try {
    const video = document.createElement("video");
    video.muted = true;
    video.srcObject = stream;
    await video.play();
    await new Promise((resolve) => setTimeout(resolve, 300)); // 첫 화면이 그려질 때까지 잠시 기다린다
    const canvas = document.createElement("canvas");
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    canvas.getContext("2d").drawImage(video, 0, 0);
    openCropper(kind, canvas);
  } finally {
    stream.getTracks().forEach((track) => track.stop());
  }
}

function openCropper(kind, source) {
  const overlay = document.createElement("div");
  overlay.className = "cropper";
  const box = document.createElement("div");
  box.className = "cropper-box";
  const guide = cell("p", "작업허가서 부분을 마우스로 드래그해 고르세요. 고르지 않으면 화면 전체를 넣습니다.", "cropper-guide");
  const stage = document.createElement("div");
  stage.className = "cropper-stage";
  const view = document.createElement("canvas");
  const maxWidth = Math.min(window.innerWidth - 80, 1400);
  const maxHeight = window.innerHeight - 200;
  const ratio = Math.min(1, maxWidth / source.width, maxHeight / source.height);
  view.width = Math.round(source.width * ratio);
  view.height = Math.round(source.height * ratio);
  view.getContext("2d").drawImage(source, 0, 0, view.width, view.height);
  const marquee = document.createElement("div");
  marquee.className = "cropper-marquee hidden";
  stage.append(view, marquee);

  let start = null;
  let rect = null;
  const point = (event) => {
    const bounds = view.getBoundingClientRect();
    return {
      x: Math.max(0, Math.min(view.width, event.clientX - bounds.left)),
      y: Math.max(0, Math.min(view.height, event.clientY - bounds.top)),
    };
  };
  view.addEventListener("pointerdown", (event) => {
    start = point(event);
    view.setPointerCapture(event.pointerId);
  });
  view.addEventListener("pointermove", (event) => {
    if (!start) return;
    const p = point(event);
    rect = { x: Math.min(start.x, p.x), y: Math.min(start.y, p.y), w: Math.abs(p.x - start.x), h: Math.abs(p.y - start.y) };
    Object.assign(marquee.style, { left: `${rect.x}px`, top: `${rect.y}px`, width: `${rect.w}px`, height: `${rect.h}px` });
    marquee.classList.remove("hidden");
  });
  view.addEventListener("pointerup", () => { start = null; });

  const close = () => overlay.remove();
  const use = cell("button", "넣기", "primary");
  use.type = "button";
  use.addEventListener("click", () => {
    const area = rect && rect.w > 10 && rect.h > 10
      ? { x: rect.x / ratio, y: rect.y / ratio, w: rect.w / ratio, h: rect.h / ratio }
      : { x: 0, y: 0, w: source.width, h: source.height };
    const out = document.createElement("canvas");
    out.width = Math.round(area.w);
    out.height = Math.round(area.h);
    out.getContext("2d").drawImage(source, area.x, area.y, area.w, area.h, 0, 0, out.width, out.height);
    out.toBlob((blob) => {
      close();
      addCapture(kind, blob);
    }, "image/png");
  });
  const cancel = cell("button", "취소");
  cancel.type = "button";
  cancel.addEventListener("click", close);
  const actions = document.createElement("div");
  actions.className = "actions";
  actions.append(use, cancel);
  box.append(guide, stage, actions);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
}

// Ctrl+V: 캡처 중이면 그 칸에, 아니면 작업허가서 칸에 붙여 넣는다
document.addEventListener("paste", (event) => {
  const image = [...(event.clipboardData ? event.clipboardData.files : [])].find((f) => f.type.startsWith("image/"));
  if (!image) return;
  event.preventDefault();
  addCapture(capture ? capture.kind : "permit", image);
});

async function runReview() {
  reviewButton.disabled = true;
  showStatus("서류를 읽는 중입니다…");
  const docs = {};
  const problems = [];
  for (const [kind, list] of Object.entries(files)) {
    if (!list.length) continue;
    const parts = [];
    for (const [index, file] of list.entries()) {
      try {
        const ocr = OCR_SLOTS.includes(kind) && IMAGE_NAME.test(file.name);
        const part = await readDocument(file, {
          ocr,
          onProgress: (message) => {
            if (message.status === "recognizing text") {
              showStatus(`${DOC_KINDS[kind]} 그림 글자 인식 중 (${index + 1}/${list.length}) ${Math.round(message.progress * 100)}%`);
            } else if (/loading|initializ/.test(message.status)) {
              showStatus("글자 인식 준비 중입니다(처음 한 번은 몇 초 걸립니다)…");
            }
          },
        });
        // 허가서 캡처는 글자가 있어야 한다(계획서·기타의 그림은 도면·사진이라 글자가 없을 수 있음)
        if (ocr && kind === "permit" && part.lines.length <= 1) problems.push(`${file.name}: 그림에서 글자를 읽지 못했습니다.`);
        // 첫 줄은 파일 이름이므로 글자가 그것뿐이면 스캔 이미지로 본다(사진 파일 제외)
        if (part.lines.length <= 1 && !IMAGE_NAME.test(file.name)) {
          problems.push(`${file.name}: 글자가 없습니다(스캔 이미지일 수 있음).`);
        }
        parts.push(part);
      } catch (error) {
        problems.push(`${file.name}: ${error.message || "읽지 못했습니다."}`);
      }
    }
    if (parts.length) docs[kind] = analyzeDocument(kind, parts);
  }
  lastReview = { ...reviewDocuments(docs), docs, reviewedAt: new Date() };
  renderResult(lastReview);
  if (problems.length) showStatus(problems.join(" / "), true);
  else statusEl.classList.add("hidden");
  reviewButton.disabled = false;
  // 자동 시험용(주소 끝이 'report'일 때만): 결과를 시험 서버로 보낸다
  if (/report$/.test(location.hash)) {
    fetch("/__report", {
      method: "POST",
      body: JSON.stringify({
        problems,
        results: lastReview.results,
        summary: lastReview.summary,
        extracted: Object.fromEntries(Object.entries(docs).map(([kind, doc]) => [kind, doc.files.map((f) => ({ file: f.fileName, lines: f.lines, names: f.names, dates: f.dates }))])),
      }),
    });
  }
}

function cell(tag, text, className) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text !== undefined) el.textContent = text;
  return el;
}

function renderResult({ results, summary, docs }) {
  resultPanel.classList.remove("hidden");

  const counts = { ok: 0, warn: 0, fail: 0, manual: 0 };
  results.forEach((r) => {
    if (r.status in counts) counts[r.status] += 1;
  });
  const summaryEl = document.getElementById("summary");
  summaryEl.innerHTML = "";
  [["fail", "미흡(반려 사유)"], ["warn", "확인 필요"], ["manual", "직접 확인"], ["ok", "적합"]].forEach(([status, label]) => {
    const tile = cell("div", undefined, `tile ${status}`);
    tile.append(cell("div", String(counts[status]), "num"), cell("div", label, "label"));
    summaryEl.appendChild(tile);
  });

  const info = document.getElementById("workInfo");
  info.innerHTML = "";
  const dl = document.createElement("dl");
  dl.className = "work-info";
  [
    ["작업 종류", summary.workTypes.join(", ") || "알 수 없음"],
    ["작업계획서 작성 대상", summary.planTargets.join(", ") || "해당 없음"],
    ["특별교육 대상", summary.educationTypes.join(" / ") || "해당 없음"],
    ["작업일", summary.workDates.length ? summary.workDates.join(", ") : "알 수 없음"],
    ["작업자", summary.workers.length ? `${summary.workers.length}명 (${summary.workers.join(", ")})` : "명단을 읽지 못함"],
  ].forEach(([term, value]) => dl.append(cell("dt", term), cell("dd", value)));
  info.appendChild(dl);

  const sectionsEl = document.getElementById("resultSections");
  sectionsEl.innerHTML = "";
  SECTION_ORDER.forEach((section) => {
    const rows = results.filter((r) => r.section === section);
    if (!rows.length) return;
    sectionsEl.appendChild(cell("h3", section, "section-title"));
    const table = cell("table", undefined, "result");
    table.innerHTML = '<colgroup><col class="c-status"><col class="c-title"><col></colgroup>'
      + "<thead><tr><th>결과</th><th>항목</th><th>근거·확인 내용</th></tr></thead>";
    const tbody = document.createElement("tbody");
    // 미흡 → 확인 필요 → 직접 확인 → 적합 → 참고 순
    const rank = { fail: 0, warn: 1, manual: 2, ok: 3, info: 4 };
    [...rows].sort((a, b) => rank[a.status] - rank[b.status]).forEach((r) => {
      const tr = document.createElement("tr");
      const statusTd = document.createElement("td");
      statusTd.appendChild(cell("span", STATUS_LABEL[r.status], `badge ${r.status}`));
      const titleTd = cell("td", r.title);
      if (r.doc && section !== "제출 서류") titleTd.appendChild(cell("span", r.doc, "basis"));
      const detailTd = cell("td", r.detail);
      if (r.basis) detailTd.appendChild(cell("span", r.basis, "basis"));
      tr.append(statusTd, titleTd, detailTd);
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    sectionsEl.appendChild(table);
  });

  const extracted = document.getElementById("extracted");
  extracted.innerHTML = "";
  const table = document.createElement("table");
  table.innerHTML = "<thead><tr><th>서류</th><th>파일</th><th>작업 종류</th><th>날짜</th><th>이름</th><th>교육시간</th><th>사진</th></tr></thead>";
  const tbody = document.createElement("tbody");
  Object.entries(DOC_KINDS).forEach(([kind, label]) => {
    const doc = docs[kind];
    if (!doc) return;
    doc.files.forEach((file) => {
      const tr = document.createElement("tr");
      [label, file.fileName, file.workTypes.map((t) => t.name).join(", ") || "-", file.dates.join(", ") || "-",
        file.names.join(", ") || "-", kind === "education" && file.hours !== null ? `${file.hours}시간` : "-", String(file.images)]
        .forEach((value) => tr.appendChild(cell("td", value)));
      tbody.appendChild(tr);
    });
  });
  table.appendChild(tbody);
  extracted.appendChild(table);
  resultPanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

function exportExcel() {
  if (!lastReview) return;
  const { results, summary, reviewedAt } = lastReview;
  const stamp = `${reviewedAt.getFullYear()}${String(reviewedAt.getMonth() + 1).padStart(2, "0")}${String(reviewedAt.getDate()).padStart(2, "0")}`;
  const head = [
    ["작업허가 서류 검토 결과"],
    ["검토일시", reviewedAt.toLocaleString("ko-KR")],
    ["작업 종류", summary.workTypes.join(", ")],
    ["작업계획서 작성 대상", summary.planTargets.join(", ")],
    ["특별교육 대상", summary.educationTypes.join(" / ")],
    ["작업일", summary.workDates.join(", ")],
    ["작업자", summary.workers.join(", ")],
    [],
    ["구분", "결과", "항목", "대상 서류", "확인 내용", "근거"],
  ];
  const rows = SECTION_ORDER.flatMap((section) =>
    results.filter((r) => r.section === section)
      .map((r) => [section, STATUS_LABEL[r.status], r.title, r.doc, r.detail, r.basis])
  );
  const sheet = XLSX.utils.aoa_to_sheet([...head, ...rows]);
  sheet["!cols"] = [{ wch: 14 }, { wch: 10 }, { wch: 36 }, { wch: 22 }, { wch: 70 }, { wch: 50 }];
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "검토결과");
  XLSX.writeFile(book, `작업허가_서류검토_${stamp}.xlsx`);
}

// 시험용 두 번째 장면(전기+화기): 주소 끝에 #demo2
const SAMPLE_FILES_2 = {
  permit: ["samples/scenario2/작업허가서.xlsx"],
  plan: ["samples/scenario2/작업계획서.pdf"],
  risk: ["samples/위험성평가_샘플.xlsx"],
  education: ["samples/scenario2/특별교육_전기화기.pdf"],
  attachments: ["samples/scenario2/화기감시자_지정서.pdf"],
};

// 시험용 세 번째 장면(장면 1을 Word·PowerPoint로): 주소 끝에 #demo3
const SAMPLE_FILES_3 = {
  ...SAMPLE_FILES,
  plan: ["samples/office/작업계획서.docx"],
  education: ["samples/office/특별안전교육_밀폐공간.pptx"],
};

// 시험용 네 번째 장면(작업계획서·특별교육을 압축으로, 한글 이름은 Windows 방식 CP949): 주소 끝에 #demo4
const SAMPLE_FILES_4 = {
  ...SAMPLE_FILES,
  plan: ["samples/zip/작업계획서_묶음.zip"],
  education: ["samples/zip/특별교육_묶음.zip"],
};

async function loadSamples(sampleFiles = SAMPLE_FILES) {
  showStatus("샘플 서류를 불러오는 중입니다…");
  try {
    for (const [kind, paths] of Object.entries(sampleFiles)) {
      files[kind] = [];
      for (const path of paths) {
        const resp = await fetch(path);
        if (!resp.ok) throw new Error(path);
        const blob = await resp.blob();
        await addFiles(kind, [new File([blob], path.split("/").pop(), { type: blob.type })]);
      }
    }
    renderSlots();
    await runReview();
  } catch (error) {
    showStatus("샘플을 불러오지 못했습니다. 인터넷 주소(웹)로 열었을 때 사용할 수 있습니다.", true);
  }
}

reviewButton.addEventListener("click", runReview);
document.getElementById("sampleButton").addEventListener("click", () => loadSamples());
document.getElementById("resetButton").addEventListener("click", () => {
  Object.keys(files).forEach((kind) => { files[kind] = []; });
  lastReview = null;
  resultPanel.classList.add("hidden");
  statusEl.classList.add("hidden");
  renderSlots();
});
document.getElementById("excelButton").addEventListener("click", exportExcel);
document.getElementById("printButton").addEventListener("click", () => window.print());

renderSlots();
if (/^#demo(report)?$/.test(location.hash)) loadSamples();
if (/^#demo2(report)?$/.test(location.hash)) loadSamples(SAMPLE_FILES_2);
if (/^#demo3(report)?$/.test(location.hash)) loadSamples(SAMPLE_FILES_3);
if (/^#demo4(report)?$/.test(location.hash)) loadSamples(SAMPLE_FILES_4);
