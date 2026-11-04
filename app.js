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
const IMAGE_SLOTS = ["permit", "education", "attachments"];
const OCR_SLOTS = ["permit", "attachments"];
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
  return /\.(pdf|xlsx|xlsm|xls)$/i.test(file.name)
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
    input.accept = IMAGE_SLOTS.includes(kind) ? ".pdf,.xlsx,.xlsm,.xls,image/*" : ".pdf,.xlsx,.xlsm,.xls";
    input.id = `file-${kind}`;
    input.addEventListener("change", () => addFiles(kind, input.files));

    const pick = document.createElement("label");
    pick.className = "pick";
    pick.htmlFor = input.id;
    pick.textContent = "파일 선택 또는 끌어 놓기";

    // 작업허가서: Safe_b 화면을 캡처 도구로 찍어 바로 넣는다
    let captureButton = null;
    if (kind === "permit") {
      captureButton = document.createElement("button");
      captureButton.type = "button";
      captureButton.className = "capture";
      captureButton.textContent = "📷 Safe_b 화면 캡처";
      captureButton.addEventListener("click", () => startCapture(kind));
    }

    const list = document.createElement("ul");
    list.className = "file-list";
    files[kind].forEach((file, index) => {
      const item = document.createElement("li");
      const name = document.createElement("span");
      name.textContent = file.name;
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

    slot.append(...[title, captureButton, pick, input, list].filter(Boolean));
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

function addFiles(kind, fileList) {
  const added = [...fileList].filter((file) => acceptable(kind, file));
  if (added.length < fileList.length) {
    showStatus(IMAGE_SLOTS.includes(kind) ? "PDF, 엑셀(xlsx), 사진 파일만 넣을 수 있습니다." : "PDF 또는 엑셀(xlsx) 파일만 넣을 수 있습니다.", true);
  }
  files[kind].push(...added);
  renderSlots();
}

// ── Safe_b 화면 캡처 ──
// 버튼 → Windows 캡처 도구(ms-screenclip:) 실행 → 사용자가 영역을 드래그 → 클립보드에 들어온 그림을
// 이 화면으로 돌아왔을 때 자동으로 가져온다. 클립보드 읽기를 막았으면 Ctrl+V로 붙여 넣으면 된다.
const CAPTURE_WAIT_MS = 3 * 60 * 1000;
let capture = null;

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

async function startCapture(kind) {
  let baseline = null;
  let clipboardBlocked = !navigator.clipboard || !navigator.clipboard.read;
  if (!clipboardBlocked) {
    try {
      baseline = await blobSignature(await clipboardImage());
    } catch (error) {
      clipboardBlocked = error && error.name === "NotAllowedError";
    }
  }
  capture = { kind, baseline, until: Date.now() + CAPTURE_WAIT_MS, clipboardBlocked };
  showStatus(clipboardBlocked
    ? "캡처 도구에서 Safe_b의 작업허가서 영역을 드래그한 뒤, 이 화면에서 Ctrl+V를 누르세요."
    : "캡처 도구에서 Safe_b의 작업허가서 영역을 드래그하세요. 이 화면으로 돌아오면 자동으로 들어옵니다.");
  // Windows 10·11 캡처 도구(Win+Shift+S와 같음). 브라우저가 '열기'를 물으면 허용한다.
  window.location.href = "ms-screenclip:";
  if (!clipboardBlocked) waitForClipboard();
}

async function waitForClipboard() {
  while (capture && Date.now() < capture.until) {
    await new Promise((resolve) => setTimeout(resolve, 700));
    if (!capture || !document.hasFocus()) continue; // 클립보드는 이 화면이 앞에 있을 때만 읽힌다
    try {
      const blob = await clipboardImage();
      if (blob && (await blobSignature(blob)) !== capture.baseline) {
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
        if (ocr && part.lines.length <= 1) problems.push(`${file.name}: 그림에서 글자를 읽지 못했습니다.`);
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

async function loadSamples(sampleFiles = SAMPLE_FILES) {
  showStatus("샘플 서류를 불러오는 중입니다…");
  try {
    for (const [kind, paths] of Object.entries(sampleFiles)) {
      files[kind] = [];
      for (const path of paths) {
        const resp = await fetch(path);
        if (!resp.ok) throw new Error(path);
        const blob = await resp.blob();
        files[kind].push(new File([blob], path.split("/").pop()));
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
