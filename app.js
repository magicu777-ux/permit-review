// 화면: 서류 칸(드래그·선택) → 검토 → 결과 표·엑셀 저장
/* global XLSX, DOC_KINDS, readDocument, analyzeDocument, reviewDocuments */
// DOC_KINDS·readDocument 등은 review.js에서 전역으로 선언된다.

const SLOT_HINTS = {
  permit: "작업 종류·일자·작업자",
  plan: "작업방법·안전대책",
  risk: "위험요인·감소대책",
  education: "교육시간·이수자 명단",
};
const STATUS_LABEL = { ok: "적합", warn: "확인 필요", fail: "미흡", info: "참고" };
const SECTION_ORDER = ["제출 서류", "작업 정보", "작업계획서", "위험성평가", "특별안전교육", "서류 간 일치"];

// 샘플(가상 서류) — samples 폴더
const SAMPLE_FILES = {
  permit: ["samples/작업허가서_샘플.xlsx"],
  plan: ["samples/작업계획서_샘플.pdf"],
  risk: ["samples/위험성평가_샘플.xlsx"],
  education: ["samples/특별안전교육_샘플.pdf"],
};

const files = { permit: [], plan: [], risk: [], education: [] };
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

function acceptable(file) {
  return /\.(pdf|xlsx|xlsm|xls)$/i.test(file.name);
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
    input.accept = ".pdf,.xlsx,.xlsm,.xls";
    input.id = `file-${kind}`;
    input.addEventListener("change", () => addFiles(kind, input.files));

    const pick = document.createElement("label");
    pick.className = "pick";
    pick.htmlFor = input.id;
    pick.textContent = "파일 선택 또는 끌어 놓기";

    const list = document.createElement("ul");
    list.className = "file-list";
    files[kind].forEach((file, index) => {
      const item = document.createElement("li");
      const name = document.createElement("span");
      name.textContent = file.name;
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

    slot.append(title, pick, input, list);
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
  const added = [...fileList].filter(acceptable);
  if (added.length < fileList.length) {
    showStatus("PDF 또는 엑셀(xlsx) 파일만 넣을 수 있습니다.", true);
  }
  files[kind].push(...added);
  renderSlots();
}

async function runReview() {
  reviewButton.disabled = true;
  showStatus("서류를 읽는 중입니다…");
  const docs = {};
  const problems = [];
  for (const [kind, list] of Object.entries(files)) {
    if (!list.length) continue;
    const parts = [];
    for (const file of list) {
      try {
        const part = await readDocument(file);
        if (!part.lines.length) problems.push(`${file.name}: 글자가 없습니다(스캔 이미지일 수 있음).`);
        parts.push(part);
      } catch (error) {
        problems.push(`${file.name}: ${error.message || "읽지 못했습니다."}`);
      }
    }
    if (parts.some((part) => part.lines.length)) docs[kind] = analyzeDocument(kind, parts);
  }
  lastReview = { ...reviewDocuments(docs), docs, reviewedAt: new Date() };
  renderResult(lastReview);
  if (problems.length) showStatus(problems.join(" / "), true);
  else statusEl.classList.add("hidden");
  reviewButton.disabled = false;
}

function cell(tag, text, className) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text !== undefined) el.textContent = text;
  return el;
}

function renderResult({ results, summary, docs }) {
  resultPanel.classList.remove("hidden");

  const counts = { ok: 0, warn: 0, fail: 0 };
  results.forEach((r) => {
    if (r.status in counts) counts[r.status] += 1;
  });
  const summaryEl = document.getElementById("summary");
  summaryEl.innerHTML = "";
  [["fail", "미흡"], ["warn", "확인 필요"], ["ok", "적합"]].forEach(([status, label]) => {
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
    // 미흡 → 확인 필요 → 적합 → 참고 순
    const rank = { fail: 0, warn: 1, ok: 2, info: 3 };
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
  table.innerHTML = "<thead><tr><th>서류</th><th>작업 종류</th><th>날짜</th><th>이름</th><th>교육시간</th><th>줄 수</th></tr></thead>";
  const tbody = document.createElement("tbody");
  Object.entries(DOC_KINDS).forEach(([kind, label]) => {
    const doc = docs[kind];
    if (!doc) return;
    const tr = document.createElement("tr");
    [label, doc.workTypes.map((t) => t.name).join(", ") || "-", doc.dates.join(", ") || "-",
      doc.names.join(", ") || "-", doc.hours === null ? "-" : `${doc.hours}시간`, String(doc.lines.length)]
      .forEach((value) => tr.appendChild(cell("td", value)));
    tbody.appendChild(tr);
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

async function loadSamples() {
  showStatus("샘플 서류를 불러오는 중입니다…");
  try {
    for (const [kind, paths] of Object.entries(SAMPLE_FILES)) {
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
document.getElementById("sampleButton").addEventListener("click", loadSamples);
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
if (location.hash === "#demo") loadSamples();
