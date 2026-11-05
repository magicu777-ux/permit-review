// 화면: 서류 칸(드래그·선택) → 검토 → 결과 표·엑셀 저장
/* global XLSX, DOC_KINDS, readDocument, analyzeDocument, reviewDocuments, mergeCapturedParts, unzipFile, detectDrm */
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
const SECTION_ORDER = ["제출 서류", "작업허가서(Safe_B)", "작업 정보", "작업계획서", "화기 작업", "위험성평가", "특별안전교육", "서류 간 일치"];

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

    // 브라우저가 창을 직접 찍고 필요한 영역을 잘라 넣는다(작업허가서는 Safe_B 화면, 그 밖의 칸은
    // 문서보안(DRM)으로 파일을 못 읽을 때 문서를 열어 놓고 쪽마다 찍어 넣는다). 찍은 그림은 글자 인식으로 읽는다.
    const windowCaptureButton = document.createElement("button");
    windowCaptureButton.type = "button";
    windowCaptureButton.className = kind === "permit" ? "capture" : "capture secondary";
    windowCaptureButton.textContent = kind === "permit" ? "🖥 Safe_B 창 선택 캡처" : "🖥 창 선택 캡처";
    windowCaptureButton.addEventListener("click", () => startWindowCapture(kind));

    const list = document.createElement("ul");
    list.className = "file-list";
    const shownGroups = new Set();
    files[kind].forEach((file, index) => {
      const item = document.createElement("li");
      if (file.captureGroup) {
        // 한 번에 찍은 여러 쪽은 문서 하나로 보여 주고, ✕는 묶음 전체를 뺀다
        if (shownGroups.has(file.captureGroup)) return;
        shownGroups.add(file.captureGroup);
        const group = files[kind].filter((f) => f.captureGroup === file.captureGroup);
        const name = cell("span", ` 📷 ${file.captureName} — 캡처 ${group.length}쪽`);
        const thumb = document.createElement("img");
        thumb.className = "thumb";
        thumb.alt = "";
        thumb.src = URL.createObjectURL(file);
        thumb.title = "첫 쪽 크게 보기";
        thumb.addEventListener("click", () => window.open(thumb.src, "_blank"));
        name.prepend(thumb);
        const remove = cell("button", "✕");
        remove.type = "button";
        remove.title = "이 캡처 문서 빼기";
        remove.addEventListener("click", () => {
          files[kind] = files[kind].filter((f) => f.captureGroup !== file.captureGroup);
          renderSlots();
        });
        item.append(name, remove);
        list.appendChild(item);
        return;
      }
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
      if (file.drm) {
        // 암호화 파일: 이 PC에서 열어 놓은 문서 창을 찍어 이 파일 대신 넣는다
        item.classList.add("locked");
        name.prepend("🔒 ");
        const captureThis = cell("button", "🖥 열어서 캡처", "capture-this");
        captureThis.type = "button";
        captureThis.title = "이 파일을 더블클릭해 연 다음 누르세요";
        captureThis.addEventListener("click", () => startWindowCapture(kind, file));
        item.append(name, captureThis, remove);
      } else {
        item.append(name, remove);
      }
      list.appendChild(item);
    });

    slot.append(...[title, windowCaptureButton, pick, input, list].filter(Boolean));
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
  const excluded = [];
  // 검토 대상이 아닌 파일(Swing 작업계획서 등)은 넣지 않고 알려 준다
  const exclusion = (file) => (window.PERMIT_RULES.excludeFiles || [])
    .find((rule) => new RegExp(rule.pattern, "i").test(file.zipPath || file.name));
  const take = (file) => {
    const rule = exclusion(file);
    if (rule) excluded.push(`${file.name}(${rule.reason})`);
    else if (acceptable(kind, file)) added.push(file);
    else skipped.push(file.name);
  };
  for (const file of incoming) {
    if (/\.zip$/i.test(file.name)) {
      showStatus(`${file.name} 압축을 푸는 중입니다…`);
      try {
        const unpacked = await unzipFile(file);
        unpacked.files.forEach(take);
        skipped.push(...unpacked.skipped);
        if (!unpacked.files.length) skipped.push(`${file.name}(빈 압축)`);
      } catch (error) {
        skipped.push(`${file.name}(${error.message || "압축을 풀지 못함"})`);
      }
    } else {
      take(file);
    }
  }
  // 문서보안(DRM) 파일은 넣을 때 바로 알아 둔다 → 목록에 🔒와 '열어서 캡처' 버튼을 보여 준다
  for (const file of added) {
    if (IMAGE_NAME.test(file.name)) continue;
    try {
      file.drm = detectDrm(file.name, await file.slice(0, 1024).arrayBuffer());
    } catch (error) {
      file.drm = null;
    }
  }
  files[kind].push(...added);
  renderSlots();
  const locked = added.filter((f) => f.drm);
  const notes = [];
  if (locked.length) {
    notes.push(`🔒 암호화(문서보안) 파일 ${locked.length}개는 내용을 읽을 수 없습니다. 이 PC에서 압축을 풀고 파일을 더블클릭해 연 다음, `
      + `목록의 '🖥 열어서 캡처'를 눌러 쪽마다 찍어 넣으세요(찍은 문서가 그 파일을 대신합니다)`);
  }
  if (excluded.length) notes.push(`검토 대상이 아니어서 뺐습니다: ${excluded.join(", ")}`);
  if (skipped.length) notes.push(`읽을 수 없는 파일은 건너뛰었습니다: ${skipped.join(", ")} — PDF, 엑셀, Word(docx), PowerPoint(pptx), 사진, zip만 넣을 수 있습니다.`);
  if (notes.length) showStatus(notes.join(" / "), Boolean(skipped.length || locked.length));
  else if (incoming.some((f) => /\.zip$/i.test(f.name))) showStatus(`압축을 풀어 ${added.length}개 파일을 ${DOC_KINDS[kind]} 칸에 넣었습니다.`);
}

// ── 캡처한 그림 넣기 ──
function addCapture(kind, blob, group) {
  const now = new Date();
  const stamp = [now.getHours(), now.getMinutes(), now.getSeconds()].map((n) => String(n).padStart(2, "0")).join("");
  const extension = (blob.type.split("/")[1] || "png").replace("jpeg", "jpg");
  const prefix = kind === "permit" ? "Safe_B_작업허가서" : DOC_KINDS[kind];
  const file = new File([blob], `${prefix}_캡처_${stamp}.${extension}`, { type: blob.type });
  file.captured = true; // 찍은 화면은 어느 칸이든 글자 인식으로 읽는다
  if (group) Object.assign(file, group); // 한 번의 캡처로 찍은 여러 쪽은 문서 하나(captureGroup)로 묶는다
  files[kind].push(file);
  renderSlots();
  showStatus(`캡처한 화면을 ${DOC_KINDS[kind]} 칸에 넣었습니다. 더 캡처하거나 '검토하기'를 누르세요.`);
}

// ── 창 선택 캡처(브라우저 화면 공유 기능) ──
// 문서보안(DRM) 파일은 이 PC에서 더블클릭해 열면 보이므로, 열린 창을 찍어 글자 인식으로 읽는다.
// 창과 잘라낼 영역은 한 번만 고르고, 쪽을 넘기면(화면이 바뀌면) 자동으로 찍는다. 찍은 쪽들은 문서 하나로 묶는다.
const AUTO_CAPTURE_MS = 800;
let captureSession = null;

function grabFrame(video, area) {
  const canvas = document.createElement("canvas");
  const a = area || { x: 0, y: 0, w: video.videoWidth, h: video.videoHeight };
  canvas.width = Math.round(a.w);
  canvas.height = Math.round(a.h);
  canvas.getContext("2d").drawImage(video, a.x, a.y, a.w, a.h, 0, 0, canvas.width, canvas.height);
  return canvas;
}

// 화면이 바뀌었는지 보기 위한 작은 흑백 지문(48×27)과 두 지문의 차이(0~255)
function frameSignature(canvas) {
  const small = document.createElement("canvas");
  small.width = 48;
  small.height = 27;
  const context = small.getContext("2d");
  context.drawImage(canvas, 0, 0, 48, 27);
  const px = context.getImageData(0, 0, 48, 27).data;
  const sig = new Float32Array(48 * 27);
  for (let i = 0; i < sig.length; i += 1) sig[i] = (px[i * 4] + px[i * 4 + 1] + px[i * 4 + 2]) / 3;
  return sig;
}

function signatureDiff(a, b) {
  let sum = 0;
  for (let i = 0; i < a.length; i += 1) sum += Math.abs(a[i] - b[i]);
  return sum / a.length;
}

// target: '열어서 캡처'를 누른 암호화 파일(있으면 그 파일 이름으로 찍고, 끝내면 그 파일을 대신한다)
async function startWindowCapture(kind, target) {
  if (captureSession) {
    showStatus("이미 캡처 중입니다. 화면 아래 막대의 '끝내기'를 누른 뒤 다시 시작하세요.", true);
    return;
  }
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
  const video = document.createElement("video");
  video.muted = true;
  video.srcObject = stream;
  await video.play();
  await new Promise((resolve) => setTimeout(resolve, 300)); // 첫 화면이 그려질 때까지 잠시 기다린다
  const stop = () => stream.getTracks().forEach((track) => track.stop());
  const guide = kind === "permit"
    ? "작업허가서 부분을 마우스로 드래그해 고르세요. 고르지 않으면 화면 전체를 넣습니다."
    : "문서 내용(슬라이드·쪽) 부분을 드래그해 고르세요. 이 영역을 다음 쪽에도 그대로 씁니다. 고르지 않으면 화면 전체.";
  openCropper(grabFrame(video), guide, (area) => {
    if (!area) {
      stop();
      return;
    }
    beginCaptureSession(kind, stream, video, area, target);
  });
}

function beginCaptureSession(kind, stream, video, area, target) {
  const groupCount = new Set(files[kind].filter((f) => f.captureGroup).map((f) => f.captureGroup)).size;
  const targetName = target ? (target.zipPath || target.name).split("/").pop().replace(/\.[^.]+$/, "") : "";
  const session = {
    kind, stream, video, area, pages: 0, lastSig: null, pendingSig: null,
    group: `capture-${Date.now()}`,
    name: targetName || (kind === "permit" ? "Safe_B 작업허가서" : `${DOC_KINDS[kind]} 캡처 ${groupCount + 1}`),
  };
  captureSession = session;

  const bar = document.createElement("div");
  bar.className = "capture-bar";
  const info = cell("span", "", "capture-info");
  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.value = session.name;
  nameInput.title = "문서 이름(예: 특별안전교육(전기), 작업계획서_크레인) — 서류 구분에 쓰입니다";
  nameInput.addEventListener("input", () => { session.name = nameInput.value.trim() || session.name; });
  const autoLabel = document.createElement("label");
  const autoBox = document.createElement("input");
  autoBox.type = "checkbox";
  autoBox.checked = true;
  autoLabel.append(autoBox, " 쪽 넘기면 자동으로 찍기");
  const shotButton = cell("button", "📸 지금 찍기");
  shotButton.type = "button";
  const endButton = cell("button", "끝내기", "primary");
  endButton.type = "button";
  bar.append(info, nameInput, autoLabel, shotButton, endButton);
  document.body.appendChild(bar);
  const updateInfo = () => {
    info.textContent = `📷 ${DOC_KINDS[kind]} 캡처 중 · ${session.pages}쪽 찍음 — 문서 창에서 쪽을 넘기세요(PageDown)`;
  };

  const shoot = (canvas, sig) => new Promise((resolve) => canvas.toBlob((blob) => {
    session.pages += 1;
    session.lastSig = sig;
    // 작업허가서는 한 장짜리 화면이라 묶지 않는다
    addCapture(kind, blob, kind === "permit" ? null
      : { captureGroup: session.group, capturePage: session.pages, captureName: session.name });
    updateInfo();
    resolve();
  }, "image/png"));

  const tick = async () => {
    if (captureSession !== session || !autoBox.checked) return;
    const canvas = grabFrame(video, area);
    const sig = frameSignature(canvas);
    if (session.lastSig && signatureDiff(sig, session.lastSig) < 4) {
      session.pendingSig = null; // 같은 쪽
      return;
    }
    // 화면이 바뀐 직후(넘김 효과 중)는 건너뛰고, 다음 확인에서도 그대로면 새 쪽으로 찍는다
    if (session.pendingSig && signatureDiff(sig, session.pendingSig) < 2) {
      session.pendingSig = null;
      await shoot(canvas, sig);
    } else {
      session.pendingSig = sig;
    }
  };
  const timer = setInterval(tick, AUTO_CAPTURE_MS);

  const end = () => {
    if (captureSession !== session) return;
    clearInterval(timer);
    stream.getTracks().forEach((track) => track.stop());
    bar.remove();
    captureSession = null;
    files[kind].filter((f) => f.captureGroup === session.group).forEach((f) => { f.captureName = session.name; });
    // 찍은 문서가 암호화 파일을 대신한다(파일 이름은 문서 이름으로 남음)
    if (target && session.pages) files[kind] = files[kind].filter((f) => f !== target);
    renderSlots();
    showStatus(`'${session.name}' ${session.pages}쪽을 ${DOC_KINDS[kind]} 칸에 넣었습니다. '검토하기'를 누르면 글자를 인식해 검토합니다.`);
  };
  shotButton.addEventListener("click", () => {
    const canvas = grabFrame(video, area);
    shoot(canvas, frameSignature(canvas));
  });
  endButton.addEventListener("click", end);
  stream.getVideoTracks()[0].addEventListener("ended", end); // 브라우저의 '공유 중지'를 눌러도 끝낸다

  updateInfo();
  const first = grabFrame(video, area);
  shoot(first, frameSignature(first)); // 처음 보이는 쪽은 바로 찍는다
  if (kind === "permit") setTimeout(end, 300); // 작업허가서는 한 장이면 충분
}

// onDone(area): 고른 영역(원본 화면 좌표) 또는 전체, 취소하면 null
function openCropper(source, guideText, onDone) {
  const overlay = document.createElement("div");
  overlay.className = "cropper";
  const box = document.createElement("div");
  box.className = "cropper-box";
  const guide = cell("p", guideText, "cropper-guide");
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
  const use = cell("button", "이 영역으로 시작", "primary");
  use.type = "button";
  use.addEventListener("click", () => {
    const area = rect && rect.w > 10 && rect.h > 10
      ? { x: rect.x / ratio, y: rect.y / ratio, w: rect.w / ratio, h: rect.h / ratio }
      : { x: 0, y: 0, w: source.width, h: source.height };
    close();
    onDone(area);
  });
  const cancel = cell("button", "취소");
  cancel.type = "button";
  cancel.addEventListener("click", () => {
    close();
    onDone(null);
  });
  const actions = document.createElement("div");
  actions.className = "actions";
  actions.append(use, cancel);
  box.append(guide, stage, actions);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
}

// Ctrl+V: 클립보드의 그림(Win+Shift+S로 찍은 것 등)을 작업허가서 칸에 붙여 넣는다
document.addEventListener("paste", (event) => {
  const image = [...(event.clipboardData ? event.clipboardData.files : [])].find((f) => f.type.startsWith("image/"));
  if (!image) return;
  event.preventDefault();
  addCapture("permit", image);
});

async function runReview() {
  reviewButton.disabled = true;
  showStatus("서류를 읽는 중입니다…");
  const docs = {};
  const problems = [];
  const lockedNames = [];
  for (const [kind, list] of Object.entries(files)) {
    if (!list.length) continue;
    const parts = [];
    for (const [index, file] of list.entries()) {
      try {
        const ocr = (OCR_SLOTS.includes(kind) || file.captured) && IMAGE_NAME.test(file.name);
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
        if (part.drm) {
          lockedNames.push(file.name);
        } else if (part.lines.length <= 1 && !IMAGE_NAME.test(file.name)) {
          // 첫 줄은 파일 이름이므로 글자가 그것뿐이면 스캔 이미지로 본다(사진 파일 제외)
          problems.push(`${file.name}: 글자가 없습니다(스캔 이미지일 수 있음).`);
        }
        if (file.captureGroup) {
          part.captureGroup = file.captureGroup;
          part.captureName = file.captureName;
        }
        parts.push(part);
      } catch (error) {
        problems.push(`${file.name}: ${error.message || "읽지 못했습니다."}`);
      }
    }
    // 창 선택 캡처로 찍은 쪽들은 찍은 순서대로 문서 하나로 합친다
    const merged = [];
    parts.forEach((part) => {
      if (!part.captureGroup) {
        merged.push(part);
        return;
      }
      if (merged.some((p) => p.captureGroup === part.captureGroup)) return;
      const group = parts.filter((p) => p.captureGroup === part.captureGroup);
      merged.push({ ...mergeCapturedParts(part.captureName, group), captureGroup: part.captureGroup });
    });
    if (merged.length) docs[kind] = analyzeDocument(kind, merged);
  }
  if (lockedNames.length) {
    // 암호화 파일은 이름만 확인했다 → 다음에 할 일(열어서 캡처)을 알려 준다
    problems.unshift(`🔒 암호화 파일 ${lockedNames.length}개는 파일 이름으로 제출 여부만 확인했습니다(내용 항목은 '직접 확인'). `
      + `내용까지 검토하려면 각 파일을 이 PC에서 더블클릭해 연 다음, 위 목록에서 그 파일의 '🖥 열어서 캡처'로 쪽마다 찍고 다시 '검토하기'를 누르세요. `
      + `(${lockedNames.join(", ")})`);
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
        slots: slotsEl.innerText,
        results: lastReview.results,
        summary: lastReview.summary,
        extracted: Object.fromEntries(Object.entries(docs).map(([kind, doc]) => [kind, doc.files.map((f) => ({
          file: f.fileName, lines: f.lines, names: f.names, dates: f.dates,
          options: (f.options || []).map((o) => `p${o.page} ${o.filled === null ? o.density : o.filled} @${o.textIndex} ${o.text.slice(0, 40)}`),
        }))])),
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
  permit: ["samples/scenario2/작업허가서_캡처.png"],
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
      if (!paths.length) continue;
      files[kind] = [];
      for (const path of paths) {
        const resp = await fetch(path);
        if (!resp.ok) throw new Error(path);
        const blob = await resp.blob();
        await addFiles(kind, [new File([blob], decodeURIComponent(path.split("/").pop()), { type: blob.type })]);
      }
    }
    renderSlots();
    await runReview();
  } catch (error) {
    showStatus("샘플을 불러오지 못했습니다. 인터넷 주소(웹)로 열었을 때 사용할 수 있습니다.", true);
  }
}

// 시험용 다섯 번째 장면(장면 2의 작업계획서를 '창 선택 캡처'로 쪽마다 찍은 것처럼): 주소 끝에 #demo5
// 실제 캡처는 화면 공유 허락이 필요해 자동 시험이 안 되므로, PDF 쪽을 그림으로 그려 캡처 묶음으로 넣는다.
async function loadCaptureDemo() {
  const { plan, ...rest } = SAMPLE_FILES_2;
  const resp = await fetch(plan[0]);
  const pdf = await window.pdfjsLib.getDocument({ data: await resp.arrayBuffer() }).promise;
  const group = { captureGroup: "capture-demo", captureName: "작업계획서(전기·화기) 캡처" };
  const pages = [];
  for (let n = 1; n <= pdf.numPages; n += 1) {
    const page = await pdf.getPage(n);
    const viewport = page.getViewport({ scale: 1.6 });
    const canvas = document.createElement("canvas");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const context = canvas.getContext("2d");
    context.fillStyle = "#fff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: context, viewport }).promise;
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
    pages.push(Object.assign(new File([blob], `작업계획서_캡처_${n}.png`, { type: "image/png" }), group, { captured: true, capturePage: n }));
  }
  files.plan = pages;
  await loadSamples(rest);
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
if (/^#demo5(report)?$/.test(location.hash)) loadCaptureDemo();
