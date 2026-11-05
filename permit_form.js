// =========================================================
// 작업허가서(Safe_B 화면) 칸 점검
// ---------------------------------------------------------
// - 필수 입력 칸이 채워졌는지(항목 이름 뒤 ~ 다음 항목 이름 앞의 글자)
// - High Risk·긴급작업 여부가 '아니오'인지(라디오 버튼 채움)
// - 작업구분에서 고른 작업에 맞는 서류(특별교육·작업계획서·하역운반기계 계획서 1쪽 체크)가 있는지, 반대로
//   서류에는 있는데 작업구분에서 빠진 작업은 없는지
// 선택 여부는 review.js의 options(선택 표시 + 글자, 채움 정도)로 판단한다.
// =========================================================

/* global hasAny, findLine, DOC_KINDS */

const PERMIT_SECTION = "작업허가서(Safe_B)";
const DENSITY_GAP = 0.15;

function compact(text) {
  return (text || "").replace(/\s+/g, "");
}

// 한 묶음(작업구분 칸들, 한 쌍의 예/아니오) 안에서 채움 정도를 비교해 선택 여부를 정한다
function decideFilled(options) {
  const unknown = options.filter((o) => o.filled === null || o.filled === undefined);
  if (!unknown.length) return;
  const values = unknown.map((o) => o.density);
  const max = Math.max(...values);
  const min = Math.min(...values);
  if (max - min >= DENSITY_GAP) {
    const threshold = (max + min) / 2;
    unknown.forEach((o) => { o.filled = o.density > threshold; });
  } else if (max < 0.1) {
    unknown.forEach((o) => { o.filled = false; });
  }
}

// 글자 인식이 조금 틀린 버튼 이름('자량형', '시저령')을 두 글자 묶음이 가장 많이 겹치는 작업구분으로 짝짓는다.
// 확실할 때만(겹침 2개 이상, 2등보다 많음) 돌려주고, 아니면 null.
function bigrams(text) {
  const t = compact(text).replace(/[^가-힣A-Za-z0-9.]/g, "");
  const grams = new Set();
  for (let i = 0; i + 1 < t.length; i += 1) grams.add(t.slice(i, i + 2));
  return grams;
}

function closestCategory(text, categories) {
  const grams = bigrams(text);
  const scores = categories.map((c) => [...bigrams(c.name)].filter((g) => grams.has(g)).length);
  const best = Math.max(...scores);
  if (best < 2 || scores.filter((s) => s === best).length > 1) return null;
  return scores.indexOf(best);
}

function labelMatches(line, labels) {
  const found = [];
  labels.forEach((label) => {
    for (const m of line.matchAll(new RegExp(label.label, "g"))) {
      found.push({ name: label.name, start: m.index, end: m.index + m[0].length });
    }
  });
  // 겹치면 긴 쪽만(예: '작업장소 구분' 안의 '구분')
  return found
    .sort((a, b) => a.start - b.start || b.end - a.end)
    .filter((m, i, all) => !all.some((o, j) => j !== i && o.start <= m.start && o.end >= m.end && (o.end - o.start) > (m.end - m.start)));
}

function cleanValue(value) {
  return value
    .replace(/[|ㅣ_:：'"」「.,·\-~*\s]+/g, " ")
    // 입력 칸의 검색(🔍 → 'Q')·펼침(▾ → 'v') 아이콘으로 읽힌 한 글자는 값이 아니다(빈 칸도 아이콘은 보인다)
    .replace(/(^|\s)[Qqv∨⌄](?=\s|$)/g, " ")
    .trim();
}

// 칸마다 어떤 항목인지 정한다: 이름이 읽힌 칸은 이름으로, 못 읽은 칸은 앞뒤로 이름이 읽힌 칸 사이의
// 빈 개수가 Safe_B 화면 순서(order)와 맞을 때 그 순서로 채운다.
function labelFormCells(cells, labels, order) {
  const named = cells.map((cell) => {
    const text = cell.label.replace(/\*/g, " ");
    const hit = labels.find((l) => new RegExp(l.label).test(text) || new RegExp(l.label).test(text.replace(/\s/g, "")));
    return { ...cell, name: hit ? hit.name : null };
  });
  const indexOf = (name) => order.indexOf(name);
  const anchors = named.map((c, i) => [i, indexOf(c.name)]).filter(([, o]) => o >= 0);
  for (let k = 0; k + 1 < anchors.length; k += 1) {
    const [i1, o1] = anchors[k];
    const [i2, o2] = anchors[k + 1];
    if (i2 - i1 === o2 - o1) for (let i = i1 + 1; i < i2; i += 1) if (!named[i].name) named[i].name = order[o1 + (i - i1)];
  }
  // 맨 앞·맨 뒤 칸: 이웃한 이름 칸 바로 앞·뒤면 순서로 채운다
  if (anchors.length) {
    const [firstIndex, firstOrder] = anchors[0];
    for (let i = firstIndex - 1, o = firstOrder - 1; i >= 0 && o >= 0 && !named[i].name; i -= 1, o -= 1) {
      if (/신청자|관리\s*감독자|안전\s*관리자|관리\s*책임자/.test(named[i].label)) break;
      named[i].name = order[o];
    }
    const [lastIndex, lastOrder] = anchors[anchors.length - 1];
    for (let i = lastIndex + 1, o = lastOrder + 1; i < named.length && o < order.length && !named[i].name; i += 1, o += 1) {
      named[i].name = order[o];
    }
  }
  return named;
}

function readPermitForm(permit) {
  const rules = window.PERMIT_RULES.permitForm;
  const allLabels = [...rules.fields, ...rules.mustBeNo, { name: "작업구분", label: rules.categoryLabel }];
  const lines = permit.lines.filter((line) => !line.startsWith("[파일]"));
  const form = { fields: {}, yesNo: {}, categories: [], categoriesDecided: false, categoryLines: new Set() };

  // 1) 필수 입력 칸
  //    Safe_B 화면은 항목 이름 칸·값 칸을 따로 읽은 결과(formCells)를 먼저 쓴다(더 정확함). 버튼으로 고르는 칸은 버튼 색으로 판단.
  const cells = labelFormCells(permit.formCells || [], allLabels.concat(rules.otherLabels), rules.cellOrder);
  rules.fields.forEach((field) => {
    const cell = field.choice ? null : cells.find((c) => c.name === field.name);
    if (cell) {
      const value = cleanValue(cell.value);
      const meaningful = value.replace(/[^가-힣A-Za-z0-9]/g, "");
      form.fields[field.name] = { found: true, value, empty: !meaningful || rules.placeholders.includes(compact(value)) };
      return;
    }
    for (let i = 0; i < lines.length; i += 1) {
      const matches = labelMatches(lines[i], allLabels);
      const index = matches.findIndex((m) => m.name === field.name);
      if (index < 0) continue;
      const end = matches[index].end;
      const next = index + 1 < matches.length ? matches[index + 1].start : lines[i].length;
      let value = cleanValue(lines[i].slice(end, next));
      // 항목 이름 아래 줄에 값이 있는 화면
      if (!value && i + 1 < lines.length && !labelMatches(lines[i + 1], allLabels).length) value = cleanValue(lines[i + 1]);
      const meaningful = value.replace(/[^가-힣A-Za-z0-9]/g, "");
      let empty = !meaningful || rules.placeholders.includes(compact(value));
      // 버튼으로 고르는 칸(작업장소 구분): Safe_B 버튼이 읽혔으면 파란(선택된) 버튼이 있어야 한다
      if (field.choice) {
        const chips = (permit.options || []).filter((o) => o.chip && o.lineText === lines[i] && o.textIndex >= end && o.textIndex < next);
        if (chips.length) {
          const chosen = chips.filter((o) => o.filled);
          empty = !chosen.length;
          value = chosen.length ? `선택됨(${chosen.length}개)` : "선택된 버튼 없음";
        }
      }
      form.fields[field.name] = { found: true, value, empty };
      return;
    }
    form.fields[field.name] = { found: false, value: "", empty: true };
  });

  // 2) 예/아니오
  rules.mustBeNo.forEach((item) => {
    const lineIndex = lines.findIndex((line) => labelMatches(line, allLabels).some((m) => m.name === item.name));
    if (lineIndex < 0) {
      form.yesNo[item.name] = { found: false, answer: null };
      return;
    }
    const line = lines[lineIndex];
    const matches = labelMatches(line, allLabels);
    const index = matches.findIndex((m) => m.name === item.name);
    const start = matches[index].end;
    const end = index + 1 < matches.length ? matches[index + 1].start : line.length;
    const pair = (permit.options || []).filter((o) => o.lineText === line && o.textIndex >= start && o.textIndex < end)
      .sort((a, b) => (a.x ?? a.textIndex) - (b.x ?? b.textIndex)).slice(0, 2);
    let answer = null;
    if (pair.length === 2 && pair.every((o) => o.chip)) {
      // Safe_B 버튼: 둘 중 더 파란 쪽(차이가 분명할 때)
      const [a, b] = pair;
      pair.forEach((o) => { o.filled = false; });
      if (Math.abs(a.blue - b.blue) >= 0.3) (a.blue > b.blue ? a : b).filled = true;
    }
    if (pair.length === 2) {
      decideFilled(pair);
      // 글자가 읽혔으면 글자로, 아니면 화면 순서(예 → 아니오)로 어느 쪽인지 정한다
      const isNo = (o, i) => (/아니|no/i.test(o.text) ? true : /^예$|yes/i.test(compact(o.text)) ? false : i === 1);
      const chosen = pair.filter((o) => o.filled);
      if (chosen.length === 1) answer = isNo(chosen[0], pair.indexOf(chosen[0])) ? "아니오" : "예";
    }
    if (!answer) {
      // 선택 표시가 없고 값이 글자로만 있는 화면(예: '아니오')
      const value = compact(line.slice(start, end));
      if (/아니[오요]|^N$|^No$/i.test(value) && !/(^|[^니])예/.test(value.replace(/아니[오요]/g, ""))) answer = "아니오";
      else if (/^예$|^Y$|^Yes$/i.test(value)) answer = "예";
    }
    form.yesNo[item.name] = { found: true, answer, densities: pair.map((o) => o.density) };
  });

  // 3) 작업구분: '작업구분' 줄과, 그 위·아래로 이어지는 선택지 줄(항목 이름이 없는 줄)
  //    Safe_B는 '작업 구분' 이름이 두 줄 버튼의 가운데에 있어, 버튼 첫 줄이 이름 줄보다 위에 읽힌다.
  let startIndex = lines.findIndex((line) => labelMatches(line, allLabels).some((m) => m.name === "작업구분"));
  if (startIndex < 0) {
    // '작업 구분' 이름을 못 읽었으면, 버튼 이름이 작업구분과 2개 이상 맞는 줄을 작업구분 줄로 본다
    startIndex = lines.findIndex((line) => (permit.options || [])
      .filter((o) => o.chip && o.lineText === line)
      .filter((o) => rules.categories.some((c) => new RegExp(c.match).test(compact(o.text))) || closestCategory(o.text, rules.categories) !== null)
      .length >= 2);
  }
  if (startIndex >= 0) {
    const hasOptions = (line) => (permit.options || []).some((o) => o.lineText === line);
    const isChoiceLine = (line) => !labelMatches(line, allLabels).length && hasOptions(line);
    let first = startIndex;
    while (first - 1 >= 0 && isChoiceLine(lines[first - 1])) first -= 1;
    let last = startIndex;
    while (last + 1 < lines.length && isChoiceLine(lines[last + 1])) last += 1;
    // 선택 표시가 없는 화면(조회 화면 등)은 다음 항목 이름 줄 전까지
    if (last === startIndex && !hasOptions(lines[startIndex])) {
      while (last + 1 < lines.length && !labelMatches(lines[last + 1], allLabels).length) last += 1;
    }
    const area = lines.slice(first, last + 1);
    area.forEach((line) => form.categoryLines.add(line));
    const options = (permit.options || []).filter((o) => form.categoryLines.has(o.lineText));
    const areaText = compact(area.join(" ").replace(new RegExp(rules.categoryLabel), ""));
    // 버튼 상자는 화면 위치(줄 순서 → 왼쪽부터) 순서로 놓는다
    const chips = options.filter((o) => o.chip)
      .sort((a, b) => area.indexOf(a.lineText) - area.indexOf(b.lineText) || (a.x ?? a.textIndex) - (b.x ?? b.textIndex));
    if (chips.length) {
      // Safe_B 버튼: 버튼 수가 작업구분 수와 같으면 순서대로 짝짓는다(이름을 못 읽은 버튼도 알아낼 수 있음).
      // 다르면 이름이 읽힌 버튼을 기준점으로 삼고, 기준점 사이 버튼 수가 맞는 구간만 순서로 채운다.
      const assigned = new Array(chips.length).fill(null);
      if (chips.length === rules.categories.length) {
        chips.forEach((chip, i) => { assigned[i] = i; });
      } else {
        chips.forEach((chip, i) => {
          const index = rules.categories.findIndex((c) => new RegExp(c.match).test(compact(chip.text)));
          assigned[i] = index >= 0 ? index : closestCategory(chip.text, rules.categories);
        });
        const anchors = assigned.map((c, i) => [i, c]).filter(([, c]) => c !== null);
        for (let k = 0; k + 1 < anchors.length; k += 1) {
          const [i1, c1] = anchors[k];
          const [i2, c2] = anchors[k + 1];
          if (i2 - i1 === c2 - c1) for (let i = i1 + 1; i < i2; i += 1) assigned[i] = c1 + (i - i1);
        }
      }
      rules.categories.forEach((category, index) => {
        const i = assigned.indexOf(index);
        form.categories.push({ ...category, option: i >= 0 ? chips[i] : null, selected: i >= 0 ? chips[i].filled : null });
      });
      form.categoriesDecided = form.categories.every((c) => c.selected !== null);
      // 못 짝지은 칸이 있어도, 파란(선택된) 버튼이 모두 짝지어졌으면 선택 판단은 된 것으로 본다
      if (!form.categoriesDecided && chips.filter((c) => c.filled).every((c) => assigned[chips.indexOf(c)] !== null)) {
        form.categories.forEach((c) => { if (c.selected === null) c.selected = false; });
        form.categoriesDecided = true;
      }
      return form;
    }
    rules.categories.forEach((category) => {
      const pattern = new RegExp(category.match);
      const option = options.find((o) => pattern.test(compact(o.text)));
      if (option) form.categories.push({ ...category, option });
      else if (new RegExp(category.match.replace(/^\^/, "")).test(areaText)) form.categories.push({ ...category, option: null });
    });
    const withOption = form.categories.filter((c) => c.option);
    if (withOption.length) {
      const optionList = withOption.map((c) => c.option);
      decideFilled(optionList);
      form.categories.forEach((c) => { c.selected = c.option ? c.option.filled : null; });
    } else {
      // 선택 표시가 전혀 없으면 화면에 적힌 작업이 고른 작업이다(조회 화면처럼 고른 것만 보이는 경우)
      form.categories.forEach((c) => { c.selected = true; });
    }
    form.categoriesDecided = form.categories.length > 0 && form.categories.every((c) => c.selected !== null && c.selected !== undefined);
  }
  return form;
}

// 작업구분에서 고른 작업 → 검토에 쓰는 작업 종류(review.js workTypes id)
function permitWorkTypeIds(form) {
  return [...new Set(form.categories.filter((c) => c.selected).flatMap((c) => c.workTypes))];
}

function reviewPermitForm(form, docs, add) {
  const rules = window.PERMIT_RULES.permitForm;
  const { plan, education } = docs;
  const label = DOC_KINDS.permit;

  Object.entries(form.fields).forEach(([name, field]) => {
    if (!field.found) {
      add(PERMIT_SECTION, label, `${name} 입력`, "warn", `화면에서 '${name}' 항목을 찾지 못했습니다. 캡처 범위에 들어갔는지 확인하세요.`);
    } else {
      add(PERMIT_SECTION, label, `${name} 입력`, field.empty ? "fail" : "ok",
        field.empty ? "비어 있거나 선택되지 않았습니다." : `“${field.value.slice(0, 60)}”`);
    }
  });

  Object.entries(form.yesNo).forEach(([name, item]) => {
    let status = "manual";
    let detail = "선택 상태를 읽지 못했습니다. 화면에서 '아니오'로 선택됐는지 확인하세요.";
    if (!item.found) {
      status = "warn";
      detail = `화면에서 '${name}' 항목을 찾지 못했습니다. 캡처 범위에 들어갔는지 확인하세요.`;
    } else if (item.answer === "아니오") {
      status = "ok";
      detail = "'아니오'로 선택되어 있습니다.";
    } else if (item.answer === "예") {
      status = "fail";
      detail = "'예'로 선택되어 있습니다. 반드시 '아니오'여야 합니다.";
    }
    add(PERMIT_SECTION, label, `${name}: 아니오`, status, detail);
  });

  // 작업구분 ↔ 서류
  if (!form.categories.length) {
    add(PERMIT_SECTION, label, "작업구분", "warn", "작업구분 항목을 찾지 못했습니다. 캡처 범위에 들어갔는지 확인하세요.");
    return;
  }
  const selected = form.categories.filter((c) => c.selected);
  add(PERMIT_SECTION, label, "작업구분 선택", form.categoriesDecided ? (selected.length ? "info" : "warn") : "manual",
    form.categoriesDecided
      ? (selected.length ? `선택: ${selected.map((c) => c.name).join(", ")}` : "선택된 작업구분이 없습니다.")
      : `일부 칸의 선택 여부를 읽지 못했습니다(읽은 선택: ${selected.map((c) => c.name).join(", ") || "없음"}). 화면에서 확인하세요.`);

  const planFiles = plan ? plan.files : [];
  const educationFiles = education ? education.files : [];
  const handlingPlans = planFiles.filter((f) => hasAny(f.text, rules.handlingPlanTitle));
  const page1Options = (file) => {
    const options = (file.options || []).filter((o) => o.page === 0);
    decideFilled(options);
    return options;
  };
  const page1Text = (file) => (file.pages && file.pages.length ? file.pages[0].lines : file.lines).join(" ");

  selected.forEach((category) => {
    category.requires.forEach((req) => {
      const title = `[${category.name}] ${req.label}`;
      if (req.doc === "education") {
        const files = educationFiles.filter((f) => hasAny(f.text, req.keywords));
        add(PERMIT_SECTION, "작업구분 ↔ 특별안전교육", title, files.length ? "ok" : "fail",
          files.length ? files.map((f) => f.fileName).join(", ") : "이 작업의 특별교육 실시 결과가 없습니다.");
      } else if (req.doc === "planTitle") {
        const files = planFiles.filter((f) => hasAny(f.text, req.keywords));
        add(PERMIT_SECTION, "작업구분 ↔ 작업계획서", title, files.length ? "ok" : "fail",
          files.length ? files.map((f) => f.fileName).join(", ") : `${req.label}가 첨부되지 않았습니다.`);
      } else if (req.doc === "handlingPage1") {
        if (!handlingPlans.length) {
          add(PERMIT_SECTION, "작업구분 ↔ 작업계획서", title, "fail", "차량계 하역운반기계 작업계획서가 첨부되지 않았습니다.");
          return;
        }
        let best = null;
        handlingPlans.forEach((file) => {
          const option = page1Options(file).find((o) => hasAny(o.text, req.keywords));
          if (option && (!best || option.filled)) best = { file, option };
          else if (!best && hasAny(page1Text(file), req.keywords)) best = { file, option: null };
        });
        let status = "fail";
        let detail = "계획서 1쪽에 이 장비 항목이 없습니다.";
        if (best && best.option && best.option.filled === true) {
          status = "ok";
          detail = `${best.file.fileName} 1쪽 “${best.option.text.slice(0, 40)}” 체크됨`;
        } else if (best && best.option && best.option.filled === false) {
          detail = `${best.file.fileName} 1쪽 “${best.option.text.slice(0, 40)}”이(가) 체크되어 있지 않습니다.`;
        } else if (best) {
          status = "manual";
          detail = `${best.file.fileName} 1쪽에 항목은 있으나 체크 여부를 읽지 못했습니다. 직접 확인하세요.`;
        }
        add(PERMIT_SECTION, "작업구분 ↔ 작업계획서", title, status, detail);
      }
    });
  });

  // 반대 방향: 서류에는 있는데 작업구분에서 고르지 않은 작업
  if (!form.categoriesDecided) return;
  const isSelected = (id) => selected.some((c) => c.id === id);
  const missing = [];
  rules.categories.forEach((category) => {
    if (isSelected(category.id)) return;
    category.requires.forEach((req) => {
      let evidence = null;
      if (req.doc === "education") {
        evidence = educationFiles.find((f) => hasAny(f.text, req.keywords));
      } else if (req.doc === "planTitle") {
        evidence = planFiles.find((f) => hasAny(f.text, req.keywords));
      } else if (req.doc === "handlingPage1") {
        evidence = handlingPlans.find((f) => page1Options(f).some((o) => o.filled && hasAny(o.text, req.keywords)));
      }
      if (evidence) missing.push(`${category.name}(${evidence.fileName})`);
    });
  });
  if (missing.length) {
    add(PERMIT_SECTION, "작업구분 ↔ 서류", "서류에는 있는데 작업구분에서 빠진 작업", "warn",
      `${missing.join(", ")} — 실제로 하는 작업이면 작업구분에서 선택해야 합니다.`);
  }
}
