const GROUP_CONFIG = [
  { key: 'knowledge', label: '知の活用力', cols: ['知①', '知②'], requiredWeights: [200, 200] },
  { key: 'management', label: 'マネジメント能力', cols: ['管①', '管②'], requiredWeights: [150, 100] },
  { key: 'communication', label: 'コミュニケーション能力', cols: ['コ①', 'コ②'], requiredWeights: [150, 100] },
  { key: 'teamwork', label: 'チームワーク力', cols: ['チ①', 'チ②'], requiredWeights: [150, 100] },
  { key: 'internationality', label: '国際性', cols: ['国①', '国②'], requiredWeights: [150, 100] },
  { key: 'research', label: '研究力', cols: ['研①', '研②', '研③'], requiredWeights: [100, 500, 100] },
  { key: 'specializedKnowledge', label: '専門知識', cols: ['専①', '専②', '専③'], requiredWeights: [100, 100, 500] },
  { key: 'ethics', label: '倫理観', cols: ['倫①', '倫②'], requiredWeights: [150, 50] }
];

const DEFAULT_ROW_NAMES = {
  11: '情報理工前期特別研究A',
  12: '情報理工前期特別研究B',
  13: '情報理工前期特別研究C',
  14: '情報理工前期特別研究D',
  15: '',
  16: '',
  17: '',
  18: '',
  19: '',
  20: ''
};

const DEFAULT_ROW_LABELS = {
  11: '11行: 情報理工前期特別研究A',
  12: '12行: 情報理工前期特別研究B',
  13: '13行: 情報理工前期特別研究C',
  14: '14行: 情報理工前期特別研究D',
  15: '15行: 空き行',
  16: '16行: 空き行',
  17: '17行: 空き行',
  18: '18行: 空き行',
  19: '19行: 空き行',
  20: '20行: 空き行'
};

const COLUMN_LABELS = GROUP_CONFIG.flatMap(group => group.cols);
const REQUIRED_POINTS18 = GROUP_CONFIG.flatMap(group => group.requiredWeights);
const COLUMN_LETTERS = Array.from({ length: 18 }, (_, idx) => String.fromCharCode('B'.charCodeAt(0) + idx));

const state = {
  dataset: null,
  courses: [],
  workbook: null,
  workbookFileName: '',
  sheetName: 'M中間評価',
  selectedCourse: null,
  previewPoints18: Array(18).fill(0),
  planRows: {}
};

const els = {};

document.addEventListener('DOMContentLoaded', async () => {
  bindElements();
  wireEvents();
  renderRowOptions();
  renderPreviewGrid();
  renderPlanSection();
  updateActionStates();
  await loadDataset();
});

function bindElements() {
  els.xlsxFile = document.getElementById('xlsxFile');
  els.workbookStatus = document.getElementById('workbookStatus');
  els.autofillRequiredBtn = document.getElementById('autofillRequiredBtn');
  els.downloadBtn = document.getElementById('downloadBtn');
  els.searchInput = document.getElementById('searchInput');
  els.searchMeta = document.getElementById('searchMeta');
  els.searchResults = document.getElementById('searchResults');
  els.selectedCourseBox = document.getElementById('selectedCourseBox');
  els.splitMode = document.getElementById('splitMode');
  els.targetRow = document.getElementById('targetRow');
  els.previewGrid = document.getElementById('previewGrid');
  els.previewTotal = document.getElementById('previewTotal');
  els.planAddBtn = document.getElementById('planAddBtn');
  els.writeRowBtn = document.getElementById('writeRowBtn');
  els.applyPlanBtn = document.getElementById('applyPlanBtn');
  els.reloadPlanBtn = document.getElementById('reloadPlanBtn');
  els.clearPlanBtn = document.getElementById('clearPlanBtn');
  els.planList = document.getElementById('planList');
  els.planCourseCount = document.getElementById('planCourseCount');
  els.planCredits = document.getElementById('planCredits');
  els.planPointsTotal = document.getElementById('planPointsTotal');
  els.extraPointsTotal = document.getElementById('extraPointsTotal');
  els.projectedTotal = document.getElementById('projectedTotal');
  els.unmetCount = document.getElementById('unmetCount');
  els.totalsBody = document.getElementById('totalsBody');
  els.totalsFoot = document.getElementById('totalsFoot');
  els.datasetInfo = document.getElementById('datasetInfo');
}

function wireEvents() {
  els.xlsxFile.addEventListener('change', onWorkbookSelected);
  els.searchInput.addEventListener('input', () => renderSearchResults(searchCourses(els.searchInput.value)));
  els.splitMode.addEventListener('change', onSplitModeChanged);
  els.planAddBtn.addEventListener('click', addSelectedCourseToPlan);
  els.writeRowBtn.addEventListener('click', writeSelectedCourseToWorkbook);
  els.autofillRequiredBtn.addEventListener('click', autofillExistingRows);
  els.applyPlanBtn.addEventListener('click', applyPlanToWorkbook);
  els.reloadPlanBtn.addEventListener('click', reloadPlanFromWorkbook);
  els.clearPlanBtn.addEventListener('click', clearPlanRows);
  els.downloadBtn.addEventListener('click', downloadWorkbook);
}

async function loadDataset() {
  try {
    const response = await fetch('./data/courses.sample.json');
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    state.dataset = await response.json();
    state.courses = state.dataset.courses || [];
    els.searchMeta.textContent = `${state.courses.length} 件のサンプル科目を読み込みました。`;
    els.datasetInfo.innerHTML = [
      `<div><strong>データ件数:</strong> ${state.dataset.meta?.count ?? state.courses.length} 件</div>`,
      `<div><strong>データソース:</strong> <a href="${escapeHtml(state.dataset.meta?.sourceUrl || '#')}" target="_blank" rel="noreferrer">公開カリキュラム・マップ</a></div>`,
      `<div><strong>備考:</strong> ${escapeHtml((state.dataset.meta?.notes || []).join(' / '))}</div>`
    ].join('');
    renderSearchResults(searchCourses(''));

    if (state.workbook) {
      syncPlanFromWorkbook();
      setWorkbookStatus(`読込済みの Excel から履修計画を再解釈しました。`, 'ok');
    }
  } catch (error) {
    console.error(error);
    els.searchMeta.textContent = '科目データの読込に失敗しました。';
    els.datasetInfo.textContent = 'data/courses.sample.json を確認してください。';
  }
}

function searchCourses(rawTerm) {
  const term = normalize(rawTerm);
  const courses = state.courses.slice();
  if (!term) {
    return courses.slice(0, 12);
  }

  const scored = courses
    .map(course => ({ course, score: scoreCourse(course, term) }))
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score || a.course.code.localeCompare(b.course.code, 'ja'));

  return scored.slice(0, 24).map(item => item.course);
}

function scoreCourse(course, term) {
  const code = normalize(course.code);
  const ja = normalize(course.nameJa);
  const en = normalize(course.nameEn);
  const aliases = (course.aliases || []).map(normalize);

  if (code === term) return 1000;
  if (ja === term || en === term || aliases.includes(term)) return 900;
  if (code.startsWith(term)) return 800;
  if (ja.includes(term)) return 700;
  if (en.includes(term)) return 650;
  if (aliases.some(alias => alias.includes(term))) return 600;
  return 0;
}

function renderSearchResults(courses) {
  if (!courses.length) {
    els.searchResults.innerHTML = '<div class="status status--warn">該当科目が見つかりません。</div>';
    return;
  }

  const fragment = document.createDocumentFragment();
  courses.forEach(course => {
    const card = document.createElement('article');
    card.className = 'result-card';

    const p8 = course.points8 || {};
    const metaChips = [
      chip(`単位 ${course.credits ?? '—'}`),
      chip(`合計 ${course.total ?? sumPoints8(p8)} 点`),
      chip(`source p.${course.sourcePage ?? '—'}`)
    ].join('');

    card.innerHTML = `
      <div class="result-card__top">
        <div>
          <div class="result-card__code">${escapeHtml(course.code)}</div>
          <h3>${escapeHtml(course.nameJa)}</h3>
          <p>${escapeHtml(course.nameEn || '')}</p>
        </div>
        <button class="button button--primary" type="button">選択</button>
      </div>
      <div class="result-card__meta">${metaChips}</div>
    `;

    card.querySelector('button').addEventListener('click', () => selectCourse(course));
    fragment.appendChild(card);
  });

  els.searchResults.replaceChildren(fragment);
}

function selectCourse(course) {
  state.selectedCourse = course;
  recomputePreview();
  renderSelectedCourse(course);
  updateActionStates();
}

function renderSelectedCourse(course) {
  if (!course) {
    els.selectedCourseBox.classList.add('empty');
    els.selectedCourseBox.textContent = 'まだ科目が選択されていません。';
    return;
  }

  const p8 = course.points8 || {};
  const hasPoints8 = Object.keys(p8).length > 0;
  const pointsHtml = hasPoints8
    ? GROUP_CONFIG.map(group => {
        const value = Number(p8[group.key] || 0);
        return `
          <div class="point8-item">
            <span>${escapeHtml(group.label)}</span>
            <strong>${value}</strong>
          </div>
        `;
      }).join('')
    : `
      <div class="point8-item point8-item--wide">
        <span>18 観点の直接入力値</span>
        <strong>${sumPoints18(course.points18 || [])} 点</strong>
      </div>
    `;

  const subline = [course.nameEn || '', course.credits ? `${course.credits} 単位` : '', course.code || '']
    .filter(Boolean)
    .join(' / ');

  els.selectedCourseBox.classList.remove('empty');
  els.selectedCourseBox.innerHTML = `
    <div class="result-card__code">${escapeHtml(course.code || 'カスタム')}</div>
    <h3>${escapeHtml(course.nameJa || '名称未設定')}</h3>
    <div class="muted">${escapeHtml(subline)}</div>
    <div class="points8">${pointsHtml}</div>
  `;
}

function onSplitModeChanged() {
  recomputePreview();
  if (state.workbook) {
    renderPlanSection();
  }
}

function recomputePreview() {
  if (!state.selectedCourse) {
    state.previewPoints18 = Array(18).fill(0);
    updatePreviewInputs();
    return;
  }

  if (Array.isArray(state.selectedCourse.points18) && state.selectedCourse.points18.length === 18) {
    state.previewPoints18 = clonePoints18(state.selectedCourse.points18);
  } else {
    state.previewPoints18 = expandPoints8To18(state.selectedCourse.points8 || {}, els.splitMode.value);
  }
  updatePreviewInputs();
}

function expandPoints8To18(points8, mode) {
  const values = [];
  GROUP_CONFIG.forEach(group => {
    const total = Number(points8[group.key] || 0);
    const weights = mode === 'even'
      ? Array(group.cols.length).fill(1)
      : group.requiredWeights;
    values.push(...allocateByWeights(total, weights));
  });
  return values;
}

function allocateByWeights(total, weights) {
  const safeTotal = Math.max(0, Math.round(Number(total || 0)));
  if (safeTotal === 0) {
    return weights.map(() => 0);
  }

  const weightSum = weights.reduce((sum, value) => sum + value, 0);
  const raw = weights.map(weight => (safeTotal * weight) / weightSum);
  const floors = raw.map(value => Math.floor(value));
  let remainder = safeTotal - floors.reduce((sum, value) => sum + value, 0);

  const order = raw
    .map((value, index) => ({ index, frac: value - Math.floor(value) }))
    .sort((a, b) => b.frac - a.frac || a.index - b.index);

  const allocated = floors.slice();
  for (let i = 0; i < remainder; i += 1) {
    allocated[order[i % order.length].index] += 1;
  }
  return allocated;
}

function renderPreviewGrid() {
  els.previewGrid.innerHTML = '';
  COLUMN_LABELS.forEach((label, index) => {
    const wrapper = document.createElement('label');
    wrapper.className = 'preview-cell';
    wrapper.innerHTML = `
      <span class="preview-cell__label">${COLUMN_LETTERS[index]}列 / ${escapeHtml(label)}</span>
      <input type="number" min="0" step="1" value="0">
    `;
    const input = wrapper.querySelector('input');
    input.addEventListener('input', () => {
      state.previewPoints18[index] = asInt(input.value);
      updatePreviewTotal();
    });
    els.previewGrid.appendChild(wrapper);
  });
  updatePreviewInputs();
}

function updatePreviewInputs() {
  const inputs = [...els.previewGrid.querySelectorAll('input')];
  inputs.forEach((input, index) => {
    input.value = String(asInt(state.previewPoints18[index] || 0));
  });
  updatePreviewTotal();
}

function updatePreviewTotal() {
  const total = sumPoints18(state.previewPoints18);
  const sourceTotal = state.selectedCourse
    ? Number(
        state.selectedCourse.total
        ?? (Array.isArray(state.selectedCourse.points18) ? sumPoints18(state.selectedCourse.points18) : sumPoints8(state.selectedCourse.points8 || {}))
      )
    : 0;
  const diff = total - sourceTotal;
  const suffix = diff === 0 ? '' : ` / 元データとの差 ${diff > 0 ? '+' : ''}${diff}`;
  els.previewTotal.textContent = `合計 ${total} 点${suffix}`;
}

async function onWorkbookSelected(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  if (!window.XlsxPopulate) {
    setWorkbookStatus('XlsxPopulate の読込に失敗しています。ネットワーク接続を確認してください。', 'warn');
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    state.workbook = await window.XlsxPopulate.fromDataAsync(buffer);
    state.workbookFileName = file.name;
    state.sheetName = detectSheetName(state.workbook);
    renderRowOptions();
    const imported = syncPlanFromWorkbook();
    setWorkbookStatus(`読込完了: ${file.name} / シート: ${state.sheetName} / 履修計画 ${imported} 行を取り込みました。`, 'ok');
    updateActionStates();
  } catch (error) {
    console.error(error);
    state.workbook = null;
    state.planRows = {};
    renderPlanSection();
    updateActionStates();
    setWorkbookStatus('Excel の読込に失敗しました。ファイル形式を確認してください。', 'warn');
  }
}

function detectSheetName(workbook) {
  const preferred = ['M中間評価', 'M最終審査'];
  const names = workbook.sheets().map(sheet => sheet.name());
  return preferred.find(name => names.includes(name)) || names[0] || 'Sheet1';
}

function renderRowOptions() {
  els.targetRow.innerHTML = '';
  for (let row = 11; row <= 20; row += 1) {
    const option = document.createElement('option');
    option.value = String(row);
    option.textContent = getRowLabel(row);
    els.targetRow.appendChild(option);
  }
}

function getRowLabel(row) {
  const planned = state.planRows[row];
  if (planned) {
    return `${row}行: ${planned.nameJa}`;
  }

  if (!state.workbook) {
    return DEFAULT_ROW_LABELS[row] || `${row}行`;
  }
  try {
    const sheet = state.workbook.sheet(state.sheetName);
    const courseName = String(sheet.cell(`A${row}`).value() || '').trim();
    if (courseName) {
      return `${row}行: ${courseName}`;
    }
  } catch (error) {
    console.warn(error);
  }
  return DEFAULT_ROW_LABELS[row] || `${row}行`;
}

function setWorkbookStatus(message, kind = 'muted') {
  els.workbookStatus.textContent = message;
  els.workbookStatus.className = `status status--${kind}`;
}

function findCourseByNameOrAlias(rawName) {
  const target = normalize(rawName);
  if (!target) return null;

  return state.courses.find(course => {
    const candidates = [course.nameJa, course.nameEn, course.code, ...(course.aliases || [])]
      .filter(Boolean)
      .map(normalize);
    return candidates.includes(target);
  }) || null;
}

function rowIsEmpty(sheet, row) {
  return COLUMN_LETTERS.every(letter => {
    const value = sheet.cell(`${letter}${row}`).value();
    return value === undefined || value === null || value === '' || Number(value) === 0;
  });
}

function writePreviewToRow(sheet, row, courseName) {
  sheet.cell(`A${row}`).value(courseName);
  COLUMN_LETTERS.forEach((letter, index) => {
    sheet.cell(`${letter}${row}`).value(asInt(state.previewPoints18[index] || 0));
  });
  const totalCell = sheet.cell(`T${row}`);
  if (!String(totalCell.formula() || '').trim()) {
    totalCell.formula(`SUM(B${row}:S${row})`);
  }
}

function writePoints18ToRow(sheet, row, courseName, points18) {
  sheet.cell(`A${row}`).value(courseName);
  COLUMN_LETTERS.forEach((letter, index) => {
    sheet.cell(`${letter}${row}`).value(asInt(points18[index] || 0));
  });
  const totalCell = sheet.cell(`T${row}`);
  if (!String(totalCell.formula() || '').trim()) {
    totalCell.formula(`SUM(B${row}:S${row})`);
  }
}

function clearCourseRow(sheet, row) {
  sheet.cell(`A${row}`).value(DEFAULT_ROW_NAMES[row] || '');
  COLUMN_LETTERS.forEach(letter => {
    sheet.cell(`${letter}${row}`).value('');
  });
  const totalCell = sheet.cell(`T${row}`);
  if (!String(totalCell.formula() || '').trim()) {
    totalCell.formula(`SUM(B${row}:S${row})`);
  }
}

function resolveCoursePoints18(course) {
  if (Array.isArray(course?.points18) && course.points18.length === 18) {
    return clonePoints18(course.points18);
  }
  return expandPoints8To18(course?.points8 || {}, els.splitMode.value);
}

function makePlanEntryFromSelection(row) {
  return {
    row,
    code: state.selectedCourse?.code || '',
    nameJa: state.selectedCourse?.nameJa || '',
    nameEn: state.selectedCourse?.nameEn || '',
    credits: Number(state.selectedCourse?.credits || 0),
    sourceTotal: Number(state.selectedCourse?.total ?? sumPoints18(state.previewPoints18)),
    points18: clonePoints18(state.previewPoints18)
  };
}

function addSelectedCourseToPlan() {
  if (!state.selectedCourse) return;
  const row = Number(els.targetRow.value || 15);
  state.planRows[row] = makePlanEntryFromSelection(row);
  renderRowOptions();
  renderPlanSection();
  selectNextAvailableRow(row);
  setWorkbookStatus(`${row}行を履修計画に追加 / 更新しました。`, 'ok');
}

function writeSelectedCourseToWorkbook() {
  if (!state.workbook || !state.selectedCourse) return;
  const row = Number(els.targetRow.value || 15);
  state.planRows[row] = makePlanEntryFromSelection(row);
  const sheet = state.workbook.sheet(state.sheetName);
  writePreviewToRow(sheet, row, state.selectedCourse.nameJa);
  renderRowOptions();
  renderPlanSection();
  selectNextAvailableRow(row);
  setWorkbookStatus(`${row}行へ「${state.selectedCourse.nameJa}」を書き込みました。`, 'ok');
}

function applyPlanToWorkbook() {
  if (!state.workbook) return;
  const sheet = state.workbook.sheet(state.sheetName);
  for (let row = 11; row <= 20; row += 1) {
    const entry = state.planRows[row];
    if (entry) {
      writePoints18ToRow(sheet, row, entry.nameJa, entry.points18);
    } else {
      clearCourseRow(sheet, row);
    }
  }
  renderRowOptions();
  setWorkbookStatus(`履修計画 ${getPlanEntries().length} 行を Excel に反映しました。`, 'ok');
}

function syncPlanFromWorkbook() {
  if (!state.workbook) {
    state.planRows = {};
    renderPlanSection();
    updateActionStates();
    return 0;
  }

  const imported = {};
  const sheet = state.workbook.sheet(state.sheetName);

  for (let row = 11; row <= 20; row += 1) {
    const name = String(sheet.cell(`A${row}`).value() || '').trim();
    const points18FromSheet = COLUMN_LETTERS.map(letter => asInt(sheet.cell(`${letter}${row}`).value()));
    const hasAnyPoints = points18FromSheet.some(value => value > 0);
    const matched = findCourseByNameOrAlias(name);
    const isNamedRow = Boolean(name);

    if (!isNamedRow && !hasAnyPoints) {
      continue;
    }

    const resolvedName = matched?.nameJa || name || DEFAULT_ROW_NAMES[row] || '';
    const resolvedPoints18 = hasAnyPoints
      ? points18FromSheet
      : (matched ? resolveCoursePoints18(matched) : points18FromSheet);

    if (!resolvedName && sumPoints18(resolvedPoints18) === 0) {
      continue;
    }

    imported[row] = {
      row,
      code: matched?.code || '',
      nameJa: resolvedName,
      nameEn: matched?.nameEn || '',
      credits: Number(matched?.credits || 0),
      sourceTotal: Number(matched?.total ?? sumPoints18(resolvedPoints18)),
      points18: clonePoints18(resolvedPoints18)
    };
  }

  state.planRows = imported;
  renderRowOptions();
  renderPlanSection();
  updateActionStates();
  return Object.keys(imported).length;
}

function reloadPlanFromWorkbook() {
  if (!state.workbook) return;
  const imported = syncPlanFromWorkbook();
  setWorkbookStatus(`Excel から履修計画を再読込しました。${imported} 行を反映しています。`, 'ok');
}

function clearPlanRows() {
  state.planRows = {};
  renderRowOptions();
  renderPlanSection();
  updateActionStates();
  setWorkbookStatus('履修計画をクリアしました。', 'ok');
}

function removePlanRow(row) {
  delete state.planRows[row];
  renderRowOptions();
  renderPlanSection();
  updateActionStates();
  setWorkbookStatus(`${row}行を履修計画から削除しました。`, 'ok');
}

function editPlanRow(row) {
  const entry = state.planRows[row];
  if (!entry) return;

  els.targetRow.value = String(row);
  const matched = findCourseByNameOrAlias(entry.code || entry.nameJa);
  state.selectedCourse = matched || {
    code: entry.code,
    nameJa: entry.nameJa,
    nameEn: entry.nameEn,
    credits: entry.credits,
    total: sumPoints18(entry.points18),
    points18: clonePoints18(entry.points18)
  };
  state.previewPoints18 = clonePoints18(entry.points18);
  renderSelectedCourse(state.selectedCourse);
  updatePreviewInputs();
  updateActionStates();
  setWorkbookStatus(`${row}行の内容を編集パネルへ読み込みました。`, 'ok');
}

function selectNextAvailableRow(currentRow) {
  const candidates = [];
  for (let row = currentRow + 1; row <= 20; row += 1) {
    candidates.push(row);
  }
  for (let row = 11; row <= currentRow; row += 1) {
    candidates.push(row);
  }
  const nextEmpty = candidates.find(row => !state.planRows[row]);
  els.targetRow.value = String(nextEmpty || currentRow);
}

function getPlanEntries() {
  return Object.values(state.planRows).sort((a, b) => a.row - b.row);
}

function getPlanTotals18() {
  const totals = Array(18).fill(0);
  getPlanEntries().forEach(entry => {
    entry.points18.forEach((value, index) => {
      totals[index] += asInt(value);
    });
  });
  return totals;
}

function getExtraTotals18() {
  const totals = Array(18).fill(0);
  if (!state.workbook) return totals;

  try {
    const sheet = state.workbook.sheet(state.sheetName);
    for (let row = 24; row <= 26; row += 1) {
      COLUMN_LETTERS.forEach((letter, index) => {
        totals[index] += asInt(sheet.cell(`${letter}${row}`).value());
      });
    }
  } catch (error) {
    console.warn(error);
  }
  return totals;
}

function renderPlanSection() {
  renderPlanList();
  renderTotalsTable();
  updateActionStates();
}

function renderPlanList() {
  const entries = getPlanEntries();
  if (!entries.length) {
    els.planList.innerHTML = '<div class="status status--muted">まだ履修計画はありません。科目を選んで「履修計画に追加 / 更新」を押してください。</div>';
  } else {
    const html = entries.map(entry => {
      const total = sumPoints18(entry.points18);
      const codeText = entry.code ? `${entry.code} / ` : '';
      return `
        <article class="plan-card">
          <div class="plan-card__top">
            <div>
              <div class="plan-card__row">${entry.row}行</div>
              <h3>${escapeHtml(entry.nameJa)}</h3>
              <p>${escapeHtml(codeText)}${escapeHtml(entry.credits ? `${entry.credits} 単位` : '単位不明')}</p>
            </div>
            <div class="plan-card__total">${total} 点</div>
          </div>
          <div class="actions-row actions-row--compact">
            <button class="button button--small" type="button" data-action="edit" data-row="${entry.row}">編集</button>
            <button class="button button--small button--danger" type="button" data-action="remove" data-row="${entry.row}">削除</button>
          </div>
        </article>
      `;
    }).join('');

    els.planList.innerHTML = html;
    els.planList.querySelectorAll('button[data-action="remove"]').forEach(button => {
      button.addEventListener('click', () => removePlanRow(Number(button.dataset.row)));
    });
    els.planList.querySelectorAll('button[data-action="edit"]').forEach(button => {
      button.addEventListener('click', () => editPlanRow(Number(button.dataset.row)));
    });
  }

  const planTotals = getPlanTotals18();
  const extraTotals = getExtraTotals18();
  const projectedTotals = planTotals.map((value, index) => value + extraTotals[index]);
  const unmetCount = REQUIRED_POINTS18.filter((required, index) => projectedTotals[index] < required).length;

  els.planCourseCount.textContent = String(entries.length);
  els.planCredits.textContent = String(entries.reduce((sum, entry) => sum + Number(entry.credits || 0), 0));
  els.planPointsTotal.textContent = String(sumPoints18(planTotals));
  els.extraPointsTotal.textContent = String(sumPoints18(extraTotals));
  els.projectedTotal.textContent = String(sumPoints18(projectedTotals));
  els.unmetCount.textContent = `${unmetCount} / 18`;
}

function renderTotalsTable() {
  const planTotals = getPlanTotals18();
  const extraTotals = getExtraTotals18();
  const projectedTotals = planTotals.map((value, index) => value + extraTotals[index]);

  els.totalsBody.innerHTML = COLUMN_LABELS.map((label, index) => {
    const required = REQUIRED_POINTS18[index];
    const plan = planTotals[index];
    const extra = extraTotals[index];
    const projected = projectedTotals[index];
    const deficit = Math.max(required - projected, 0);
    const status = deficit === 0
      ? '<span class="badge badge--ok">OK</span>'
      : `<span class="badge badge--warn">不足 ${deficit}</span>`;

    return `
      <tr>
        <th>${escapeHtml(`${COLUMN_LETTERS[index]} / ${label}`)}</th>
        <td class="num">${required}</td>
        <td class="num">${plan}</td>
        <td class="num">${extra}</td>
        <td class="num">${projected}</td>
        <td class="num ${deficit > 0 ? 'num--deficit' : 'num--ok'}">${deficit}</td>
        <td>${status}</td>
      </tr>
    `;
  }).join('');

  const requiredTotal = sumPoints18(REQUIRED_POINTS18);
  const planTotal = sumPoints18(planTotals);
  const extraTotal = sumPoints18(extraTotals);
  const projectedTotal = sumPoints18(projectedTotals);
  const deficitTotal = REQUIRED_POINTS18.reduce((sum, required, index) => sum + Math.max(required - projectedTotals[index], 0), 0);
  const overallGap = Math.max(requiredTotal - projectedTotal, 0);

  els.totalsFoot.innerHTML = `
    <tr>
      <th>合計</th>
      <td class="num">${requiredTotal}</td>
      <td class="num">${planTotal}</td>
      <td class="num">${extraTotal}</td>
      <td class="num">${projectedTotal}</td>
      <td class="num ${deficitTotal > 0 ? 'num--deficit' : 'num--ok'}">${deficitTotal}</td>
      <td>${overallGap === 0 ? '<span class="badge badge--ok">総点 OK</span>' : `<span class="badge badge--warn">総点不足 ${overallGap}</span>`}</td>
    </tr>
  `;
}

function updateActionStates() {
  const hasWorkbook = Boolean(state.workbook);
  const hasSelection = Boolean(state.selectedCourse);
  const hasPlan = getPlanEntries().length > 0;

  els.autofillRequiredBtn.disabled = !hasWorkbook;
  els.downloadBtn.disabled = !hasWorkbook;
  els.planAddBtn.disabled = !hasSelection;
  els.writeRowBtn.disabled = !hasWorkbook || !hasSelection;
  els.applyPlanBtn.disabled = !hasWorkbook || !hasPlan;
  els.reloadPlanBtn.disabled = !hasWorkbook;
  els.clearPlanBtn.disabled = !hasPlan;
}

function autofillExistingRows() {
  if (!state.workbook) return;
  const sheet = state.workbook.sheet(state.sheetName);
  let filled = 0;

  for (let row = 11; row <= 20; row += 1) {
    const courseName = String(sheet.cell(`A${row}`).value() || '').trim();
    if (!courseName || !rowIsEmpty(sheet, row)) continue;
    const matched = findCourseByNameOrAlias(courseName);
    if (!matched) continue;

    const points18 = resolveCoursePoints18(matched);
    writePoints18ToRow(sheet, row, matched.nameJa, points18);
    filled += 1;
  }

  syncPlanFromWorkbook();
  if (filled > 0) {
    setWorkbookStatus(`既存の必修 / 既入力科目から ${filled} 行を自動入力しました。`, 'ok');
  } else {
    setWorkbookStatus('自動入力できる既存行は見つかりませんでした。', 'warn');
  }
}

async function downloadWorkbook() {
  if (!state.workbook) return;
  try {
    const blob = await state.workbook.outputAsync();
    const link = document.createElement('a');
    const baseName = state.workbookFileName.replace(/\.(xlsx|xlsm|xls)$/i, '');
    const downloadName = `${baseName || '達成度評価'}_入力支援.xlsx`;
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = downloadName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setWorkbookStatus(`更新済みファイル「${downloadName}」をダウンロードしました。`, 'ok');
  } catch (error) {
    console.error(error);
    setWorkbookStatus('Excel の出力に失敗しました。', 'warn');
  }
}

function sumPoints8(points8) {
  return Object.values(points8 || {}).reduce((sum, value) => sum + Number(value || 0), 0);
}

function sumPoints18(values) {
  return (values || []).reduce((sum, value) => sum + asInt(value), 0);
}

function clonePoints18(values) {
  return Array.from({ length: 18 }, (_, index) => asInt(values?.[index] || 0));
}

function chip(text) {
  return `<span class="chip">${escapeHtml(text)}</span>`;
}

function normalize(value) {
  return String(value || '')
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s　]+/g, '')
    .replace(/[()（）]/g, '')
    .replace(/[・･]/g, '');
}

function asInt(value) {
  const num = Number(value);
  return Number.isFinite(num) ? Math.max(0, Math.round(num)) : 0;
}

function escapeHtml(value) {
  return String(value || '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
