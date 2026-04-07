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
const COLUMN_LETTERS = Array.from({ length: 18 }, (_, idx) => String.fromCharCode('B'.charCodeAt(0) + idx));

const state = {
  dataset: null,
  courses: [],
  workbook: null,
  workbookFileName: '',
  sheetName: 'M中間評価',
  selectedCourse: null,
  previewPoints18: Array(18).fill(0)
};

const els = {};

document.addEventListener('DOMContentLoaded', async () => {
  bindElements();
  wireEvents();
  renderRowOptions();
  renderPreviewGrid();
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
  els.writeRowBtn = document.getElementById('writeRowBtn');
  els.datasetInfo = document.getElementById('datasetInfo');
}

function wireEvents() {
  els.xlsxFile.addEventListener('change', onWorkbookSelected);
  els.searchInput.addEventListener('input', () => renderSearchResults(searchCourses(els.searchInput.value)));
  els.splitMode.addEventListener('change', recomputePreview);
  els.writeRowBtn.addEventListener('click', writeSelectedCourseToWorkbook);
  els.autofillRequiredBtn.addEventListener('click', autofillExistingRows);
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
  els.writeRowBtn.disabled = !state.workbook;
}

function renderSelectedCourse(course) {
  const p8 = course.points8 || {};
  const pointsHtml = GROUP_CONFIG.map(group => {
    const value = Number(p8[group.key] || 0);
    return `
      <div class="point8-item">
        <span>${escapeHtml(group.label)}</span>
        <strong>${value}</strong>
      </div>
    `;
  }).join('');

  els.selectedCourseBox.classList.remove('empty');
  els.selectedCourseBox.innerHTML = `
    <div class="result-card__code">${escapeHtml(course.code)}</div>
    <h3>${escapeHtml(course.nameJa)}</h3>
    <div class="muted">${escapeHtml(course.nameEn || '')}</div>
    <div class="points8">${pointsHtml}</div>
  `;
}

function recomputePreview() {
  if (!state.selectedCourse) {
    state.previewPoints18 = Array(18).fill(0);
    updatePreviewInputs();
    return;
  }

  if (Array.isArray(state.selectedCourse.points18) && state.selectedCourse.points18.length === 18) {
    state.previewPoints18 = state.selectedCourse.points18.map(asInt);
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
  const total = state.previewPoints18.reduce((sum, value) => sum + asInt(value), 0);
  const sourceTotal = state.selectedCourse
    ? (state.selectedCourse.total ?? sumPoints8(state.selectedCourse.points8 || {}))
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
    setWorkbookStatus(`読込完了: ${file.name} / シート: ${state.sheetName}`, 'ok');
    els.autofillRequiredBtn.disabled = false;
    els.downloadBtn.disabled = false;
    els.writeRowBtn.disabled = !state.selectedCourse;
  } catch (error) {
    console.error(error);
    state.workbook = null;
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

function writeSelectedCourseToWorkbook() {
  if (!state.workbook || !state.selectedCourse) return;
  const row = Number(els.targetRow.value || 15);
  const sheet = state.workbook.sheet(state.sheetName);
  writePreviewToRow(sheet, row, state.selectedCourse.nameJa);
  renderRowOptions();
  setWorkbookStatus(`${row}行へ「${state.selectedCourse.nameJa}」を書き込みました。`, 'ok');
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

    const points18 = Array.isArray(matched.points18) && matched.points18.length === 18
      ? matched.points18.map(asInt)
      : expandPoints8To18(matched.points8 || {}, els.splitMode.value);

    writePoints18ToRow(sheet, row, matched.nameJa, points18);
    filled += 1;
  }

  renderRowOptions();
  if (filled > 0) {
    setWorkbookStatus(`既存の必修/既入力科目から ${filled} 行を自動入力しました。`, 'ok');
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
