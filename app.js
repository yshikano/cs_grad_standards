
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

const GROUP_KEYS = GROUP_CONFIG.map(group => group.key);
const GROUP_TOTAL_REQUIREMENTS = GROUP_CONFIG.map(group => group.requiredWeights.reduce((sum, value) => sum + value, 0));
const GROUP_LABELS = GROUP_CONFIG.map(group => group.label);
const COLUMN_LABELS = GROUP_CONFIG.flatMap(group => group.cols);
const COLUMN_LETTERS = Array.from({ length: 18 }, (_, index) => String.fromCharCode('B'.charCodeAt(0) + index));
const REQUIRED_POINTS18 = GROUP_CONFIG.flatMap(group => group.requiredWeights);
const REQUIRED_TOTAL = sumNumbers(REQUIRED_POINTS18);
const CUSTOM_MODE_OPTIONS = {
  competence8: GROUP_LABELS,
  viewpoint18: COLUMN_LABELS
};
const STYLE_KEYS = [
  'bold',
  'italic',
  'underline',
  'strikethrough',
  'subscript',
  'superscript',
  'fontSize',
  'fontFamily',
  'fontGenericFamily',
  'fontScheme',
  'fontColor',
  'fill',
  'border',
  'horizontalAlignment',
  'justifyLastLine',
  'readingOrder',
  'relativeIndent',
  'shrinkToFit',
  'textRotation',
  'verticalAlignment',
  'wrapText',
  'indent',
  'numberFormat',
  'hidden',
  'locked'
];

const DEFAULT_ROW_NAMES = {
  11: '情報理工前期特別研究A',
  12: '情報理工前期特別研究B',
  13: '情報理工前期特別研究C',
  14: '情報理工前期特別研究D'
};

const DEFAULT_ROW_LABELS = {
  11: '11行: 情報理工前期特別研究A',
  12: '12行: 情報理工前期特別研究B',
  13: '13行: 情報理工前期特別研究C',
  14: '14行: 情報理工前期特別研究D'
};

const state = {
  dataset: null,
  courses: [],
  workbook: null,
  workbookFileName: '',
  sheetName: 'M中間評価',
  selectedCourse: null,
  previewPoints18: Array(18).fill(0),
  planRows: {},
  customValues8: Array(8).fill(0),
  customValues18: Array(18).fill(0)
};

const els = {};

document.addEventListener('DOMContentLoaded', async () => {
  bindElements();
  wireEvents();
  renderPreviewGrid();
  renderCustomValueGrid();
  renderSelectedCourse(null);
  renderRowOptions();
  renderSimulationSection();
  renderPlanSection();
  updateActionStates();
  await loadDataset();
});

function bindElements() {
  els.xlsxFile = document.getElementById('xlsxFile');
  els.workbookStatus = document.getElementById('workbookStatus');
  els.searchInput = document.getElementById('searchInput');
  els.searchMeta = document.getElementById('searchMeta');
  els.searchResults = document.getElementById('searchResults');

  els.customName = document.getElementById('customName');
  els.customCode = document.getElementById('customCode');
  els.customCredits = document.getElementById('customCredits');
  els.customInputMode = document.getElementById('customInputMode');
  els.customValueGrid = document.getElementById('customValueGrid');
  els.customValueMeta = document.getElementById('customValueMeta');
  els.customUseBtn = document.getElementById('customUseBtn');
  els.customClearBtn = document.getElementById('customClearBtn');

  els.selectedCourseBox = document.getElementById('selectedCourseBox');
  els.splitMode = document.getElementById('splitMode');
  els.targetRow = document.getElementById('targetRow');
  els.previewGrid = document.getElementById('previewGrid');
  els.previewTotal = document.getElementById('previewTotal');
  els.planAddBtn = document.getElementById('planAddBtn');

  els.simulationNote = document.getElementById('simulationNote');
  els.simulatedSubtotalTotal = document.getElementById('simulatedSubtotalTotal');
  els.simulatedExtraTotal = document.getElementById('simulatedExtraTotal');
  els.simulatedProjectedTotal = document.getElementById('simulatedProjectedTotal');
  els.simulatedDiffTotal = document.getElementById('simulatedDiffTotal');
  els.simulatedUnmetCount = document.getElementById('simulatedUnmetCount');
  els.simulationBody = document.getElementById('simulationBody');
  els.simulationFoot = document.getElementById('simulationFoot');

  els.applyPlanBtn = document.getElementById('applyPlanBtn');
  els.downloadBtn = document.getElementById('downloadBtn');
  els.reloadPlanBtn = document.getElementById('reloadPlanBtn');
  els.clearPlanBtn = document.getElementById('clearPlanBtn');
  els.syncStatus = document.getElementById('syncStatus');
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
  els.customInputMode.addEventListener('change', () => {
    renderCustomValueGrid();
    updateCustomValueMeta();
  });
  els.customCredits.addEventListener('input', updateCustomValueMeta);
  els.customName.addEventListener('input', updateCustomValueMeta);
  els.customCode.addEventListener('input', updateCustomValueMeta);
  els.customUseBtn.addEventListener('click', useCustomAsSelection);
  els.customClearBtn.addEventListener('click', clearCustomBuilder);

  els.splitMode.addEventListener('change', () => {
    if (state.selectedCourse) {
      recomputePreview();
    } else {
      renderSimulationSection();
    }
  });
  els.targetRow.addEventListener('change', () => {
    if (state.selectedCourse) {
      recomputePreview();
    } else {
      renderSimulationSection();
    }
  });

  els.planAddBtn.addEventListener('click', addSelectedCourseToPlan);
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

    els.searchMeta.textContent = `${state.courses.length} 件の科目データを読み込みました。`;
    const notes = (state.dataset.meta?.notes || []).map(text => escapeHtml(text)).join(' / ');
    els.datasetInfo.innerHTML = [
      `<div><strong>データ件数:</strong> ${state.dataset.meta?.count ?? state.courses.length} 件</div>`,
      `<div><strong>データソース:</strong> <a href="${escapeHtml(state.dataset.meta?.sourceUrl || '#')}" target="_blank" rel="noreferrer">公開カリキュラム・マップ</a></div>`,
      `<div><strong>メモ:</strong> ${notes || 'なし'}</div>`
    ].join('');

    renderSearchResults(searchCourses(''));

    if (state.workbook) {
      syncPlanFromWorkbook();
      setWorkbookStatus('科目データを読み直しました。現在の Excel を基にドラフトを再解釈しています。', currentStatusKindAfterPlanChange());
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
    return courses.slice(0, 16);
  }

  const scored = courses
    .map(course => ({ course, score: scoreCourse(course, term) }))
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score || a.course.code.localeCompare(b.course.code, 'ja'));

  return scored.slice(0, 30).map(item => item.course);
}

function scoreCourse(course, term) {
  const code = normalize(course.code);
  const ja = normalize(course.nameJa);
  const en = normalize(course.nameEn);
  const aliases = (course.aliases || []).map(normalize);

  if (code === term) return 1200;
  if (ja === term || en === term || aliases.includes(term)) return 1000;
  if (code.startsWith(term)) return 900;
  if (ja.includes(term)) return 800;
  if (en.includes(term)) return 760;
  if (aliases.some(alias => alias.includes(term))) return 700;
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

    const sourceKindLabel = course.sourceKind === 'exact8'
      ? '8コンピテンス確定'
      : (course.sourceKind === 'sequence' ? '非ゼロ列を自動補完' : 'カスタム');
    const metaChips = [
      chip(course.code || 'コードなし'),
      chip(`単位 ${course.credits ?? '—'}`),
      chip(`合計 ${course.total ?? '—'} 点`),
      chip(sourceKindLabel)
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

function renderCustomValueGrid() {
  const mode = els.customInputMode.value;
  const labels = CUSTOM_MODE_OPTIONS[mode] || GROUP_LABELS;
  const values = mode === 'viewpoint18' ? state.customValues18 : state.customValues8;

  els.customValueGrid.innerHTML = '';
  labels.forEach((label, index) => {
    const wrapper = document.createElement('label');
    wrapper.className = 'preview-cell preview-cell--compact';
    wrapper.innerHTML = `
      <span class="preview-cell__label">${mode === 'viewpoint18' ? `${COLUMN_LETTERS[index]}列 / ` : ''}${escapeHtml(label)}</span>
      <input type="number" min="0" step="1" value="${asInt(values[index])}">
    `;
    const input = wrapper.querySelector('input');
    input.addEventListener('input', () => {
      if (mode === 'viewpoint18') {
        state.customValues18[index] = asInt(input.value);
      } else {
        state.customValues8[index] = asInt(input.value);
      }
      updateCustomValueMeta();
    });
    els.customValueGrid.appendChild(wrapper);
  });

  updateCustomValueMeta();
}

function updateCustomValueMeta() {
  const mode = els.customInputMode.value;
  const values = mode === 'viewpoint18' ? state.customValues18 : state.customValues8;
  const inputTotal = sumNumbers(values);
  const credits = asNumber(els.customCredits.value);
  const targetTotal = credits > 0 ? Math.round(credits * 100) : 0;

  let message = '元のコンピテンス / 観点を入力すると、1単位 100 点へ正規化した上で不足差分を見ながら自動配点します。';
  let kind = 'muted';

  if (!els.customName.value.trim() && inputTotal === 0 && targetTotal === 0) {
    kind = 'muted';
  } else if (credits <= 0) {
    message = '単位数を入れてください。外部科目は 1 単位 = 100 点で扱います。';
    kind = 'warn';
  } else if (inputTotal <= 0) {
    message = `${credits} 単位なので最終合計は ${targetTotal} 点になります。元の配点を 1 つ以上入れてください。`;
    kind = 'warn';
  } else if (inputTotal === targetTotal) {
    message = `入力合計 ${inputTotal} 点。${targetTotal} 点のまま使います。smart 配点では不足差分を優先しつつ、この入力分布を尊重します。`;
    kind = 'ok';
  } else {
    message = `入力合計 ${inputTotal} 点を ${targetTotal} 点へ正規化して使います。smart 配点では不足差分を優先しつつ、この入力分布を尊重します。`;
    kind = 'pending';
  }

  els.customValueMeta.textContent = message;
  els.customValueMeta.className = `status status--${kind}`;
}

function clearCustomBuilder() {
  els.customName.value = '';
  els.customCode.value = '';
  els.customCredits.value = '';
  state.customValues8 = Array(8).fill(0);
  state.customValues18 = Array(18).fill(0);
  renderCustomValueGrid();
  updateCustomValueMeta();
}

function buildCustomCourseFromForm() {
  const name = els.customName.value.trim();
  const code = els.customCode.value.trim();
  const credits = asNumber(els.customCredits.value);
  const total = credits > 0 ? Math.round(credits * 100) : 0;
  const mode = els.customInputMode.value;
  const rawValues = (mode === 'viewpoint18' ? state.customValues18 : state.customValues8).map(asInt);
  const inputTotal = sumNumbers(rawValues);

  if (!name) {
    throw new Error('科目名を入力してください。');
  }
  if (!(credits > 0)) {
    throw new Error('単位数を入力してください。');
  }
  if (!(inputTotal > 0)) {
    throw new Error('元の配点を 1 つ以上入力してください。');
  }

  const course = {
    code,
    nameJa: name,
    nameEn: '',
    credits,
    total,
    isExternal: true,
    sourceKind: mode === 'viewpoint18' ? 'external18' : 'external8',
    aliases: []
  };

  if (mode === 'viewpoint18') {
    course.points18 = clonePoints18(rawValues);
  } else {
    course.points8 = arrayToPoints8(rawValues);
  }

  return course;
}

function useCustomAsSelection() {
  try {
    const course = buildCustomCourseFromForm();
    selectCourse(course);
    setWorkbookStatus(`「${course.nameJa}」を選択中の科目へセットしました。`, currentStatusKindAfterPlanChange());
  } catch (error) {
    setWorkbookStatus(error.message, 'warn');
  }
}

function selectCourse(course) {
  state.selectedCourse = cloneCourseSource(course);
  renderSelectedCourse(state.selectedCourse);
  recomputePreview();
  updateActionStates();
}

function renderSelectedCourse(course) {
  if (!course) {
    els.selectedCourseBox.classList.add('empty');
    els.selectedCourseBox.innerHTML = 'まだ科目が選択されていません。';
    return;
  }

  const subline = [
    course.code || 'コード未設定',
    course.credits ? `${course.credits} 単位` : '',
    course.isExternal ? '他研究群科目' : 'システム情報工学研究群科目'
  ].filter(Boolean).join(' / ');

  let pointsHtml = '';
  if (Array.isArray(course.points18) && course.points18.length === 18) {
    pointsHtml = `
      <div class="point8-item point8-item--wide">
        <span>元データ</span>
        <strong>18観点入力 / 合計 ${sumPoints18(course.points18)} 点</strong>
      </div>
    `;
  } else if (course.points8 && Object.keys(course.points8).length > 0) {
    pointsHtml = GROUP_CONFIG.map(group => `
      <div class="point8-item">
        <span>${escapeHtml(group.label)}</span>
        <strong>${asInt(course.points8[group.key])}</strong>
      </div>
    `).join('');
  } else if (Array.isArray(course.rawCompetenceValues) && course.rawCompetenceValues.length > 0) {
    pointsHtml = `
      <div class="point8-item point8-item--wide">
        <span>公開 PDF から読める非ゼロ配点列</span>
        <strong>${course.rawCompetenceValues.map(asInt).join(' / ')}</strong>
        <div class="help-text">不足差分を見ながら 8 コンピテンスへ補完し、その後 18 観点へ自動展開します。</div>
      </div>
    `;
  } else {
    pointsHtml = `
      <div class="point8-item point8-item--wide">
        <span>元データ</span>
        <strong>配点情報なし</strong>
      </div>
    `;
  }

  const badges = [
    chip(course.sourceKind === 'sequence' ? '自動補完対象' : '元配点あり'),
    chip(`最終合計 ${asInt(resolveCourseTotal(course))} 点`)
  ].join('');

  els.selectedCourseBox.classList.remove('empty');
  els.selectedCourseBox.innerHTML = `
    <div class="result-card__code">${escapeHtml(course.code || 'カスタム')}</div>
    <h3>${escapeHtml(course.nameJa || '名称未設定')}</h3>
    <div class="muted">${escapeHtml(subline)}</div>
    <div class="selected-course__source">${badges}</div>
    <div class="points8">${pointsHtml}</div>
  `;
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
      renderSimulationSection();
    });
    els.previewGrid.appendChild(wrapper);
  });

  updatePreviewInputs();
}

function updatePreviewInputs() {
  const inputs = [...els.previewGrid.querySelectorAll('input')];
  inputs.forEach((input, index) => {
    input.value = String(asInt(state.previewPoints18[index]));
  });
  updatePreviewTotal();
  renderSimulationSection();
}

function updatePreviewTotal() {
  const previewTotal = sumPoints18(state.previewPoints18);
  const sourceTotal = state.selectedCourse ? resolveCourseTotal(state.selectedCourse) : 0;
  const diff = previewTotal - sourceTotal;
  const suffix = diff === 0 ? '' : ` / 元データとの差 ${formatSigned(diff)}`;
  els.previewTotal.textContent = `合計 ${previewTotal} 点${suffix}`;
}

function recomputePreview() {
  if (!state.selectedCourse) {
    state.previewPoints18 = Array(18).fill(0);
    updatePreviewInputs();
    return;
  }

  const targetRow = Number(els.targetRow.value || 11);
  const positiveDeficits18 = getPositiveDeficits18(state.planRows, targetRow);
  state.previewPoints18 = resolveCoursePoints18(state.selectedCourse, {
    mode: els.splitMode.value,
    positiveDeficits18
  });

  updatePreviewInputs();
}

function resolveCoursePoints18(course, options = {}) {
  const mode = options.mode || 'smart';
  const positiveDeficits18 = clonePoints18(options.positiveDeficits18 || Array(18).fill(0));
  const total = resolveCourseTotal(course);

  if (Array.isArray(course?.points18) && course.points18.length === 18) {
    return resolveFromOriginal18(course.points18, total, mode, positiveDeficits18);
  }

  const points8 = resolveCoursePoints8(course, total, mode, positiveDeficits18);
  return expandPoints8To18(points8, mode, positiveDeficits18);
}

function resolveFromOriginal18(original18, total, mode, positiveDeficits18) {
  const base = clonePoints18(original18);
  const allowedMask = base.some(value => asInt(value) > 0)
    ? base.map(value => asInt(value) > 0)
    : Array(18).fill(true);

  if (mode === 'even') {
    return allocateByWeights(total, allowedMask.map(flag => flag ? 1 : 0));
  }

  if (mode === 'required') {
    return allocateByWeights(total, REQUIRED_POINTS18.map((value, index) => allowedMask[index] ? value : 0));
  }

  return allocateComposite(total, base, positiveDeficits18, REQUIRED_POINTS18, allowedMask);
}

function resolveCoursePoints8(course, total, mode, positiveDeficits18) {
  if (course?.points8 && Object.keys(course.points8).length > 0) {
    const raw = GROUP_KEYS.map(key => asInt(course.points8[key]));
    const normalized = normalizeVectorToTotal(raw, total);
    return arrayToPoints8(normalized);
  }

  if (Array.isArray(course?.rawCompetenceValues) && course.rawCompetenceValues.length > 0) {
    const raw = course.rawCompetenceValues.map(asInt);
    const normalized = normalizeVectorToTotal(raw, total);
    const placed = placeSequenceIntoGroups(normalized, positiveDeficits18, mode);
    return arrayToPoints8(placed);
  }

  return zeroPoints8();
}

function placeSequenceIntoGroups(sequence, positiveDeficits18, mode) {
  const k = sequence.length;
  if (!k) return Array(8).fill(0);
  if (k >= 8) return normalizeVectorToTotal(sequence.slice(0, 8), sumNumbers(sequence.slice(0, 8)));

  const combos = choosePositions(8, k);
  const groupNeeds = getGroupNeedTotals(positiveDeficits18);
  const scoringWeights = mode === 'smart'
    ? (sumNumbers(groupNeeds) > 0 ? groupNeeds : GROUP_TOTAL_REQUIREMENTS)
    : (mode === 'required' ? GROUP_TOTAL_REQUIREMENTS : Array(8).fill(1));

  let bestScore = Number.NEGATIVE_INFINITY;
  let best = null;

  combos.forEach(positions => {
    const arr = Array(8).fill(0);
    positions.forEach((position, index) => {
      arr[position] = sequence[index];
    });

    const fitScore = arr.reduce((sum, value, index) => sum + (value * scoringWeights[index]), 0);
    const tieBreaker = positions.reduce((sum, value) => sum - (value * 0.001), 0);
    const score = fitScore + tieBreaker;

    if (score > bestScore) {
      bestScore = score;
      best = arr;
    }
  });

  return best || Array(8).fill(0);
}

function expandPoints8To18(points8, mode, positiveDeficits18) {
  const output = [];
  let offset = 0;

  GROUP_CONFIG.forEach(group => {
    const total = asInt(points8[group.key]);
    const span = group.cols.length;
    const needSlice = positiveDeficits18.slice(offset, offset + span);

    let allocated;
    if (mode === 'even') {
      allocated = allocateByWeights(total, Array(span).fill(1));
    } else if (mode === 'required') {
      allocated = allocateByWeights(total, group.requiredWeights);
    } else {
      allocated = allocateComposite(total, group.requiredWeights, needSlice, group.requiredWeights, Array(span).fill(true));
    }

    output.push(...allocated);
    offset += span;
  });

  return clonePoints18(output);
}

function allocateComposite(total, baseWeights, needWeights, fallbackWeights, allowedMask) {
  const safeTotal = asInt(total);
  const length = Math.max(
    Array.isArray(baseWeights) ? baseWeights.length : 0,
    Array.isArray(needWeights) ? needWeights.length : 0,
    Array.isArray(fallbackWeights) ? fallbackWeights.length : 0,
    Array.isArray(allowedMask) ? allowedMask.length : 0
  );

  const mask = Array.from({ length }, (_, index) => {
    if (!Array.isArray(allowedMask) || !allowedMask.length) return true;
    return Boolean(allowedMask[index]);
  });

  const base = normalizeMaskedWeights(baseWeights, mask);
  const need = normalizeMaskedWeights((needWeights || []).map(value => Math.max(0, asIntSigned(value))), mask);
  const fallback = normalizeMaskedWeights(fallbackWeights, mask);

  let weights;
  if (sumNumbers(base) > 0 && sumNumbers(need) > 0) {
    weights = base.map((value, index) => (value * 0.4) + (need[index] * 0.6));
  } else if (sumNumbers(need) > 0) {
    weights = need;
  } else if (sumNumbers(base) > 0) {
    weights = base;
  } else if (sumNumbers(fallback) > 0) {
    weights = fallback;
  } else {
    weights = mask.map(flag => flag ? 1 : 0);
  }

  return allocateByWeights(safeTotal, weights);
}

function normalizeMaskedWeights(values, mask) {
  const arr = Array.from({ length: mask.length }, (_, index) => {
    if (!mask[index]) return 0;
    return Math.max(0, Number(values?.[index] || 0));
  });

  const total = arr.reduce((sum, value) => sum + value, 0);
  if (total <= 0) return arr;
  return arr.map(value => value / total);
}

function normalizeVectorToTotal(values, total) {
  const safeValues = (values || []).map(asInt);
  const safeTotal = asInt(total);
  const current = sumNumbers(safeValues);
  if (safeTotal <= 0) return safeValues.map(() => 0);
  if (current === safeTotal) return safeValues.slice();
  if (current <= 0) return safeValues.map(() => 0);
  return allocateByWeights(safeTotal, safeValues);
}

function allocateByWeights(total, weights) {
  const safeTotal = asInt(total);
  const safeWeights = (weights || []).map(value => Math.max(0, Number(value || 0)));

  if (!safeWeights.length) return [];
  if (safeTotal <= 0) return safeWeights.map(() => 0);

  let weightSum = safeWeights.reduce((sum, value) => sum + value, 0);
  let workingWeights = safeWeights.slice();
  if (weightSum <= 0) {
    workingWeights = safeWeights.map(() => 1);
    weightSum = workingWeights.length;
  }

  const raw = workingWeights.map(weight => (safeTotal * weight) / weightSum);
  const floors = raw.map(value => Math.floor(value));
  let remainder = safeTotal - floors.reduce((sum, value) => sum + value, 0);

  const order = raw
    .map((value, index) => ({ index, frac: value - Math.floor(value) }))
    .sort((a, b) => b.frac - a.frac || a.index - b.index);

  const allocated = floors.slice();
  if (!order.length) return allocated;

  for (let i = 0; i < remainder; i += 1) {
    allocated[order[i % order.length].index] += 1;
  }

  return allocated;
}

function choosePositions(length, pick) {
  const results = [];
  const walk = (start, prefix) => {
    if (prefix.length === pick) {
      results.push(prefix.slice());
      return;
    }
    for (let i = start; i < length; i += 1) {
      prefix.push(i);
      walk(i + 1, prefix);
      prefix.pop();
    }
  };
  walk(0, []);
  return results;
}

function getGroupNeedTotals(positiveDeficits18) {
  let offset = 0;
  return GROUP_CONFIG.map(group => {
    const slice = positiveDeficits18.slice(offset, offset + group.cols.length);
    offset += group.cols.length;
    return sumNumbers(slice);
  });
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

    if (isPlanSyncedWithWorkbook()) {
      setWorkbookStatus(`読込完了: ${file.name} / シート: ${state.sheetName} / ${imported} 行のドラフトが Excel と同期しています。`, 'ok');
    } else {
      setWorkbookStatus(`読込完了: ${file.name} / シート: ${state.sheetName} / ${imported} 行のドラフトを作成しました。まだ Excel へは反映していません。`, 'pending');
    }
  } catch (error) {
    console.error(error);
    state.workbook = null;
    state.workbookFileName = '';
    state.planRows = {};
    renderRowOptions();
    renderSimulationSection();
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

function detectSheetLayout(sheet) {
  const scanMax = 500;
  const labelRows = [];
  for (let row = 1; row <= scanMax; row += 1) {
    const label = normalize(sheet.cell(`A${row}`).value());
    if (!label) continue;
    if (['小計', '授業科目以外', '合計', '不足分'].includes(label)) {
      labelRows.push({ row, label });
    }
  }

  const courseSubtotal = labelRows.find(item => item.label === '小計' && item.row >= 11);
  const extraHeader = labelRows.find(item => item.label === '授業科目以外' && item.row > (courseSubtotal?.row || 11));
  const extraSubtotal = labelRows.find(item => item.label === '小計' && item.row > (extraHeader?.row || Number.MAX_SAFE_INTEGER));
  const totalRow = labelRows.find(item => item.label === '合計' && item.row > (extraSubtotal?.row || Number.MAX_SAFE_INTEGER));
  const deficitRow = labelRows.find(item => item.label === '不足分' && item.row > (totalRow?.row || Number.MAX_SAFE_INTEGER));

  const courseSubtotalRow = courseSubtotal?.row || 21;
  const extraHeaderRow = extraHeader?.row || 23;
  const extraSubtotalRow = extraSubtotal?.row || 27;
  const total = totalRow?.row || 29;
  const deficit = deficitRow?.row || 30;
  const extraStartRow = extraHeaderRow + 1;
  const extraEndRow = extraSubtotalRow - 1;

  return {
    courseStartRow: 11,
    courseEndRow: courseSubtotalRow - 1,
    courseSubtotalRow,
    blankAfterCourseSubtotalRow: courseSubtotalRow + 1,
    extraHeaderRow,
    extraStartRow,
    extraEndRow,
    extraSubtotalRow,
    blankAfterExtraSubtotalRow: extraSubtotalRow + 1,
    totalRow: total,
    deficitRow: deficit
  };
}

function renderRowOptions(selectedValue = null) {
  const maxRow = getSelectableMaxRow();
  const currentValue = Number((selectedValue ?? els.targetRow.value) || 11);
  const safeSelected = Math.min(Math.max(11, currentValue), maxRow);

  els.targetRow.innerHTML = '';
  for (let row = 11; row <= maxRow; row += 1) {
    const option = document.createElement('option');
    option.value = String(row);
    option.textContent = getRowLabel(row);
    els.targetRow.appendChild(option);
  }
  els.targetRow.value = String(safeSelected);
}

function getSelectableMaxRow() {
  let maxRow = 20;
  const highestPlanned = getHighestPlannedRow();
  if (highestPlanned > 0) {
    maxRow = Math.max(maxRow, highestPlanned + 1);
  }
  if (state.workbook) {
    try {
      const sheet = state.workbook.sheet(state.sheetName);
      const layout = detectSheetLayout(sheet);
      maxRow = Math.max(maxRow, layout.courseEndRow + 1);
    } catch (error) {
      console.warn(error);
    }
  }
  return maxRow;
}

function getRowLabel(row) {
  const planned = state.planRows[row];
  if (planned) {
    return `${row}行: ${displayCourseName(planned)}`;
  }

  if (state.workbook) {
    try {
      const sheet = state.workbook.sheet(state.sheetName);
      const layout = detectSheetLayout(sheet);
      if (row <= layout.courseEndRow) {
        const courseName = String(sheet.cell(`A${row}`).value() || '').trim();
        const points18 = COLUMN_LETTERS.map(letter => asInt(sheet.cell(`${letter}${row}`).value()));
        if (courseName || points18.some(value => value > 0)) {
          return `${row}行: ${courseName || '入力済み科目'}`;
        }
      }
    } catch (error) {
      console.warn(error);
    }
  }

  return DEFAULT_ROW_LABELS[row] || `${row}行: 空き行`;
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

function getDefaultClearName(row) {
  return DEFAULT_ROW_NAMES[row] || '';
}

function writePoints18ToRow(sheet, row, courseName, points18) {
  sheet.cell(`A${row}`).value(courseName || '');
  COLUMN_LETTERS.forEach((letter, index) => {
    sheet.cell(`${letter}${row}`).value(asInt(points18[index]));
  });
  sheet.cell(`T${row}`).formula(`SUM(B${row}:S${row})`);
}

function clearCourseRow(sheet, row) {
  sheet.cell(`A${row}`).value(getDefaultClearName(row));
  COLUMN_LETTERS.forEach(letter => {
    sheet.cell(`${letter}${row}`).value('');
  });
  sheet.cell(`T${row}`).formula(`SUM(B${row}:S${row})`);
}

function makePlanEntryFromSelection(row) {
  const course = cloneCourseSource(state.selectedCourse);
  return {
    row,
    code: course.code || '',
    nameJa: course.nameJa || '',
    nameEn: course.nameEn || '',
    credits: Number(course.credits || 0),
    sourceTotal: resolveCourseTotal(course),
    points18: clonePoints18(state.previewPoints18),
    sourceCourse: course,
    isExternal: Boolean(course.isExternal),
    sourceKind: course.sourceKind || 'manual'
  };
}

function addSelectedCourseToPlan() {
  if (!state.selectedCourse) return;
  const row = Number(els.targetRow.value || 11);
  state.planRows[row] = makePlanEntryFromSelection(row);
  renderRowOptions();
  selectNextAvailableRow(row);
  state.selectedCourse = null;
  state.previewPoints18 = Array(18).fill(0);
  renderSelectedCourse(null);
  updatePreviewInputs();
  renderPlanSection();
  setWorkbookStatus(`${row}行を履修計画ドラフトに追加 / 更新しました。まだ Excel へは反映していません。`, currentStatusKindAfterPlanChange());
}

function selectNextAvailableRow(currentRow) {
  const maxRow = getSelectableMaxRow() + 1;
  const candidates = [];
  for (let row = currentRow + 1; row <= maxRow; row += 1) {
    candidates.push(row);
  }
  for (let row = 11; row <= currentRow; row += 1) {
    candidates.push(row);
  }
  const nextEmpty = candidates.find(row => !state.planRows[row]) || maxRow;
  renderRowOptions(nextEmpty);
  renderSimulationSection();
}

function clonePlanRows(source = state.planRows) {
  const cloned = {};
  Object.entries(source).forEach(([row, entry]) => {
    cloned[row] = {
      ...entry,
      row: Number(row),
      points18: clonePoints18(entry.points18),
      sourceCourse: entry.sourceCourse ? cloneCourseSource(entry.sourceCourse) : null
    };
  });
  return cloned;
}

function getPlanEntries(planRows = state.planRows) {
  return Object.values(planRows)
    .map(entry => ({ ...entry, row: Number(entry.row) }))
    .sort((a, b) => a.row - b.row);
}

function getHighestPlannedRow(planRows = state.planRows) {
  const entries = getPlanEntries(planRows);
  return entries.length ? Math.max(...entries.map(entry => entry.row)) : 0;
}

function getPlanTotals18(planRows = state.planRows) {
  const totals = Array(18).fill(0);
  getPlanEntries(planRows).forEach(entry => {
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
    const layout = detectSheetLayout(sheet);
    for (let row = layout.extraStartRow; row <= layout.extraEndRow; row += 1) {
      COLUMN_LETTERS.forEach((letter, index) => {
        totals[index] += asInt(sheet.cell(`${letter}${row}`).value());
      });
    }
  } catch (error) {
    console.warn(error);
  }

  return totals;
}

function getMetrics(planRows = state.planRows) {
  const subtotal18 = getPlanTotals18(planRows);
  const extra18 = getExtraTotals18();
  const projected18 = subtotal18.map((value, index) => value + extra18[index]);
  const diff18 = projected18.map((value, index) => value - REQUIRED_POINTS18[index]);
  const unmetCount = diff18.filter(value => value < 0).length;

  return {
    subtotal18,
    extra18,
    projected18,
    diff18,
    subtotalTotal: sumPoints18(subtotal18),
    extraTotal: sumPoints18(extra18),
    projectedTotal: sumPoints18(projected18),
    diffTotal: sumPoints18(projected18) - REQUIRED_TOTAL,
    unmetCount
  };
}

function getPositiveDeficits18(planRows = state.planRows, excludingRow = null) {
  const working = clonePlanRows(planRows);
  if (excludingRow !== null) {
    delete working[excludingRow];
  }
  const metrics = getMetrics(working);
  return metrics.diff18.map(value => value < 0 ? Math.abs(value) : 0);
}

function buildPlanRowsWithPreview() {
  if (!state.selectedCourse) {
    return clonePlanRows(state.planRows);
  }

  const row = Number(els.targetRow.value || 11);
  const simulated = clonePlanRows(state.planRows);
  simulated[row] = makePlanEntryFromSelection(row);
  return simulated;
}

function renderSimulationSection() {
  const row = Number(els.targetRow?.value || 11);
  const replacing = state.planRows[row];
  const simulatedRows = buildPlanRowsWithPreview();
  const metrics = getMetrics(simulatedRows);

  if (!state.selectedCourse) {
    els.simulationNote.textContent = '科目を選ぶと、追加後の授業科目小計・見込み合計・必要点との差分をここで確認できます。';
    els.simulationNote.className = 'status status--muted';
  } else if (replacing) {
    els.simulationNote.textContent = `${row}行は現在「${displayCourseName(replacing)}」です。このシミュレーションでは、その内容を選択中の科目で置き換えます。smart 配点では不足差分を見ながら配点を再計算しています。`;
    els.simulationNote.className = 'status status--pending';
  } else {
    els.simulationNote.textContent = `${row}行へ選択中の科目を追加した場合のシミュレーションです。smart 配点では不足差分を見ながら配点を自動調整しています。`;
    els.simulationNote.className = 'status status--muted';
  }

  els.simulatedSubtotalTotal.textContent = String(metrics.subtotalTotal);
  els.simulatedExtraTotal.textContent = String(metrics.extraTotal);
  els.simulatedProjectedTotal.textContent = String(metrics.projectedTotal);
  els.simulatedDiffTotal.textContent = formatSigned(metrics.diffTotal);
  els.simulatedDiffTotal.className = metrics.diffTotal >= 0 ? 'value value--ok' : 'value value--deficit';
  els.simulatedUnmetCount.textContent = `${metrics.unmetCount} / 18`;

  els.simulationBody.innerHTML = COLUMN_LABELS.map((label, index) => {
    const required = REQUIRED_POINTS18[index];
    const subtotal = metrics.subtotal18[index];
    const projected = metrics.projected18[index];
    const diff = metrics.diff18[index];
    const status = diff >= 0
      ? '<span class="badge badge--ok">OK</span>'
      : `<span class="badge badge--warn">不足 ${Math.abs(diff)}</span>`;

    return `
      <tr>
        <th>${escapeHtml(`${COLUMN_LETTERS[index]} / ${label}`)}</th>
        <td class="num">${required}</td>
        <td class="num">${subtotal}</td>
        <td class="num">${projected}</td>
        <td class="num ${diff >= 0 ? 'num--ok' : 'num--deficit'}">${formatSigned(diff)}</td>
        <td>${status}</td>
      </tr>
    `;
  }).join('');

  els.simulationFoot.innerHTML = `
    <tr>
      <th>合計</th>
      <td class="num">${REQUIRED_TOTAL}</td>
      <td class="num">${metrics.subtotalTotal}</td>
      <td class="num">${metrics.projectedTotal}</td>
      <td class="num ${metrics.diffTotal >= 0 ? 'num--ok' : 'num--deficit'}">${formatSigned(metrics.diffTotal)}</td>
      <td>${metrics.diffTotal >= 0 ? '<span class="badge badge--ok">総点 OK</span>' : `<span class="badge badge--warn">総点不足 ${Math.abs(metrics.diffTotal)}</span>`}</td>
    </tr>
  `;
}

function renderPlanSection() {
  renderSyncStatus();
  renderPlanList();
  renderTotalsTable();
  updateActionStates();
}

function renderSyncStatus() {
  if (!state.workbook) {
    els.syncStatus.innerHTML = '<span class="badge">Excel 未読込</span><span class="muted"> まだ一括反映先の Excel がありません。</span>';
    return;
  }

  if (isPlanSyncedWithWorkbook()) {
    els.syncStatus.innerHTML = '<span class="badge badge--ok">同期済み</span><span class="muted"> UI 上の履修計画ドラフトと Excel の授業科目欄が一致しています。</span>';
    return;
  }

  els.syncStatus.innerHTML = '<span class="badge badge--pending">未反映</span><span class="muted"> UI 上のドラフトに未反映の変更があります。最後に「履修計画を Excel に一括反映」を押してください。</span>';
}

function renderPlanList() {
  const entries = getPlanEntries();

  if (!entries.length) {
    els.planList.innerHTML = '<div class="status status--muted">まだ履修計画ドラフトはありません。科目を選んで「履修計画ドラフトに追加 / 更新」を押してください。</div>';
  } else {
    els.planList.innerHTML = entries.map(entry => {
      const total = sumPoints18(entry.points18);
      const tags = [
        entry.code ? chip(entry.code) : '',
        entry.isExternal ? chip('他研究群') : chip('シス情科目'),
        chip(`${entry.credits ? `${entry.credits} 単位 / ` : ''}${total} 点`)
      ].join('');
      return `
        <article class="plan-card">
          <div class="plan-card__top">
            <div>
              <div class="plan-card__row">${entry.row}行</div>
              <h3>${escapeHtml(displayCourseName(entry))}</h3>
              <p>${escapeHtml(entry.nameEn || '')}</p>
              <div class="plan-card__meta">${tags}</div>
            </div>
            <div class="plan-card__total">${total} 点</div>
          </div>
          <div class="actions-row actions-row--compact">
            <button class="button button--small" type="button" data-action="up" data-row="${entry.row}">上へ</button>
            <button class="button button--small" type="button" data-action="down" data-row="${entry.row}">下へ</button>
            <button class="button button--small" type="button" data-action="edit" data-row="${entry.row}">編集</button>
            <button class="button button--small button--danger" type="button" data-action="remove" data-row="${entry.row}">削除</button>
          </div>
        </article>
      `;
    }).join('');

    els.planList.querySelectorAll('button[data-action="remove"]').forEach(button => {
      button.addEventListener('click', () => removePlanRow(Number(button.dataset.row)));
    });
    els.planList.querySelectorAll('button[data-action="edit"]').forEach(button => {
      button.addEventListener('click', () => editPlanRow(Number(button.dataset.row)));
    });
    els.planList.querySelectorAll('button[data-action="up"]').forEach(button => {
      button.addEventListener('click', () => movePlanRow(Number(button.dataset.row), -1));
    });
    els.planList.querySelectorAll('button[data-action="down"]').forEach(button => {
      button.addEventListener('click', () => movePlanRow(Number(button.dataset.row), +1));
    });
  }

  const metrics = getMetrics();
  els.planCourseCount.textContent = String(entries.length);
  els.planCredits.textContent = String(entries.reduce((sum, entry) => sum + Number(entry.credits || 0), 0));
  els.planPointsTotal.textContent = String(metrics.subtotalTotal);
  els.extraPointsTotal.textContent = String(metrics.extraTotal);
  els.projectedTotal.textContent = String(metrics.projectedTotal);
  els.unmetCount.textContent = `${metrics.unmetCount} / 18`;
}

function movePlanRow(row, delta) {
  const targetRow = row + delta;
  if (targetRow < 11) return;

  const entries = clonePlanRows();
  const current = entries[row];
  if (!current) return;

  const swap = entries[targetRow];
  if (swap) {
    entries[row] = { ...swap, row };
  } else {
    delete entries[row];
  }
  entries[targetRow] = { ...current, row: targetRow };

  state.planRows = entries;
  renderRowOptions(targetRow);
  if (state.selectedCourse) {
    recomputePreview();
  } else {
    renderSimulationSection();
  }
  renderPlanSection();
  setWorkbookStatus(`履修計画ドラフトの並びを更新しました。まだ Excel へは反映していません。`, currentStatusKindAfterPlanChange());
}

function renderTotalsTable() {
  const metrics = getMetrics();

  els.totalsBody.innerHTML = COLUMN_LABELS.map((label, index) => {
    const required = REQUIRED_POINTS18[index];
    const subtotal = metrics.subtotal18[index];
    const extra = metrics.extra18[index];
    const projected = metrics.projected18[index];
    const diff = metrics.diff18[index];
    const status = diff >= 0
      ? '<span class="badge badge--ok">OK</span>'
      : `<span class="badge badge--warn">不足 ${Math.abs(diff)}</span>`;

    return `
      <tr>
        <th>${escapeHtml(`${COLUMN_LETTERS[index]} / ${label}`)}</th>
        <td class="num">${required}</td>
        <td class="num">${subtotal}</td>
        <td class="num">${extra}</td>
        <td class="num">${projected}</td>
        <td class="num ${diff >= 0 ? 'num--ok' : 'num--deficit'}">${formatSigned(diff)}</td>
        <td>${status}</td>
      </tr>
    `;
  }).join('');

  els.totalsFoot.innerHTML = `
    <tr>
      <th>合計</th>
      <td class="num">${REQUIRED_TOTAL}</td>
      <td class="num">${metrics.subtotalTotal}</td>
      <td class="num">${metrics.extraTotal}</td>
      <td class="num">${metrics.projectedTotal}</td>
      <td class="num ${metrics.diffTotal >= 0 ? 'num--ok' : 'num--deficit'}">${formatSigned(metrics.diffTotal)}</td>
      <td>${metrics.diffTotal >= 0 ? '<span class="badge badge--ok">総点 OK</span>' : `<span class="badge badge--warn">総点不足 ${Math.abs(metrics.diffTotal)}</span>`}</td>
    </tr>
  `;
}

function updateActionStates() {
  const hasWorkbook = Boolean(state.workbook);
  const hasSelection = Boolean(state.selectedCourse);
  const hasPlan = getPlanEntries().length > 0;
  const synced = hasWorkbook && isPlanSyncedWithWorkbook();

  els.planAddBtn.disabled = !hasSelection;
  els.applyPlanBtn.disabled = !hasWorkbook || synced;
  els.downloadBtn.disabled = !hasWorkbook || !synced;
  els.reloadPlanBtn.disabled = !hasWorkbook;
  els.clearPlanBtn.disabled = !hasPlan;
}

function expectedCourseEndRow(planRows = state.planRows) {
  return Math.max(20, getHighestPlannedRow(planRows) || 20);
}

function isPlanSyncedWithWorkbook() {
  if (!state.workbook) return false;

  try {
    const sheet = state.workbook.sheet(state.sheetName);
    const layout = detectSheetLayout(sheet);
    const expectedEnd = expectedCourseEndRow();

    if (layout.courseEndRow !== expectedEnd) {
      return false;
    }

    for (let row = 11; row <= expectedEnd; row += 1) {
      const entry = state.planRows[row];
      const expectedName = entry ? (entry.nameJa || entry.nameEn || entry.code || '') : getDefaultClearName(row);
      const expectedPoints18 = entry ? clonePoints18(entry.points18) : Array(18).fill(0);
      const actualName = String(sheet.cell(`A${row}`).value() || '').trim();
      const actualPoints18 = COLUMN_LETTERS.map(letter => asInt(sheet.cell(`${letter}${row}`).value()));

      if (!courseNamesEquivalent(actualName, expectedName)) {
        return false;
      }

      for (let index = 0; index < expectedPoints18.length; index += 1) {
        if (actualPoints18[index] !== expectedPoints18[index]) {
          return false;
        }
      }
    }

    return true;
  } catch (error) {
    console.warn(error);
    return false;
  }
}

function courseNamesEquivalent(a, b) {
  const left = String(a || '').trim();
  const right = String(b || '').trim();
  if (!left && !right) return true;
  if (normalize(left) === normalize(right)) return true;

  const leftCourse = findCourseByNameOrAlias(left);
  const rightCourse = findCourseByNameOrAlias(right);
  if (leftCourse && rightCourse) {
    return normalize(leftCourse.code || leftCourse.nameJa) === normalize(rightCourse.code || rightCourse.nameJa);
  }

  return false;
}

function currentStatusKindAfterPlanChange() {
  if (!state.workbook) return 'muted';
  return isPlanSyncedWithWorkbook() ? 'ok' : 'pending';
}

function syncPlanFromWorkbook() {
  if (!state.workbook) {
    state.planRows = {};
    renderSimulationSection();
    renderPlanSection();
    updateActionStates();
    return 0;
  }

  const sheet = state.workbook.sheet(state.sheetName);
  const layout = detectSheetLayout(sheet);
  const imported = {};

  for (let row = 11; row <= layout.courseEndRow; row += 1) {
    const name = String(sheet.cell(`A${row}`).value() || '').trim();
    const points18 = COLUMN_LETTERS.map(letter => asInt(sheet.cell(`${letter}${row}`).value()));
    const hasAnyPoints = points18.some(value => value > 0);
    if (!hasAnyPoints) {
      continue;
    }

    const matched = findCourseByNameOrAlias(name);
    const sourceCourse = matched
      ? cloneCourseSource(matched)
      : {
          code: '',
          nameJa: name || `Row ${row}`,
          nameEn: '',
          credits: 0,
          total: sumPoints18(points18),
          points18: clonePoints18(points18),
          sourceKind: 'manual',
          aliases: []
        };

    imported[row] = {
      row,
      code: matched?.code || '',
      nameJa: name || matched?.nameJa || '',
      nameEn: matched?.nameEn || '',
      credits: Number(matched?.credits || 0),
      sourceTotal: resolveCourseTotal(sourceCourse),
      points18: clonePoints18(points18),
      sourceCourse,
      isExternal: Boolean(sourceCourse.isExternal),
      sourceKind: sourceCourse.sourceKind || 'manual'
    };
  }

  state.planRows = imported;
  renderRowOptions();
  if (state.selectedCourse) {
    recomputePreview();
  } else {
    renderSimulationSection();
  }
  renderPlanSection();
  updateActionStates();
  return Object.keys(imported).length;
}

function reloadPlanFromWorkbook() {
  if (!state.workbook) return;
  const imported = syncPlanFromWorkbook();
  setWorkbookStatus(`Excel から履修計画ドラフトを再読込しました。${imported} 行を反映しています。${isPlanSyncedWithWorkbook() ? ' Excel と同期しています。' : ' まだ Excel へは反映していません。'}`, currentStatusKindAfterPlanChange());
}

function clearPlanRows() {
  state.planRows = {};
  renderRowOptions();
  if (state.selectedCourse) {
    recomputePreview();
  } else {
    renderSimulationSection();
  }
  renderPlanSection();
  updateActionStates();
  setWorkbookStatus('UI 上の履修計画ドラフトをクリアしました。まだ Excel は更新していません。', currentStatusKindAfterPlanChange());
}

function removePlanRow(row) {
  delete state.planRows[row];
  renderRowOptions();
  if (state.selectedCourse) {
    recomputePreview();
  } else {
    renderSimulationSection();
  }
  renderPlanSection();
  updateActionStates();
  setWorkbookStatus(`${row}行を履修計画ドラフトから削除しました。まだ Excel は更新していません。`, currentStatusKindAfterPlanChange());
}

function editPlanRow(row) {
  const entry = state.planRows[row];
  if (!entry) return;

  state.selectedCourse = entry.sourceCourse ? cloneCourseSource(entry.sourceCourse) : {
    code: entry.code,
    nameJa: entry.nameJa,
    nameEn: entry.nameEn,
    credits: entry.credits,
    total: sumPoints18(entry.points18),
    points18: clonePoints18(entry.points18),
    sourceKind: 'manual'
  };
  state.previewPoints18 = clonePoints18(entry.points18);
  renderSelectedCourse(state.selectedCourse);
  renderRowOptions(row);
  updatePreviewInputs();
  updateActionStates();
  setWorkbookStatus(`${row}行の内容を編集パネルへ読み込みました。`, currentStatusKindAfterPlanChange());
}

function captureStyleRow(sheet, row) {
  return {
    height: safeRowHeight(sheet, row),
    cells: Array.from({ length: 20 }, (_, index) => captureCellStyle(sheet.cell(`${columnLetter(index + 1)}${row}`)))
  };
}

function captureCellStyle(cell) {
  const style = {};
  STYLE_KEYS.forEach(key => {
    try {
      const value = cell.style(key);
      if (value === undefined || value === null) return;
      if (typeof value === 'object' && value !== null && Object.keys(value).length === 0) return;
      style[key] = deepClone(value);
    } catch (error) {
      console.warn(error);
    }
  });
  return style;
}

function applyStyleRow(sheet, row, styleRow) {
  if (!styleRow) return;
  if (Number.isFinite(styleRow.height) && styleRow.height > 0) {
    try {
      sheet.row(row).height(styleRow.height);
    } catch (error) {
      console.warn(error);
    }
  }

  styleRow.cells.forEach((style, index) => {
    const cell = sheet.cell(`${columnLetter(index + 1)}${row}`);
    if (style && Object.keys(style).length) {
      cell.style(deepClone(style));
    }
  });
}

function captureExtraRows(sheet, layout) {
  const rows = [];
  for (let row = layout.extraStartRow; row <= layout.extraEndRow; row += 1) {
    const style = captureStyleRow(sheet, row);
    const cells = Array.from({ length: 19 }, (_, index) => {
      const address = `${columnLetter(index + 1)}${row}`;
      const cell = sheet.cell(address);
      return {
        value: deepClone(cell.value()),
        formula: String(cell.formula() || '')
      };
    });
    rows.push({ style, cells });
  }
  return rows;
}

function clearRowContents(sheet, row) {
  for (let col = 1; col <= 20; col += 1) {
    sheet.cell(`${columnLetter(col)}${row}`).value('');
  }
}

function writeFooterStructure(sheet, startCourseEndRow, templates) {
  const courseSubtotalRow = startCourseEndRow + 1;
  const blank1Row = courseSubtotalRow + 1;
  const extraHeaderRow = courseSubtotalRow + 2;
  const extraStartRow = extraHeaderRow + 1;
  const extraEndRow = extraStartRow + templates.extraRows.length - 1;
  const extraSubtotalRow = extraEndRow + 1;
  const blank2Row = extraSubtotalRow + 1;
  const totalRow = blank2Row + 1;
  const deficitRow = totalRow + 1;

  applyStyleRow(sheet, courseSubtotalRow, templates.courseSubtotalStyle);
  writeCourseSubtotalRow(sheet, courseSubtotalRow, startCourseEndRow);

  applyStyleRow(sheet, blank1Row, templates.blank1Style);
  clearRowContents(sheet, blank1Row);

  applyStyleRow(sheet, extraHeaderRow, templates.extraHeaderStyle);
  clearRowContents(sheet, extraHeaderRow);
  sheet.cell(`A${extraHeaderRow}`).value('授業科目以外');

  templates.extraRows.forEach((rowTemplate, index) => {
    const row = extraStartRow + index;
    applyStyleRow(sheet, row, rowTemplate.style);
    clearRowContents(sheet, row);
    rowTemplate.cells.forEach((snapshot, cellIndex) => {
      const address = `${columnLetter(cellIndex + 1)}${row}`;
      if (snapshot.formula) {
        sheet.cell(address).formula(snapshot.formula);
      } else {
        sheet.cell(address).value(deepClone(snapshot.value));
      }
    });
    sheet.cell(`T${row}`).formula(`SUM(B${row}:S${row})`);
  });

  applyStyleRow(sheet, extraSubtotalRow, templates.extraSubtotalStyle);
  writeExtraSubtotalRow(sheet, extraSubtotalRow, extraStartRow, extraEndRow);

  applyStyleRow(sheet, blank2Row, templates.blank2Style);
  clearRowContents(sheet, blank2Row);

  applyStyleRow(sheet, totalRow, templates.totalStyle);
  writeTotalRow(sheet, totalRow, courseSubtotalRow, extraSubtotalRow);

  applyStyleRow(sheet, deficitRow, templates.deficitStyle);
  writeDeficitRow(sheet, deficitRow, totalRow);

  return {
    courseSubtotalRow,
    blank1Row,
    extraHeaderRow,
    extraStartRow,
    extraEndRow,
    extraSubtotalRow,
    blank2Row,
    totalRow,
    deficitRow
  };
}

function writeCourseSubtotalRow(sheet, row, courseEndRow) {
  clearRowContents(sheet, row);
  sheet.cell(`A${row}`).value('小計');
  COLUMN_LETTERS.forEach(letter => {
    sheet.cell(`${letter}${row}`).formula(`SUM(${letter}9:${letter}${courseEndRow})`);
  });
  sheet.cell(`T${row}`).formula(`SUM(B${row}:S${row})`);
}

function writeExtraSubtotalRow(sheet, row, extraStartRow, extraEndRow) {
  clearRowContents(sheet, row);
  sheet.cell(`A${row}`).value('小計');
  COLUMN_LETTERS.forEach(letter => {
    sheet.cell(`${letter}${row}`).formula(`SUM(${letter}${extraStartRow}:${letter}${extraEndRow})`);
  });
  sheet.cell(`T${row}`).formula(`SUM(T${extraStartRow}:T${extraEndRow})`);
}

function writeTotalRow(sheet, row, courseSubtotalRow, extraSubtotalRow) {
  clearRowContents(sheet, row);
  sheet.cell(`A${row}`).value('合計');
  COLUMN_LETTERS.forEach(letter => {
    sheet.cell(`${letter}${row}`).formula(`${letter}${courseSubtotalRow}+${letter}${extraSubtotalRow}`);
  });
  sheet.cell(`T${row}`).formula(`T${courseSubtotalRow}+T${extraSubtotalRow}`);
}

function writeDeficitRow(sheet, row, totalRow) {
  clearRowContents(sheet, row);
  sheet.cell(`A${row}`).value('不足分');
  COLUMN_LETTERS.forEach(letter => {
    sheet.cell(`${letter}${row}`).formula(`${letter}8-${letter}${totalRow}`);
  });
  sheet.cell(`T${row}`).formula(`T8-T${totalRow}`);
}

function prepareRowTemplates(sheet, layout) {
  const courseStyleRow = captureStyleRow(sheet, Math.min(20, layout.courseEndRow));
  return {
    courseStyleRow,
    courseSubtotalStyle: captureStyleRow(sheet, layout.courseSubtotalRow),
    blank1Style: captureStyleRow(sheet, layout.blankAfterCourseSubtotalRow),
    extraHeaderStyle: captureStyleRow(sheet, layout.extraHeaderRow),
    extraRows: captureExtraRows(sheet, layout),
    extraSubtotalStyle: captureStyleRow(sheet, layout.extraSubtotalRow),
    blank2Style: captureStyleRow(sheet, layout.blankAfterExtraSubtotalRow),
    totalStyle: captureStyleRow(sheet, layout.totalRow),
    deficitStyle: captureStyleRow(sheet, layout.deficitRow)
  };
}

function ensureCourseRows(sheet, startRow, endRow, existingCourseEndRow, courseStyleRow) {
  for (let row = startRow; row <= endRow; row += 1) {
    if (row > existingCourseEndRow) {
      applyStyleRow(sheet, row, courseStyleRow);
    }
    clearCourseRow(sheet, row);
  }
}

function clearTrailingArea(sheet, fromRow, toRow) {
  if (toRow < fromRow) return;
  for (let row = fromRow; row <= toRow; row += 1) {
    clearRowContents(sheet, row);
  }
}

function applyPlanToWorkbook() {
  if (!state.workbook) return;

  try {
    const sheet = state.workbook.sheet(state.sheetName);
    const layout = detectSheetLayout(sheet);
    const templates = prepareRowTemplates(sheet, layout);
    const targetCourseEndRow = expectedCourseEndRow();
    const oldDeficitRow = layout.deficitRow;

    ensureCourseRows(sheet, 11, targetCourseEndRow, layout.courseEndRow, templates.courseStyleRow);

    getPlanEntries().forEach(entry => {
      writePoints18ToRow(sheet, entry.row, entry.nameJa || entry.nameEn || entry.code || '', entry.points18);
    });

    const newLayout = writeFooterStructure(sheet, targetCourseEndRow, templates);
    clearTrailingArea(sheet, newLayout.deficitRow + 1, oldDeficitRow);

    renderRowOptions();
    if (state.selectedCourse) {
      recomputePreview();
    } else {
      renderSimulationSection();
    }
    renderPlanSection();
    updateActionStates();
    setWorkbookStatus(`履修計画 ${getPlanEntries().length} 行を Excel に一括反映しました。小計・合計・不足分も再配置しています。`, 'ok');
  } catch (error) {
    console.error(error);
    setWorkbookStatus('Excel への一括反映に失敗しました。テンプレート構造を確認してください。', 'warn');
  }
}

async function downloadWorkbook() {
  if (!state.workbook) return;
  if (!isPlanSyncedWithWorkbook()) {
    setWorkbookStatus('先に「履修計画を Excel に一括反映」を押してからダウンロードしてください。', 'warn');
    return;
  }

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

function resolveCourseTotal(course) {
  if (!course) return 0;
  if (course.total !== undefined && course.total !== null && course.total !== '') {
    return asInt(course.total);
  }
  if (Array.isArray(course.points18) && course.points18.length === 18) {
    return sumPoints18(course.points18);
  }
  if (course.points8 && Object.keys(course.points8).length) {
    return sumPoints8(course.points8);
  }
  if (Array.isArray(course.rawCompetenceValues) && course.rawCompetenceValues.length) {
    return sumNumbers(course.rawCompetenceValues);
  }
  if (course.credits) {
    return Math.round(asNumber(course.credits) * 100);
  }
  return 0;
}

function cloneCourseSource(course) {
  return {
    code: course?.code || '',
    nameJa: course?.nameJa || '',
    nameEn: course?.nameEn || '',
    credits: Number(course?.credits || 0),
    total: resolveCourseTotal(course),
    points8: course?.points8 ? { ...course.points8 } : null,
    rawCompetenceValues: Array.isArray(course?.rawCompetenceValues) ? course.rawCompetenceValues.map(asInt) : null,
    points18: Array.isArray(course?.points18) ? clonePoints18(course.points18) : null,
    aliases: Array.isArray(course?.aliases) ? course.aliases.slice() : [],
    sourceKind: course?.sourceKind || '',
    isExternal: Boolean(course?.isExternal)
  };
}

function displayCourseName(entry) {
  return String(entry?.nameJa || entry?.nameEn || entry?.code || '名称未設定');
}

function zeroPoints8() {
  return GROUP_KEYS.reduce((acc, key) => {
    acc[key] = 0;
    return acc;
  }, {});
}

function arrayToPoints8(values) {
  return GROUP_KEYS.reduce((acc, key, index) => {
    acc[key] = asInt(values[index]);
    return acc;
  }, {});
}

function sumPoints8(points8) {
  return GROUP_KEYS.reduce((sum, key) => sum + asInt(points8?.[key]), 0);
}

function sumPoints18(values) {
  return (values || []).reduce((sum, value) => sum + asInt(value), 0);
}

function clonePoints18(values) {
  return Array.from({ length: 18 }, (_, index) => asInt(values?.[index]));
}

function sumNumbers(values) {
  return (values || []).reduce((sum, value) => sum + Number(value || 0), 0);
}

function chip(text) {
  return `<span class="chip">${escapeHtml(text)}</span>`;
}

function formatSigned(value) {
  const num = asIntSigned(value);
  if (num > 0) return `+${num}`;
  return String(num);
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

function asNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function asIntSigned(value) {
  const num = Number(value);
  return Number.isFinite(num) ? Math.round(num) : 0;
}

function columnLetter(index) {
  return String.fromCharCode('A'.charCodeAt(0) + index - 1);
}

function safeRowHeight(sheet, row) {
  try {
    const value = sheet.row(row).height();
    return Number.isFinite(Number(value)) ? Number(value) : null;
  } catch (error) {
    return null;
  }
}

function deepClone(value) {
  if (value === undefined) return undefined;
  if (value === null) return null;
  return JSON.parse(JSON.stringify(value));
}

function escapeHtml(value) {
  return String(value || '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
