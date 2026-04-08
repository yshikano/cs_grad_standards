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
const GROUP_LABELS = GROUP_CONFIG.map(group => group.label);
const GROUP_TOTAL_REQUIREMENTS = GROUP_CONFIG.map(group => group.requiredWeights.reduce((sum, value) => sum + value, 0));
const COLUMN_LABELS = GROUP_CONFIG.flatMap(group => group.cols);
const COLUMN_LETTERS = Array.from({ length: 18 }, (_, index) => String.fromCharCode('B'.charCodeAt(0) + index));
const REQUIRED_POINTS18 = GROUP_CONFIG.flatMap(group => group.requiredWeights);
const REQUIRED_TOTAL = REQUIRED_POINTS18.reduce((sum, value) => sum + value, 0);
const CUSTOM_MODE_OPTIONS = {
  competence8: GROUP_LABELS,
  viewpoint18: COLUMN_LABELS
};

const MODE_CONFIG = {
  plan: {
    key: 'plan',
    label: '履修計画',
    description: 'UI 上で履修計画を検討するモードです。Excel への書き込みは行いません。',
    canWrite: false,
    sheetCandidates: []
  },
  midterm: {
    key: 'midterm',
    label: 'M中間評価',
    description: 'M中間評価シートを対象にした実運用モードです。',
    canWrite: true,
    sheetCandidates: ['M中間評価', 'M Mid-term evaluation']
  },
  final: {
    key: 'final',
    label: 'M最終評価',
    description: 'M最終評価シートを対象にした実運用モードです。中間評価の内容を初期値にできます。',
    canWrite: true,
    sheetCandidates: ['M最終評価', 'M final evaluation']
  }
};

const MODE_KEYS = Object.keys(MODE_CONFIG);
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
const FIXED_COURSE_ROWS = [11, 12, 13, 14];
const MIN_COURSE_END_ROW = 30;
const MIN_EXTRA_ROWS = 3;

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

const ZERO_18 = Array(18).fill(0);
const INTERNATIONALITY_ALLOWED_MASK = Array.from({ length: 18 }, (_, index) => index === 8 || index === 9);

const FIXED_EXTRA_TEMPLATES = {
  report1: {
    key: 'report1',
    label: '達成度自己点検レポート1（M1終了時）',
    note: '汎用コンピテンス 2. マネジメント能力① に 20 点を固定配分します。',
    fixedPoints18: [0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
  },
  report2: {
    key: 'report2',
    label: '達成度自己点検レポート2（M2終了時）',
    note: '汎用コンピテンス 2. マネジメント能力① に 20 点を固定配分します。',
    fixedPoints18: [0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
  },
  infoss: {
    key: 'infoss',
    label: 'INFOSS情報倫理',
    note: '専門コンピテンス III 倫理観① に 50 点を固定配分します。',
    fixedPoints18: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 50, 0]
  },
  aprin: {
    key: 'aprin',
    label: 'APRIN eラーニングプログラム',
    note: '専門コンピテンス III 倫理観① に 30 点を固定配分します。',
    fixedPoints18: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 30, 0]
  }
};

const EXTRA_TYPE_OPTIONS = [
  { key: 'paper', label: '論文・発表' },
  { key: 'patent', label: '特許' },
  { key: 'labwg', label: '研究室WG' },
  { key: 'collabotics', label: 'CSスペシャルワークショップ' },
  { key: 'outsideProgram', label: '学位プログラム外活動' },
  { key: 'ta', label: 'TA' },
  { key: 'language', label: 'TOEIC / TOEFL' },
  { key: 'fixed', label: '固定項目' },
  { key: 'custom', label: 'その他（自由入力）' }
];

const EXTRA_TYPE_LABELS = EXTRA_TYPE_OPTIONS.reduce((acc, item) => {
  acc[item.key] = item.label;
  return acc;
}, {});

const state = {
  dataset: null,
  courses: [],
  workbook: null,
  workbookFileName: '',
  availableSheets: {
    midterm: null,
    final: null
  },
  currentMode: 'plan',
  selectedCourse: null,
  previewPoints18: clonePoints18(ZERO_18),
  customValues8: Array(8).fill(0),
  customValues18: Array(18).fill(0),
  nextExtraId: 1,
  modeDrafts: {
    plan: createEmptyModeDraft(),
    midterm: createEmptyModeDraft(),
    final: createEmptyModeDraft()
  },
  extraForm: createDefaultExtraForm()
};

const els = {};

document.addEventListener('DOMContentLoaded', async () => {
  bindElements();
  wireEvents();
  renderModeSwitch();
  renderRowOptions();
  renderPreviewGrid();
  renderCustomValueGrid();
  renderExtraTypeOptions();
  renderExtraPreviewGrid();
  renderFixedCourseButtons();
  renderFixedExtraButtons();
  renderExtraDynamicFields();
  renderSelectedCourse(null);
  renderModeInfo();
  renderPlanSection();
  updateActionStates();
  await loadDataset();
});

function bindElements() {
  els.xlsxFile = document.getElementById('xlsxFile');
  els.modeSwitch = document.getElementById('modeSwitch');
  els.seedFinalBtn = document.getElementById('seedFinalBtn');
  els.modeInfo = document.getElementById('modeInfo');
  els.workbookStatus = document.getElementById('workbookStatus');

  els.searchInput = document.getElementById('searchInput');
  els.searchMeta = document.getElementById('searchMeta');
  els.searchResults = document.getElementById('searchResults');
  els.datasetInfo = document.getElementById('datasetInfo');

  els.customName = document.getElementById('customName');
  els.customCode = document.getElementById('customCode');
  els.customCredits = document.getElementById('customCredits');
  els.customInputMode = document.getElementById('customInputMode');
  els.customValueGrid = document.getElementById('customValueGrid');
  els.customValueMeta = document.getElementById('customValueMeta');
  els.customUseBtn = document.getElementById('customUseBtn');
  els.customClearBtn = document.getElementById('customClearBtn');

  els.fixedCourseButtons = document.getElementById('fixedCourseButtons');
  els.selectedCourseBox = document.getElementById('selectedCourseBox');
  els.splitMode = document.getElementById('splitMode');
  els.targetRow = document.getElementById('targetRow');
  els.previewGrid = document.getElementById('previewGrid');
  els.previewTotal = document.getElementById('previewTotal');
  els.courseRecomputeBtn = document.getElementById('courseRecomputeBtn');
  els.planAddBtn = document.getElementById('planAddBtn');

  els.extraType = document.getElementById('extraType');
  els.extraLabel = document.getElementById('extraLabel');
  els.extraSplitMode = document.getElementById('extraSplitMode');
  els.extraNote = document.getElementById('extraNote');
  els.extraDynamicFields = document.getElementById('extraDynamicFields');
  els.extraMeta = document.getElementById('extraMeta');
  els.fixedExtraButtons = document.getElementById('fixedExtraButtons');
  els.extraPreviewGrid = document.getElementById('extraPreviewGrid');
  els.extraPreviewTotal = document.getElementById('extraPreviewTotal');
  els.extraRecomputeBtn = document.getElementById('extraRecomputeBtn');
  els.extraAddBtn = document.getElementById('extraAddBtn');
  els.extraClearBtn = document.getElementById('extraClearBtn');
  els.extraList = document.getElementById('extraList');

  els.applyPlanBtn = document.getElementById('applyPlanBtn');
  els.downloadBtn = document.getElementById('downloadBtn');
  els.reloadPlanBtn = document.getElementById('reloadPlanBtn');
  els.clearPlanBtn = document.getElementById('clearPlanBtn');
  els.syncStatus = document.getElementById('syncStatus');

  els.planList = document.getElementById('planList');
  els.planCourseCount = document.getElementById('planCourseCount');
  els.planCredits = document.getElementById('planCredits');
  els.planPointsTotal = document.getElementById('planPointsTotal');
  els.extraItemCount = document.getElementById('extraItemCount');
  els.extraPointsTotal = document.getElementById('extraPointsTotal');
  els.projectedTotal = document.getElementById('projectedTotal');
  els.unmetCount = document.getElementById('unmetCount');
  els.totalsBody = document.getElementById('totalsBody');
  els.totalsFoot = document.getElementById('totalsFoot');
}

function wireEvents() {
  els.xlsxFile.addEventListener('change', onWorkbookSelected);
  els.seedFinalBtn.addEventListener('click', () => seedFinalFromMidterm(true));

  els.searchInput.addEventListener('input', () => renderSearchResults(searchCourses(els.searchInput.value)));

  els.customInputMode.addEventListener('change', () => {
    renderCustomValueGrid();
    updateCustomValueMeta();
  });
  ['input', 'change'].forEach(eventName => {
    els.customCredits.addEventListener(eventName, updateCustomValueMeta);
    els.customName.addEventListener(eventName, updateCustomValueMeta);
    els.customCode.addEventListener(eventName, updateCustomValueMeta);
  });
  els.customUseBtn.addEventListener('click', useCustomAsSelection);
  els.customClearBtn.addEventListener('click', clearCustomBuilder);

  els.splitMode.addEventListener('change', recomputePreview);
  els.targetRow.addEventListener('change', () => {
    if (state.selectedCourse) {
      recomputePreview();
    }
  });
  els.courseRecomputeBtn.addEventListener('click', recomputePreview);
  els.planAddBtn.addEventListener('click', addSelectedCourseToDraft);

  els.extraType.addEventListener('change', () => {
    state.extraForm.type = els.extraType.value;
    state.extraForm.editingId = null;
    state.extraForm.points18 = clonePoints18(ZERO_18);
    renderExtraDynamicFields();
    recomputeExtraPreview();
    renderExtraList();
  });
  els.extraLabel.addEventListener('input', () => {
    state.extraForm.label = els.extraLabel.value;
    updateExtraMeta();
  });
  els.extraSplitMode.addEventListener('change', () => {
    state.extraForm.splitMode = els.extraSplitMode.value;
    recomputeExtraPreview();
  });
  els.extraNote.addEventListener('input', () => {
    state.extraForm.note = els.extraNote.value;
    updateExtraMeta();
  });
  els.extraRecomputeBtn.addEventListener('click', recomputeExtraPreview);
  els.extraAddBtn.addEventListener('click', addExtraItemToDraft);
  els.extraClearBtn.addEventListener('click', clearExtraForm);

  els.applyPlanBtn.addEventListener('click', applyCurrentModeToWorkbook);
  els.downloadBtn.addEventListener('click', downloadWorkbook);
  els.reloadPlanBtn.addEventListener('click', reloadCurrentModeFromWorkbook);
  els.clearPlanBtn.addEventListener('click', clearCurrentDraft);
}

function createEmptyModeDraft() {
  return {
    courseRows: {},
    extraItems: [],
    seededFromWorkbook: false,
    seededFromMidterm: false
  };
}

function createDefaultExtraForm() {
  return {
    editingId: null,
    type: 'paper',
    label: '',
    note: '',
    splitMode: 'smart',
    points18: clonePoints18(ZERO_18),
    paperCounts: {
      firstJournal: 0,
      firstIntl: 0,
      firstOther: 0,
      nonFirstJournal: 0,
      nonFirstIntl: 0,
      nonFirstOther: 0
    },
    patentContribution: 100,
    labwgPoints: 0,
    collaboticsPoints: 0,
    outsideProgramPoints: 0,
    taSlots: 0,
    languageKind: 'toeic',
    languageScore: 0,
    fixedKey: 'report1',
    customTotal: 0
  };
}

function renderModeSwitch() {
  els.modeSwitch.innerHTML = '';
  MODE_KEYS.forEach(mode => {
    const config = MODE_CONFIG[mode];
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `mode-switch__button${state.currentMode === mode ? ' is-active' : ''}`;
    button.textContent = config.label;
    button.addEventListener('click', () => setMode(mode));
    els.modeSwitch.appendChild(button);
  });
}

function setMode(mode) {
  if (!MODE_CONFIG[mode]) return;
  state.currentMode = mode;
  maybeSeedFinalFromMidterm();
  renderModeSwitch();
  renderModeInfo();
  renderRowOptions();
  renderFixedCourseButtons();
  renderFixedExtraButtons();
  if (state.selectedCourse) {
    recomputePreview();
  } else {
    updatePreviewInputs();
  }
  recomputeExtraPreview();
  renderPlanSection();
  updateActionStates();
}

function renderModeInfo() {
  const config = MODE_CONFIG[state.currentMode];
  const sheetName = getSheetNameForMode(state.currentMode);

  els.planAddBtn.textContent = `${config.label}の授業科目へ追加 / 更新`;
  els.extraAddBtn.textContent = `${config.label}の授業科目以外へ追加 / 更新`;
  els.applyPlanBtn.textContent = config.canWrite
    ? `${config.label}を Excel に一括反映`
    : '履修計画モードでは Excel に書き込みません';

  if (state.currentMode === 'plan') {
    els.modeInfo.textContent = MODE_CONFIG.plan.description;
    els.modeInfo.className = 'status status--muted';
    return;
  }

  if (!state.workbook) {
    els.modeInfo.textContent = `${config.description} Excel を読み込むと ${config.label} シートの内容をドラフトへ取り込みます。`;
    els.modeInfo.className = 'status status--muted';
    return;
  }

  const currentDraft = getDraft(state.currentMode);
  const syncLabel = isModeDraftSyncedWithWorkbook(state.currentMode) ? '同期済み' : '未反映';
  let text = `${config.description} 対象シート: ${sheetName || '未検出'} / ${syncLabel}`;
  if (state.currentMode === 'final' && currentDraft.seededFromMidterm) {
    text += ' / 中間評価ドラフトを初期値として使っています。';
  }
  els.modeInfo.textContent = text;
  els.modeInfo.className = `status ${isModeDraftSyncedWithWorkbook(state.currentMode) ? 'status--ok' : 'status--pending'}`;
}

function getDraft(mode = state.currentMode) {
  return state.modeDrafts[mode];
}

function getCurrentDraft() {
  return getDraft(state.currentMode);
}

function cloneModeDraft(draft) {
  return {
    courseRows: clonePlanRows(draft.courseRows),
    extraItems: (draft.extraItems || []).map(cloneExtraItem),
    seededFromWorkbook: Boolean(draft.seededFromWorkbook),
    seededFromMidterm: Boolean(draft.seededFromMidterm)
  };
}

function cloneExtraItem(item) {
  return {
    ...item,
    points18: clonePoints18(item.points18),
    formSnapshot: item.formSnapshot ? deepClone(item.formSnapshot) : null
  };
}

function hasDraftContent(draft = getCurrentDraft()) {
  return getPlanEntries(draft.courseRows).length > 0 || (draft.extraItems || []).length > 0;
}

function maybeSeedFinalFromMidterm() {
  if (!MODE_CONFIG.final.canWrite) return false;
  const finalDraft = getDraft('final');
  const midtermDraft = getDraft('midterm');
  if (hasDraftContent(finalDraft) || !hasDraftContent(midtermDraft)) {
    return false;
  }
  state.modeDrafts.final = cloneModeDraft(midtermDraft);
  state.modeDrafts.final.seededFromMidterm = true;
  return true;
}

function seedFinalFromMidterm(force = false) {
  const midtermDraft = getDraft('midterm');
  if (!hasDraftContent(midtermDraft)) {
    setWorkbookStatus('中間評価ドラフトが空なので、最終評価を初期化できません。', 'warn');
    return;
  }

  const finalDraft = getDraft('final');
  if (!force && hasDraftContent(finalDraft)) {
    return;
  }

  state.modeDrafts.final = cloneModeDraft(midtermDraft);
  state.modeDrafts.final.seededFromMidterm = true;
  if (state.currentMode === 'final') {
    renderModeInfo();
    renderRowOptions();
    if (state.selectedCourse) {
      recomputePreview();
    }
    recomputeExtraPreview();
    renderPlanSection();
    updateActionStates();
  }
  setWorkbookStatus('M中間評価のドラフト内容をもとに M最終評価ドラフトを初期化しました。', 'pending');
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
    const sourceUrl = escapeHtml(state.dataset.meta?.sourceUrl || '#');
    const courseListSourceUrl = escapeHtml(state.dataset.meta?.courseListSourceUrl || '#');
    const sourceParts = [
      `<div><strong>データ件数:</strong> ${state.dataset.meta?.count ?? state.courses.length} 件</div>`,
      `<div><strong>カリキュラム・マップ:</strong> <a href="${sourceUrl}" target="_blank" rel="noreferrer">公開 PDF</a></div>`
    ];
    if (state.dataset.meta?.courseListSourceUrl) {
      sourceParts.push(`<div><strong>開設授業科目一覧:</strong> <a href="${courseListSourceUrl}" target="_blank" rel="noreferrer">公式ページ</a></div>`);
    }
    sourceParts.push(`<div><strong>メモ:</strong> ${notes || 'なし'}</div>`);
    els.datasetInfo.innerHTML = sourceParts.join('');

    renderSearchResults(searchCourses(''));
    renderFixedCourseButtons();
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
      ? 'カリキュラム・マップ由来'
      : (course.sourceKind === 'sequence'
        ? '非ゼロ列を自動補完'
        : (course.sourceKind === 'creditsOnly' ? '単位数ベースの初期配点' : 'カスタム'));
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

  let message = '元のコンピテンス / 観点を入力すると、1単位 100 点へ正規化した上で不足差分を見ながら初期配点します。';
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
    setWorkbookStatus(`「${course.nameJa}」を選択中の科目へセットしました。`, currentStatusKindAfterChange());
  } catch (error) {
    setWorkbookStatus(error.message, 'warn');
  }
}

function renderFixedCourseButtons() {
  els.fixedCourseButtons.innerHTML = '';
  FIXED_COURSE_ROWS.forEach(row => {
    const name = DEFAULT_ROW_NAMES[row];
    const course = findCourseByNameOrAlias(name);
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'button button--small';
    button.textContent = `${row}行: ${name}`;
    button.disabled = !course;
    button.addEventListener('click', () => {
      renderRowOptions(row);
      selectCourse(course || {
        code: '',
        nameJa: name,
        nameEn: '',
        credits: 0,
        total: 0,
        points18: clonePoints18(ZERO_18),
        aliases: [name],
        sourceKind: 'manual'
      });
    });
    els.fixedCourseButtons.appendChild(button);
  });
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
  if (Array.isArray(course?.points18) && course.points18.length === 18) {
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
        <div class="help-text">不足差分を見ながら 8 コンピテンスへ補完し、その後 18 観点へ初期展開します。</div>
      </div>
    `;
  } else {
    pointsHtml = `
      <div class="point8-item point8-item--wide">
        <span>元データ</span>
        <strong>大学院スタンダード配点が未登録のため、単位数換算 ${asInt(resolveCourseTotal(course))} 点をもとに初期配点します</strong>
      </div>
    `;
  }

  const badges = [
    chip(course.sourceKind === 'sequence' ? '自動補完対象' : (course.sourceKind === 'creditsOnly' ? '単位数ベース' : '元配点あり')),
    chip(`最終合計 ${asInt(resolveCourseTotal(course))} 点`)
  ].join('');

  els.selectedCourseBox.classList.remove('empty');
  els.selectedCourseBox.innerHTML = `
    <div class="result-card__code">${escapeHtml(course.code || 'カスタム')}</div>
    <h3>${escapeHtml(course.nameJa || '名称未設定')}</h3>
    <div class="muted">${escapeHtml(subline)}</div>
    <div class="selected-course__source">${badges}</div>
    <div class="points8">${pointsHtml}</div>
    <div class="help-text">大学院スタンダードの公開カリキュラム・マップは 8 コンピテンス配点なので、下の 18 観点は初期案です。必要なら手動で修正してください。</div>
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
    state.previewPoints18 = clonePoints18(ZERO_18);
    updatePreviewInputs();
    return;
  }

  const targetRow = Number(els.targetRow.value || 11);
  const positiveDeficits18 = getPositiveDeficits18({ excludingCourseRow: targetRow });
  state.previewPoints18 = resolveCoursePoints18(state.selectedCourse, {
    mode: els.splitMode.value,
    positiveDeficits18
  });

  updatePreviewInputs();
}

function resolveCoursePoints18(course, options = {}) {
  const mode = options.mode || 'smart';
  const positiveDeficits18 = clonePoints18(options.positiveDeficits18 || ZERO_18);
  const total = resolveCourseTotal(course);

  if (Array.isArray(course?.points18) && course.points18.length === 18) {
    return resolveFromOriginal18(course.points18, total, mode, positiveDeficits18);
  }

  const has8Source = Boolean(course?.points8 && Object.keys(course.points8).length > 0)
    || (Array.isArray(course?.rawCompetenceValues) && course.rawCompetenceValues.length > 0);

  if (!has8Source) {
    if (total <= 0) return clonePoints18(ZERO_18);
    if (mode === 'even') {
      return allocateByWeights(total, Array(18).fill(1));
    }
    if (mode === 'required') {
      return allocateByWeights(total, REQUIRED_POINTS18);
    }
    return allocateComposite(total, ZERO_18, positiveDeficits18, REQUIRED_POINTS18, Array(18).fill(true));
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
  let maxRow = MIN_COURSE_END_ROW;
  const highestPlanned = getHighestPlannedRow(getCurrentDraft().courseRows);
  if (highestPlanned > 0) {
    maxRow = Math.max(maxRow, highestPlanned + 1);
  }
  const sheetName = getSheetNameForMode(state.currentMode);
  if (state.workbook && sheetName) {
    try {
      const sheet = state.workbook.sheet(sheetName);
      const layout = detectSheetLayout(sheet);
      maxRow = Math.max(maxRow, layout.courseEndRow + 1);
    } catch (error) {
      console.warn(error);
    }
  }
  return maxRow;
}

function getRowLabel(row) {
  const planned = getCurrentDraft().courseRows[row];
  if (planned) {
    return `${row}行: ${displayCourseName(planned)}`;
  }

  const sheetName = getSheetNameForMode(state.currentMode);
  if (state.workbook && sheetName) {
    try {
      const sheet = state.workbook.sheet(sheetName);
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

function addSelectedCourseToDraft() {
  if (!state.selectedCourse) return;
  const row = Number(els.targetRow.value || 11);
  getCurrentDraft().courseRows[row] = makePlanEntryFromSelection(row);
  renderRowOptions();
  selectNextAvailableRow(row);
  state.selectedCourse = null;
  state.previewPoints18 = clonePoints18(ZERO_18);
  renderSelectedCourse(null);
  updatePreviewInputs();
  renderPlanSection();
  setWorkbookStatus(`${MODE_CONFIG[state.currentMode].label}の ${row} 行を更新しました。${MODE_CONFIG[state.currentMode].canWrite ? 'まだ Excel へは反映していません。' : '履修計画モードなので Excel は更新していません。'}`, currentStatusKindAfterChange());
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
  const rows = getCurrentDraft().courseRows;
  const nextEmpty = candidates.find(row => !rows[row]) || maxRow;
  renderRowOptions(nextEmpty);
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

function clonePlanRows(source = {}) {
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

function getPlanEntries(planRows = getCurrentDraft().courseRows) {
  return Object.values(planRows)
    .map(entry => ({ ...entry, row: Number(entry.row) }))
    .sort((a, b) => a.row - b.row);
}

function getHighestPlannedRow(planRows = getCurrentDraft().courseRows) {
  const entries = getPlanEntries(planRows);
  return entries.length ? Math.max(...entries.map(entry => entry.row)) : 0;
}

function getCourseTotals18(courseRows = getCurrentDraft().courseRows) {
  const totals = Array(18).fill(0);
  getPlanEntries(courseRows).forEach(entry => {
    entry.points18.forEach((value, index) => {
      totals[index] += asInt(value);
    });
  });
  return totals;
}

function getExtraTotals18(extraItems = getCurrentDraft().extraItems) {
  const totals = Array(18).fill(0);
  (extraItems || []).forEach(item => {
    item.points18.forEach((value, index) => {
      totals[index] += asInt(value);
    });
  });
  return totals;
}

function getMetrics(draft = getCurrentDraft()) {
  const subtotal18 = getCourseTotals18(draft.courseRows);
  const extra18 = getExtraTotals18(draft.extraItems);
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

function getPositiveDeficits18(options = {}) {
  const draft = cloneModeDraft(getCurrentDraft());
  if (options.excludingCourseRow !== undefined && options.excludingCourseRow !== null) {
    delete draft.courseRows[options.excludingCourseRow];
  }
  if (options.excludingExtraId !== undefined && options.excludingExtraId !== null) {
    draft.extraItems = draft.extraItems.filter(item => item.id !== options.excludingExtraId);
  }
  const metrics = getMetrics(draft);
  return metrics.diff18.map(value => value < 0 ? Math.abs(value) : 0);
}

function renderExtraTypeOptions() {
  els.extraType.innerHTML = EXTRA_TYPE_OPTIONS.map(option => `<option value="${escapeHtml(option.key)}">${escapeHtml(option.label)}</option>`).join('');
  els.extraType.value = state.extraForm.type;
}

function renderFixedExtraButtons() {
  els.fixedExtraButtons.innerHTML = '';
  Object.values(FIXED_EXTRA_TEMPLATES).forEach(template => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'button button--small';
    button.textContent = template.label;
    button.addEventListener('click', () => {
      clearExtraForm();
      state.extraForm.type = 'fixed';
      state.extraForm.fixedKey = template.key;
      els.extraType.value = 'fixed';
      renderExtraDynamicFields();
      recomputeExtraPreview();
    });
    els.fixedExtraButtons.appendChild(button);
  });
}

function renderExtraDynamicFields() {
  const form = state.extraForm;
  els.extraLabel.value = form.label || '';
  els.extraNote.value = form.note || '';
  els.extraSplitMode.value = form.splitMode || 'smart';
  els.extraType.value = form.type;

  let html = '';
  switch (form.type) {
    case 'paper':
      html = `
        <div class="dynamic-fields__group">
          <div class="dynamic-fields__grid">
            ${renderLabeledNumberField('paper-first-journal', '筆頭著者 / 査読有り論文誌 本数', form.paperCounts.firstJournal)}
            ${renderLabeledNumberField('paper-first-intl', '筆頭著者 / 査読有り国際会議かつ発表 本数', form.paperCounts.firstIntl)}
            ${renderLabeledNumberField('paper-first-other', '筆頭著者 / それ以外の論文・発表 本数', form.paperCounts.firstOther)}
            ${renderLabeledNumberField('paper-nonfirst-journal', '非筆頭 / 査読有り論文誌 本数', form.paperCounts.nonFirstJournal)}
            ${renderLabeledNumberField('paper-nonfirst-intl', '非筆頭 / 査読有り国際会議かつ発表 本数', form.paperCounts.nonFirstIntl)}
            ${renderLabeledNumberField('paper-nonfirst-other', '非筆頭 / それ以外の論文・発表 本数', form.paperCounts.nonFirstOther)}
          </div>
          <div class="help-box">筆頭論文・非筆頭論文を含めて複数本ある場合でも、合計ではなく最も高いポイントだけを採用します。</div>
        </div>
      `;
      break;
    case 'patent':
      html = `
        ${renderLabeledNumberField('patent-contribution', '貢献率（%）', form.patentContribution, { min: 0, max: 100, step: 1 })}
        <div class="help-box">特許は 300 × 貢献率で計算します。例: 50% なら 150 点です。</div>
      `;
      break;
    case 'labwg':
      html = `
        ${renderLabeledNumberField('labwg-points', '配分点（目安: 0〜300）', form.labwgPoints, { min: 0, max: 300, step: 1 })}
        <div class="help-box">研究室に関係する活動に対して配分します。実態に合わせて指導教員と相談してください。</div>
      `;
      break;
    case 'collabotics':
      html = `
        ${renderLabeledNumberField('collabotics-points', '配分点（1〜500）', form.collaboticsPoints, { min: 0, max: 500, step: 1 })}
        <div class="help-box">CSスペシャルワークショップ（CollaboTICS）の貢献度に応じて実行委員会が判断するポイントを入力します。</div>
      `;
      break;
    case 'outsideProgram':
      html = `
        ${renderLabeledNumberField('outsideprogram-points', '配分点（目安: 0〜300）', form.outsideProgramPoints, { min: 0, step: 1 })}
        <div class="help-box">全学プロジェクト、コンテスト、インターンシップ、国外活動などをまとめて扱えます。指導教員と相談の上で点数を決めてください。</div>
      `;
      break;
    case 'ta':
      html = `
        ${renderLabeledNumberField('ta-slots', 'TA コマ数（1コマ = 1.5 時間）', form.taSlots, { min: 0, step: 1 })}
        <div class="help-box">1 コマあたり 4 ポイントで計算します。</div>
      `;
      break;
    case 'language':
      html = `
        <label>
          <span>試験種別</span>
          <select id="language-kind">
            <option value="toeic" ${form.languageKind === 'toeic' ? 'selected' : ''}>TOEIC</option>
            <option value="toefl" ${form.languageKind === 'toefl' ? 'selected' : ''}>TOEFL</option>
          </select>
        </label>
        ${renderLabeledNumberField('language-score', 'スコア', form.languageScore, { min: 0, step: 1 })}
        <div class="help-box">TOEIC / TOEFL は国際性（J, K 列）のみに配点します。40 / 60 / 80 / 100 点へ離散化します。</div>
      `;
      break;
    case 'fixed':
      html = `
        <label>
          <span>固定項目</span>
          <select id="fixed-key">
            ${Object.values(FIXED_EXTRA_TEMPLATES).map(template => `<option value="${escapeHtml(template.key)}" ${template.key === form.fixedKey ? 'selected' : ''}>${escapeHtml(template.label)}</option>`).join('')}
          </select>
        </label>
        <div class="help-box">${escapeHtml(FIXED_EXTRA_TEMPLATES[form.fixedKey]?.note || '')}</div>
      `;
      break;
    case 'custom':
      html = `
        ${renderLabeledNumberField('custom-total', '総点', form.customTotal, { min: 0, step: 1 })}
        <div class="help-box">総点を決めてから、18 観点へ手動で配点できます。</div>
      `;
      break;
    default:
      break;
  }
  els.extraDynamicFields.innerHTML = html;
  wireExtraDynamicFieldEvents();
}

function renderLabeledNumberField(id, label, value, options = {}) {
  const min = options.min ?? 0;
  const max = options.max !== undefined ? `max="${options.max}"` : '';
  const step = options.step ?? 1;
  return `
    <label>
      <span>${escapeHtml(label)}</span>
      <input id="${escapeHtml(id)}" type="number" min="${min}" ${max} step="${step}" value="${asIntSigned(value)}">
    </label>
  `;
}

function wireExtraDynamicFieldEvents() {
  const bindNumeric = (id, setter) => {
    const input = document.getElementById(id);
    if (!input) return;
    input.addEventListener('input', () => {
      setter(input.value);
      recomputeExtraPreview();
    });
  };

  if (state.extraForm.type === 'paper') {
    bindNumeric('paper-first-journal', value => { state.extraForm.paperCounts.firstJournal = asInt(value); });
    bindNumeric('paper-first-intl', value => { state.extraForm.paperCounts.firstIntl = asInt(value); });
    bindNumeric('paper-first-other', value => { state.extraForm.paperCounts.firstOther = asInt(value); });
    bindNumeric('paper-nonfirst-journal', value => { state.extraForm.paperCounts.nonFirstJournal = asInt(value); });
    bindNumeric('paper-nonfirst-intl', value => { state.extraForm.paperCounts.nonFirstIntl = asInt(value); });
    bindNumeric('paper-nonfirst-other', value => { state.extraForm.paperCounts.nonFirstOther = asInt(value); });
    return;
  }

  if (state.extraForm.type === 'patent') {
    bindNumeric('patent-contribution', value => { state.extraForm.patentContribution = asInt(value); });
    return;
  }
  if (state.extraForm.type === 'labwg') {
    bindNumeric('labwg-points', value => { state.extraForm.labwgPoints = asInt(value); });
    return;
  }
  if (state.extraForm.type === 'collabotics') {
    bindNumeric('collabotics-points', value => { state.extraForm.collaboticsPoints = asInt(value); });
    return;
  }
  if (state.extraForm.type === 'outsideProgram') {
    bindNumeric('outsideprogram-points', value => { state.extraForm.outsideProgramPoints = asInt(value); });
    return;
  }
  if (state.extraForm.type === 'ta') {
    bindNumeric('ta-slots', value => { state.extraForm.taSlots = asInt(value); });
    return;
  }
  if (state.extraForm.type === 'language') {
    const select = document.getElementById('language-kind');
    if (select) {
      select.addEventListener('change', () => {
        state.extraForm.languageKind = select.value;
        recomputeExtraPreview();
      });
    }
    bindNumeric('language-score', value => { state.extraForm.languageScore = asInt(value); });
    return;
  }
  if (state.extraForm.type === 'fixed') {
    const select = document.getElementById('fixed-key');
    if (select) {
      select.addEventListener('change', () => {
        state.extraForm.fixedKey = select.value;
        recomputeExtraPreview();
      });
    }
    return;
  }
  if (state.extraForm.type === 'custom') {
    bindNumeric('custom-total', value => { state.extraForm.customTotal = asInt(value); });
  }
}

function renderExtraPreviewGrid() {
  els.extraPreviewGrid.innerHTML = '';
  COLUMN_LABELS.forEach((label, index) => {
    const wrapper = document.createElement('label');
    wrapper.className = 'preview-cell';
    wrapper.innerHTML = `
      <span class="preview-cell__label">${COLUMN_LETTERS[index]}列 / ${escapeHtml(label)}</span>
      <input type="number" min="0" step="1" value="0">
    `;
    const input = wrapper.querySelector('input');
    input.addEventListener('input', () => {
      state.extraForm.points18[index] = asInt(input.value);
      updateExtraPreviewTotal();
      updateExtraMeta();
    });
    els.extraPreviewGrid.appendChild(wrapper);
  });
  recomputeExtraPreview();
}

function getExtraDescriptor() {
  const form = state.extraForm;
  const label = form.label.trim();
  const note = form.note.trim();

  switch (form.type) {
    case 'paper': {
      const candidates = [];
      if (asInt(form.paperCounts.firstJournal) > 0) candidates.push(600);
      if (asInt(form.paperCounts.firstIntl) > 0) candidates.push(400);
      if (asInt(form.paperCounts.firstOther) > 0) candidates.push(200);
      if (asInt(form.paperCounts.nonFirstJournal) > 0) candidates.push(150);
      if (asInt(form.paperCounts.nonFirstIntl) > 0) candidates.push(100);
      if (asInt(form.paperCounts.nonFirstOther) > 0) candidates.push(50);
      const total = candidates.length ? Math.max(...candidates) : 0;
      return {
        type: 'paper',
        name: label || '論文・発表',
        note,
        total,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: total > 0
          ? `最も高い実績に応じて ${total} 点を採用します。複数件あっても合計にはなりません。`
          : '本数を入力すると、最も高い実績に応じた点数を採用します。'
      };
    }
    case 'patent': {
      const total = Math.round(300 * Math.max(0, Math.min(100, asInt(form.patentContribution))) / 100);
      return {
        type: 'patent',
        name: label || '特許',
        note,
        total,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: `300 × 貢献率で計算し、現在は ${total} 点です。`
      };
    }
    case 'labwg': {
      const total = asInt(form.labwgPoints);
      return {
        type: 'labwg',
        name: label || '研究室WG',
        note,
        total,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: '研究室に関係する活動として自由配点できます。'
      };
    }
    case 'collabotics': {
      const total = asInt(form.collaboticsPoints);
      return {
        type: 'collabotics',
        name: label || 'CSスペシャルワークショップ',
        note,
        total,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: '貢献度に応じて 1〜500 点を自由配点できます。'
      };
    }
    case 'outsideProgram': {
      const total = asInt(form.outsideProgramPoints);
      return {
        type: 'outsideProgram',
        name: label || '学位プログラム外活動',
        note,
        total,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: '外部活動の実態に応じて自由配点できます。'
      };
    }
    case 'ta': {
      const total = asInt(form.taSlots) * 4;
      return {
        type: 'ta',
        name: label || 'TA',
        note,
        total,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: `TA ${asInt(form.taSlots)} コマ × 4 点 = ${total} 点です。`
      };
    }
    case 'language': {
      const rawScore = asInt(form.languageScore);
      const total = scoreLanguagePoints(form.languageKind, rawScore);
      const examLabel = form.languageKind === 'toefl' ? 'TOEFL' : 'TOEIC';
      return {
        type: 'language',
        name: label || `${examLabel}${rawScore ? ` ${rawScore}` : ''}`,
        note,
        total,
        allowedMask: INTERNATIONALITY_ALLOWED_MASK,
        fixedPoints18: null,
        message: `${examLabel} ${rawScore || 0} 点を ${total} 点へ離散化し、国際性（J, K 列）にのみ配点します。`
      };
    }
    case 'fixed': {
      const template = FIXED_EXTRA_TEMPLATES[form.fixedKey];
      return {
        type: 'fixed',
        name: label || template.label,
        note,
        total: sumPoints18(template.fixedPoints18),
        allowedMask: template.fixedPoints18.map(value => value > 0),
        fixedPoints18: clonePoints18(template.fixedPoints18),
        message: template.note
      };
    }
    case 'custom': {
      const total = asInt(form.customTotal);
      return {
        type: 'custom',
        name: label || 'その他',
        note,
        total,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: '総点を決めた上で 18 観点へ自由配点できます。'
      };
    }
    default:
      return {
        type: 'custom',
        name: label || 'その他',
        note,
        total: 0,
        allowedMask: Array(18).fill(true),
        fixedPoints18: null,
        message: '入力してください。'
      };
  }
}

function scoreLanguagePoints(kind, score) {
  const value = asInt(score);
  if (kind === 'toefl') {
    if (value >= 100) return 100;
    if (value >= 80) return 80;
    if (value >= 60) return 60;
    if (value >= 40) return 40;
    return 0;
  }
  if (value >= 860) return 100;
  if (value >= 730) return 80;
  if (value >= 600) return 60;
  if (value >= 470) return 40;
  return 0;
}

function recomputeExtraPreview() {
  const descriptor = getExtraDescriptor();

  if (descriptor.fixedPoints18) {
    state.extraForm.points18 = clonePoints18(descriptor.fixedPoints18);
    updateExtraPreviewInputs(descriptor);
    updateExtraMeta();
    return;
  }

  const total = descriptor.total;
  const allowedMask = descriptor.allowedMask || Array(18).fill(true);
  if (total <= 0) {
    state.extraForm.points18 = clonePoints18(ZERO_18);
    updateExtraPreviewInputs(descriptor);
    updateExtraMeta();
    return;
  }

  const positiveDeficits18 = getPositiveDeficits18({ excludingExtraId: state.extraForm.editingId });
  const currentValues = normalizePoints18AgainstMask(state.extraForm.points18, allowedMask);
  let nextValues;
  if (state.extraForm.splitMode === 'even') {
    nextValues = allocateByWeights(total, allowedMask.map(flag => flag ? 1 : 0));
  } else if (state.extraForm.splitMode === 'required') {
    nextValues = allocateByWeights(total, REQUIRED_POINTS18.map((value, index) => allowedMask[index] ? value : 0));
  } else {
    nextValues = allocateComposite(total, currentValues, positiveDeficits18, REQUIRED_POINTS18, allowedMask);
  }
  state.extraForm.points18 = clonePoints18(nextValues);
  updateExtraPreviewInputs(descriptor);
  updateExtraMeta();
}

function normalizePoints18AgainstMask(values, allowedMask) {
  return Array.from({ length: 18 }, (_, index) => allowedMask[index] ? asInt(values?.[index]) : 0);
}

function updateExtraPreviewInputs(descriptor = getExtraDescriptor()) {
  const disabledMask = descriptor.fixedPoints18
    ? Array(18).fill(true)
    : Array.from({ length: 18 }, (_, index) => !descriptor.allowedMask[index]);

  [...els.extraPreviewGrid.children].forEach((wrapper, index) => {
    const input = wrapper.querySelector('input');
    input.value = String(asInt(state.extraForm.points18[index]));
    input.disabled = disabledMask[index];
    wrapper.classList.toggle('is-locked', disabledMask[index]);
  });
  updateExtraPreviewTotal();
}

function updateExtraPreviewTotal() {
  const previewTotal = sumPoints18(state.extraForm.points18);
  const descriptor = getExtraDescriptor();
  const diff = previewTotal - descriptor.total;
  const suffix = diff === 0 ? '' : ` / 目標との差 ${formatSigned(diff)}`;
  els.extraPreviewTotal.textContent = `合計 ${previewTotal} 点${suffix}`;
}

function updateExtraMeta() {
  const descriptor = getExtraDescriptor();
  const previewTotal = sumPoints18(state.extraForm.points18);
  let kind = 'muted';
  let message = descriptor.message;

  if (descriptor.total <= 0) {
    kind = 'warn';
    message = `${descriptor.message} 現在の総点は 0 点です。`;
  } else if (previewTotal === descriptor.total) {
    kind = 'ok';
  } else {
    kind = 'pending';
    message = `${descriptor.message} 入力合計 ${previewTotal} 点を保存時に ${descriptor.total} 点へ正規化します。`;
  }

  if (state.extraForm.editingId) {
    message = `編集中: ${message}`;
  }

  els.extraMeta.textContent = message;
  els.extraMeta.className = `status status--${kind}`;
}

function normalizePreviewToTargetTotal(values, total, allowedMask) {
  const masked = normalizePoints18AgainstMask(values, allowedMask);
  const currentTotal = sumPoints18(masked);
  if (total <= 0) {
    return clonePoints18(ZERO_18);
  }
  if (currentTotal === total) {
    return clonePoints18(masked);
  }
  if (currentTotal <= 0) {
    return allocateByWeights(total, allowedMask.map(flag => flag ? 1 : 0));
  }
  return allocateByWeights(total, masked);
}

function addExtraItemToDraft() {
  const descriptor = getExtraDescriptor();
  if (!descriptor.name.trim()) {
    setWorkbookStatus('授業科目以外の表示名を入力してください。', 'warn');
    return;
  }
  if (!(descriptor.total > 0)) {
    setWorkbookStatus('授業科目以外の総点が 0 点です。条件を見直してください。', 'warn');
    return;
  }

  const normalizedPoints18 = normalizePreviewToTargetTotal(state.extraForm.points18, descriptor.total, descriptor.allowedMask);
  const items = getCurrentDraft().extraItems;
  const existingIndex = items.findIndex(item => item.id === state.extraForm.editingId);
  const item = {
    id: existingIndex >= 0 ? items[existingIndex].id : state.nextExtraId++,
    type: descriptor.type,
    name: descriptor.name,
    note: descriptor.note,
    total: descriptor.total,
    points18: clonePoints18(normalizedPoints18),
    formSnapshot: deepClone({
      ...state.extraForm,
      points18: clonePoints18(normalizedPoints18),
      editingId: null
    })
  };

  if (existingIndex >= 0) {
    items[existingIndex] = item;
  } else {
    items.push(item);
  }

  clearExtraForm();
  renderPlanSection();
  setWorkbookStatus(`${MODE_CONFIG[state.currentMode].label}の授業科目以外を更新しました。`, currentStatusKindAfterChange());
}

function clearExtraForm() {
  state.extraForm = createDefaultExtraForm();
  els.extraType.value = state.extraForm.type;
  renderExtraDynamicFields();
  recomputeExtraPreview();
}

function renderExtraList() {
  const items = getCurrentDraft().extraItems || [];

  if (!items.length) {
    els.extraList.innerHTML = '<div class="status status--muted">まだ授業科目以外はありません。上のフォームから追加してください。</div>';
    return;
  }

  els.extraList.innerHTML = items.map(item => {
    const tags = [
      chip(EXTRA_TYPE_LABELS[item.type] || '授業科目以外'),
      chip(`${item.total} 点`),
      item.note ? chip(item.note) : ''
    ].join('');

    return `
      <article class="plan-card">
        <div class="plan-card__top">
          <div>
            <div class="plan-card__row">授業科目以外</div>
            <h3>${escapeHtml(item.name)}</h3>
            <p>${escapeHtml(item.note || '')}</p>
            <div class="plan-card__meta">${tags}</div>
          </div>
          <div class="plan-card__total">${item.total} 点</div>
        </div>
        <div class="actions-row actions-row--compact">
          <button class="button button--small" type="button" data-action="edit-extra" data-id="${item.id}">編集</button>
          <button class="button button--small button--danger" type="button" data-action="remove-extra" data-id="${item.id}">削除</button>
        </div>
      </article>
    `;
  }).join('');

  els.extraList.querySelectorAll('button[data-action="edit-extra"]').forEach(button => {
    button.addEventListener('click', () => editExtraItem(Number(button.dataset.id)));
  });
  els.extraList.querySelectorAll('button[data-action="remove-extra"]').forEach(button => {
    button.addEventListener('click', () => removeExtraItem(Number(button.dataset.id)));
  });
}

function editExtraItem(id) {
  const item = (getCurrentDraft().extraItems || []).find(entry => entry.id === id);
  if (!item) return;

  state.extraForm = item.formSnapshot ? deepClone(item.formSnapshot) : createDefaultExtraForm();
  state.extraForm.editingId = id;
  state.extraForm.label = item.name;
  state.extraForm.note = item.note || '';
  state.extraForm.points18 = clonePoints18(item.points18);
  els.extraType.value = state.extraForm.type;
  renderExtraDynamicFields();
  updateExtraPreviewInputs(getExtraDescriptor());
  updateExtraMeta();
  updateActionStates();
  setWorkbookStatus('授業科目以外の内容を編集フォームへ読み込みました。', currentStatusKindAfterChange());
}

function removeExtraItem(id) {
  const draft = getCurrentDraft();
  draft.extraItems = draft.extraItems.filter(item => item.id !== id);
  if (state.extraForm.editingId === id) {
    clearExtraForm();
  }
  renderPlanSection();
  setWorkbookStatus('授業科目以外の項目を削除しました。', currentStatusKindAfterChange());
}

function renderPlanSection() {
  renderSyncStatus();
  renderPlanList();
  renderExtraList();
  renderTotalsTable();
  updateActionStates();
}

function renderSyncStatus() {
  const config = MODE_CONFIG[state.currentMode];

  if (state.currentMode === 'plan') {
    els.syncStatus.innerHTML = '<span class="badge">履修計画</span><span class="muted"> このモードは UI 上の検討用です。Excel への書き込みは行いません。</span>';
    return;
  }
  if (!state.workbook) {
    els.syncStatus.innerHTML = '<span class="badge">Excel 未読込</span><span class="muted"> まだ一括反映先の Excel がありません。</span>';
    return;
  }
  if (isModeDraftSyncedWithWorkbook(state.currentMode)) {
    els.syncStatus.innerHTML = `<span class="badge badge--ok">同期済み</span><span class="muted"> ${config.label} のドラフトと Excel シートが一致しています。</span>`;
    return;
  }
  els.syncStatus.innerHTML = `<span class="badge badge--pending">未反映</span><span class="muted"> ${config.label} のドラフトに未反映の変更があります。最後に一括反映してください。</span>`;
}

function renderPlanList() {
  const entries = getPlanEntries(getCurrentDraft().courseRows);
  if (!entries.length) {
    els.planList.innerHTML = '<div class="status status--muted">まだ授業科目ドラフトはありません。科目を選んで追加してください。</div>';
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

  const draft = getCurrentDraft();
  const metrics = getMetrics(draft);
  els.planCourseCount.textContent = String(entries.length);
  els.planCredits.textContent = String(entries.reduce((sum, entry) => sum + Number(entry.credits || 0), 0));
  els.planPointsTotal.textContent = String(metrics.subtotalTotal);
  els.extraItemCount.textContent = String((draft.extraItems || []).length);
  els.extraPointsTotal.textContent = String(metrics.extraTotal);
  els.projectedTotal.textContent = String(metrics.projectedTotal);
  els.unmetCount.textContent = `${metrics.unmetCount} / 18`;
}

function movePlanRow(row, delta) {
  const targetRow = row + delta;
  if (targetRow < 11) return;

  const draft = getCurrentDraft();
  const entries = clonePlanRows(draft.courseRows);
  const current = entries[row];
  if (!current) return;

  const swap = entries[targetRow];
  if (swap) {
    entries[row] = { ...swap, row };
  } else {
    delete entries[row];
  }
  entries[targetRow] = { ...current, row: targetRow };

  draft.courseRows = entries;
  renderRowOptions(targetRow);
  if (state.selectedCourse) {
    recomputePreview();
  }
  renderPlanSection();
  setWorkbookStatus('授業科目ドラフトの並びを更新しました。', currentStatusKindAfterChange());
}

function removePlanRow(row) {
  delete getCurrentDraft().courseRows[row];
  renderRowOptions();
  if (state.selectedCourse) {
    recomputePreview();
  }
  renderPlanSection();
  setWorkbookStatus(`${row}行をドラフトから削除しました。`, currentStatusKindAfterChange());
}

function editPlanRow(row) {
  const entry = getCurrentDraft().courseRows[row];
  if (!entry) return;

  state.selectedCourse = entry.sourceCourse ? cloneCourseSource(entry.sourceCourse) : {
    code: entry.code,
    nameJa: entry.nameJa,
    nameEn: entry.nameEn,
    credits: entry.credits,
    total: sumPoints18(entry.points18),
    points18: clonePoints18(entry.points18),
    sourceKind: 'manual',
    aliases: [entry.nameJa, entry.nameEn].filter(Boolean)
  };
  state.previewPoints18 = clonePoints18(entry.points18);
  renderSelectedCourse(state.selectedCourse);
  renderRowOptions(row);
  updatePreviewInputs();
  updateActionStates();
  setWorkbookStatus(`${row}行の内容を編集フォームへ読み込みました。`, currentStatusKindAfterChange());
}

function renderTotalsTable() {
  const metrics = getMetrics(getCurrentDraft());

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
  const config = MODE_CONFIG[state.currentMode];
  const hasWorkbook = Boolean(state.workbook);
  const hasSelection = Boolean(state.selectedCourse);
  const hasDraft = hasDraftContent();
  const synced = config.canWrite ? isModeDraftSyncedWithWorkbook(state.currentMode) : true;

  els.planAddBtn.disabled = !hasSelection;
  els.applyPlanBtn.disabled = !config.canWrite || !hasWorkbook || synced;
  els.reloadPlanBtn.disabled = !config.canWrite || !hasWorkbook;
  els.clearPlanBtn.disabled = !hasDraft;
  els.downloadBtn.disabled = !hasWorkbook || (config.canWrite && !synced);
  els.seedFinalBtn.disabled = state.currentMode !== 'final' || !hasDraftContent(getDraft('midterm'));
}

function getSheetNameForMode(mode) {
  return state.availableSheets?.[mode] || null;
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
    state.availableSheets.midterm = detectSheetNameForMode(state.workbook, 'midterm');
    state.availableSheets.final = detectSheetNameForMode(state.workbook, 'final');

    const importedMidterm = reloadModeFromWorkbook('midterm', { silent: true });
    const importedFinal = reloadModeFromWorkbook('final', { silent: true, allowSeedFromMidterm: true });
    const baseMode = hasDraftContent(getDraft('midterm')) ? 'midterm' : (hasDraftContent(getDraft('final')) ? 'final' : null);
    state.modeDrafts.plan = baseMode ? cloneModeDraft(getDraft(baseMode)) : createEmptyModeDraft();

    maybeSeedFinalFromMidterm();
    renderModeInfo();
    renderRowOptions();
    renderPlanSection();
    updateActionStates();

    const currentDraft = getDraft(state.currentMode);
    const finalSource = getDraft('final').seededFromMidterm
      ? ' / 最終評価は中間評価ドラフトを初期値にしています。'
      : '';
    setWorkbookStatus(`読込完了: ${file.name} / M中間評価 ${importedMidterm} 行 / M最終評価 ${importedFinal} 行 / 現在モードの授業科目以外 ${currentDraft.extraItems.length} 件${finalSource}`, 'ok');
  } catch (error) {
    console.error(error);
    state.workbook = null;
    state.workbookFileName = '';
    state.availableSheets.midterm = null;
    state.availableSheets.final = null;
    state.modeDrafts = {
      plan: createEmptyModeDraft(),
      midterm: createEmptyModeDraft(),
      final: createEmptyModeDraft()
    };
    renderRowOptions();
    renderPlanSection();
    updateActionStates();
    setWorkbookStatus('Excel の読込に失敗しました。ファイル形式を確認してください。', 'warn');
  }
}

function detectSheetNameForMode(workbook, mode) {
  const names = workbook.sheets().map(sheet => sheet.name());
  return MODE_CONFIG[mode].sheetCandidates.find(name => names.includes(name)) || null;
}

function detectSheetLayout(sheet) {
  const scanMax = 800;
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

function importDraftFromSheet(sheet) {
  const layout = detectSheetLayout(sheet);
  const courseRows = {};
  const extraItems = [];

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

    courseRows[row] = {
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

  for (let row = layout.extraStartRow; row <= layout.extraEndRow; row += 1) {
    const name = String(sheet.cell(`A${row}`).value() || '').trim();
    const points18 = COLUMN_LETTERS.map(letter => asInt(sheet.cell(`${letter}${row}`).value()));
    const hasAnyPoints = points18.some(value => value > 0);
    if (!name && !hasAnyPoints) continue;
    if (!hasAnyPoints) continue;

    const templateKey = guessExtraTemplateKey(name, points18);
    const snapshot = createDefaultExtraForm();
    if (templateKey) {
      snapshot.type = 'fixed';
      snapshot.fixedKey = templateKey;
    } else {
      snapshot.type = 'custom';
      snapshot.customTotal = sumPoints18(points18);
    }
    snapshot.label = name;
    snapshot.points18 = clonePoints18(points18);

    extraItems.push({
      id: state.nextExtraId++,
      type: templateKey ? 'fixed' : 'custom',
      name,
      note: '',
      total: sumPoints18(points18),
      points18: clonePoints18(points18),
      formSnapshot: snapshot
    });
  }

  return {
    courseRows,
    extraItems
  };
}

function guessExtraTemplateKey(name, points18) {
  const normalizedName = normalize(name);
  const matchedByName = Object.values(FIXED_EXTRA_TEMPLATES).find(template => normalize(template.label) === normalizedName);
  if (matchedByName) return matchedByName.key;
  const matchedByPoints = Object.values(FIXED_EXTRA_TEMPLATES).find(template => points18.every((value, index) => asInt(value) === asInt(template.fixedPoints18[index])));
  return matchedByPoints?.key || null;
}

function reloadModeFromWorkbook(mode = state.currentMode, options = {}) {
  const { silent = false, allowSeedFromMidterm = false } = options;
  if (!state.workbook) return 0;
  const sheetName = getSheetNameForMode(mode);
  if (!sheetName) {
    state.modeDrafts[mode] = createEmptyModeDraft();
    return 0;
  }

  const sheet = state.workbook.sheet(sheetName);
  const imported = importDraftFromSheet(sheet);
  state.modeDrafts[mode] = {
    courseRows: imported.courseRows,
    extraItems: imported.extraItems,
    seededFromWorkbook: true,
    seededFromMidterm: false
  };

  if (mode === 'final' && allowSeedFromMidterm && !hasDraftContent(state.modeDrafts.final) && hasDraftContent(getDraft('midterm'))) {
    state.modeDrafts.final = cloneModeDraft(getDraft('midterm'));
    state.modeDrafts.final.seededFromMidterm = true;
  }

  if (mode === state.currentMode) {
    renderModeInfo();
    renderRowOptions();
    if (state.selectedCourse) {
      recomputePreview();
    }
    recomputeExtraPreview();
    renderPlanSection();
    updateActionStates();
    if (!silent) {
      setWorkbookStatus(`${MODE_CONFIG[mode].label} を Excel から再読込しました。`, currentStatusKindAfterChange());
    }
  }

  return getPlanEntries(state.modeDrafts[mode].courseRows).length;
}

function reloadCurrentModeFromWorkbook() {
  reloadModeFromWorkbook(state.currentMode, { allowSeedFromMidterm: state.currentMode === 'final' });
}

function isModeDraftSyncedWithWorkbook(mode) {
  if (mode === 'plan') return true;
  if (!state.workbook) return false;
  const sheetName = getSheetNameForMode(mode);
  if (!sheetName) return false;

  try {
    const draft = getDraft(mode);
    const sheet = state.workbook.sheet(sheetName);
    const layout = detectSheetLayout(sheet);
    const expectedEnd = expectedCourseEndRow(draft.courseRows);
    if (layout.courseEndRow !== expectedEnd) {
      return false;
    }

    for (let row = 11; row <= expectedEnd; row += 1) {
      const entry = draft.courseRows[row];
      const expectedName = entry ? (entry.nameJa || entry.nameEn || entry.code || '') : getDefaultClearName(row);
      const expectedPoints18 = entry ? clonePoints18(entry.points18) : clonePoints18(ZERO_18);
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

    const expectedExtraRows = Math.max(MIN_EXTRA_ROWS, draft.extraItems.length || 0);
    const actualExtraRows = layout.extraEndRow - layout.extraStartRow + 1;
    if (actualExtraRows !== expectedExtraRows) {
      return false;
    }

    for (let i = 0; i < expectedExtraRows; i += 1) {
      const row = layout.extraStartRow + i;
      const item = draft.extraItems[i];
      const expectedName = item ? item.name : '';
      const expectedPoints18 = item ? clonePoints18(item.points18) : clonePoints18(ZERO_18);
      const actualName = String(sheet.cell(`A${row}`).value() || '').trim();
      const actualPoints18 = COLUMN_LETTERS.map(letter => asInt(sheet.cell(`${letter}${row}`).value()));
      if (normalize(actualName) !== normalize(expectedName)) {
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

function expectedCourseEndRow(courseRows = getCurrentDraft().courseRows) {
  return Math.max(MIN_COURSE_END_ROW, getHighestPlannedRow(courseRows) || MIN_COURSE_END_ROW);
}

function currentStatusKindAfterChange() {
  if (state.currentMode === 'plan') return 'muted';
  if (!state.workbook) return 'muted';
  return isModeDraftSyncedWithWorkbook(state.currentMode) ? 'ok' : 'pending';
}

function clearCurrentDraft() {
  const empty = createEmptyModeDraft();
  getCurrentDraft().courseRows = empty.courseRows;
  getCurrentDraft().extraItems = empty.extraItems;
  getCurrentDraft().seededFromMidterm = false;
  renderRowOptions();
  if (state.selectedCourse) {
    recomputePreview();
  }
  clearExtraForm();
  renderPlanSection();
  updateActionStates();
  setWorkbookStatus(`${MODE_CONFIG[state.currentMode].label} のドラフトをクリアしました。`, currentStatusKindAfterChange());
}

function prepareRowTemplates(sheet, layout) {
  const courseStyleRow = captureStyleRow(sheet, Math.min(MIN_COURSE_END_ROW, layout.courseEndRow));
  const extraStyleRow = captureStyleRow(sheet, layout.extraStartRow);
  return {
    courseStyleRow,
    courseSubtotalStyle: captureStyleRow(sheet, layout.courseSubtotalRow),
    blank1Style: captureStyleRow(sheet, layout.blankAfterCourseSubtotalRow),
    extraHeaderStyle: captureStyleRow(sheet, layout.extraHeaderRow),
    extraItemStyle: extraStyleRow,
    extraSubtotalStyle: captureStyleRow(sheet, layout.extraSubtotalRow),
    blank2Style: captureStyleRow(sheet, layout.blankAfterExtraSubtotalRow),
    totalStyle: captureStyleRow(sheet, layout.totalRow),
    deficitStyle: captureStyleRow(sheet, layout.deficitRow)
  };
}

function applyCurrentModeToWorkbook() {
  const config = MODE_CONFIG[state.currentMode];
  if (!config.canWrite || !state.workbook) return;
  const sheetName = getSheetNameForMode(state.currentMode);
  if (!sheetName) {
    setWorkbookStatus(`${config.label} シートが見つかりません。`, 'warn');
    return;
  }

  try {
    const sheet = state.workbook.sheet(sheetName);
    const layout = detectSheetLayout(sheet);
    const templates = prepareRowTemplates(sheet, layout);
    const draft = getCurrentDraft();
    const targetCourseEndRow = expectedCourseEndRow(draft.courseRows);
    const oldDeficitRow = layout.deficitRow;

    ensureCourseRows(sheet, 11, targetCourseEndRow, layout.courseEndRow, templates.courseStyleRow);
    for (let row = 11; row <= targetCourseEndRow; row += 1) {
      clearCourseRow(sheet, row);
    }
    getPlanEntries(draft.courseRows).forEach(entry => {
      writePoints18ToRow(sheet, entry.row, entry.nameJa || entry.nameEn || entry.code || '', entry.points18);
    });

    const newLayout = writeFooterStructure(sheet, targetCourseEndRow, templates, draft.extraItems);
    clearTrailingArea(sheet, newLayout.deficitRow + 1, oldDeficitRow);

    renderRowOptions();
    renderModeInfo();
    renderPlanSection();
    updateActionStates();
    setWorkbookStatus(`${config.label} を Excel に一括反映しました。授業科目欄は 11〜${targetCourseEndRow} 行、授業科目以外は ${Math.max(MIN_EXTRA_ROWS, draft.extraItems.length)} 行で再配置しています。`, 'ok');
  } catch (error) {
    console.error(error);
    setWorkbookStatus('Excel への一括反映に失敗しました。テンプレート構造を確認してください。', 'warn');
  }
}

function writeFooterStructure(sheet, courseEndRow, templates, extraItems) {
  const extraRowCount = Math.max(MIN_EXTRA_ROWS, (extraItems || []).length);
  const courseSubtotalRow = courseEndRow + 1;
  const blank1Row = courseSubtotalRow + 1;
  const extraHeaderRow = courseSubtotalRow + 2;
  const extraStartRow = extraHeaderRow + 1;
  const extraEndRow = extraStartRow + extraRowCount - 1;
  const extraSubtotalRow = extraEndRow + 1;
  const blank2Row = extraSubtotalRow + 1;
  const totalRow = blank2Row + 1;
  const deficitRow = totalRow + 1;

  applyStyleRow(sheet, courseSubtotalRow, templates.courseSubtotalStyle);
  writeCourseSubtotalRow(sheet, courseSubtotalRow, courseEndRow);

  applyStyleRow(sheet, blank1Row, templates.blank1Style);
  clearRowContents(sheet, blank1Row);

  applyStyleRow(sheet, extraHeaderRow, templates.extraHeaderStyle);
  clearRowContents(sheet, extraHeaderRow);
  sheet.cell(`A${extraHeaderRow}`).value('授業科目以外');

  for (let i = 0; i < extraRowCount; i += 1) {
    const row = extraStartRow + i;
    const item = extraItems?.[i];
    applyStyleRow(sheet, row, templates.extraItemStyle);
    clearRowContents(sheet, row);
    if (item) {
      sheet.cell(`A${row}`).value(item.name || '');
      COLUMN_LETTERS.forEach((letter, index) => {
        sheet.cell(`${letter}${row}`).value(asInt(item.points18[index]));
      });
    }
    sheet.cell(`T${row}`).formula(`SUM(B${row}:S${row})`);
  }

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
    sheet.cell(`${letter}${row}`).formula(`SUM(${letter}11:${letter}${courseEndRow})`);
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

function ensureCourseRows(sheet, startRow, endRow, existingCourseEndRow, courseStyleRow) {
  for (let row = startRow; row <= endRow; row += 1) {
    if (row > existingCourseEndRow) {
      applyStyleRow(sheet, row, courseStyleRow);
    }
    clearCourseRow(sheet, row);
  }
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

function getDefaultClearName(row) {
  return DEFAULT_ROW_NAMES[row] || '';
}

function clearRowContents(sheet, row) {
  for (let col = 1; col <= 20; col += 1) {
    sheet.cell(`${columnLetter(col)}${row}`).value('');
  }
}

function clearTrailingArea(sheet, fromRow, toRow) {
  if (toRow < fromRow) return;
  for (let row = fromRow; row <= toRow; row += 1) {
    clearRowContents(sheet, row);
  }
}

async function downloadWorkbook() {
  if (!state.workbook) return;
  if (MODE_CONFIG[state.currentMode].canWrite && !isModeDraftSyncedWithWorkbook(state.currentMode)) {
    setWorkbookStatus('先に一括反映してからダウンロードしてください。', 'warn');
    return;
  }

  try {
    const blob = await state.workbook.outputAsync();
    const link = document.createElement('a');
    const baseName = state.workbookFileName.replace(/\.(xlsx|xlsm|xls)$/i, '');
    const modeLabel = MODE_CONFIG[state.currentMode].canWrite ? MODE_CONFIG[state.currentMode].label : '履修計画';
    const downloadName = `${baseName || '達成度評価'}_${modeLabel}_入力支援.xlsx`;
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

function displayCourseName(entry) {
  return String(entry?.nameJa || entry?.nameEn || entry?.code || '名称未設定');
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

function normalizeMaskedWeights(values, mask) {
  const arr = Array.from({ length: mask.length }, (_, index) => {
    if (!mask[index]) return 0;
    return Math.max(0, Number(values?.[index] || 0));
  });
  const total = arr.reduce((sum, value) => sum + value, 0);
  if (total <= 0) return arr;
  return arr.map(value => value / total);
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
