/**
 * 폐기불량 관리시스템 - JavaScript (Flask 버전)
 */

// 전역 상태 관리
const state = {
  currentStep: 1,
  department: '',
  part: '',
  process: '',
  person: '',
  isAdmin: false,
  adminPassword: '',
  currentNumpadTarget: ''
};

// 입력된 폐기 항목 목록
let scrapEntries = [];

// 폐기사유 목록 캐시
let scrapReasons = [];

// 현재 편집 중인 마스터 데이터 시트
let currentMasterSheet = '';

// 현재 표시 중인 테이블 데이터 (수정용)
let currentTableHeaders = [];
let currentTableRows = [];

// 드롭다운용 캐시 데이터
let cachedDepartments = [];
let cachedProcesses = [];

// 추가 모달용
let currentAddSheet = '';
let currentAddLabel = '';

// TM-NO 검색 debounce용
let tmnoSearchTimer = null;
let tmnoCache = null;

// 마스터 데이터 캐시 (서버 왕복 최소화)
let masterCache = {
  departments: null,
  persons: null,
  scrapReasons: null
};

// ==================== API 호출 함수 ====================

async function apiCall(url, options = {}) {
  try {
    const response = await fetch(url, {
      headers: {
        'Content-Type': 'application/json',
        ...options.headers
      },
      ...options
    });
    return await response.json();
  } catch (error) {
    console.error('API 호출 오류:', error);
    showToast('서버 연결 오류', 'error');
    return null;
  }
}

// ==================== 화면 전환 ====================

function startInput() {
  document.getElementById('startScreen').classList.remove('active');
  document.getElementById('inputScreen').classList.add('active');
  updateProgress();
  history.pushState({ screen: 'input', step: 1 }, '');
}

function goToStart() {
  resetState();
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('startScreen').classList.add('active');
  history.pushState({ screen: 'start' }, '');
}

function resetState() {
  state.currentStep = 1;
  state.department = '';
  state.part = '';
  state.process = '';
  state.person = '';
  scrapEntries = [];
  tmnoCache = null;

  document.querySelectorAll('.step-container').forEach(c => c.classList.remove('active'));
  document.getElementById('step1').classList.add('active');
  updateProgress();
  updateEntriesTable();
}

// ==================== 진행 단계 ====================

function updateProgress() {
  document.querySelectorAll('.progress-step').forEach((step, idx) => {
    step.classList.remove('active', 'completed');
    const stepNum = idx + 1;
    if (stepNum < state.currentStep) {
      step.classList.add('completed');
    } else if (stepNum === state.currentStep) {
      step.classList.add('active');
    }
  });
}

function goToStep(step) {
  state.currentStep = step;
  document.querySelectorAll('.step-container').forEach(c => c.classList.remove('active'));
  document.getElementById('step' + step).classList.add('active');
  updateProgress();
  updateHeaderPrevBtn();
  history.pushState({ screen: 'input', step: step }, '');

  // Step 5 진입 시 선택 요약 표시 + 설비 드롭다운 로드
  if (step === 5) {
    updateSelectionSummary();
    loadMachinesForEntry();
  }
}

function prevStep(current) {
  goToStep(current - 1);
}

function headerPrevStep() {
  if (state.currentStep > 1) {
    prevStep(state.currentStep);
  }
}

function updateHeaderPrevBtn() {
  const btn = document.getElementById('headerPrevBtn');
  if (btn) {
    btn.style.display = state.currentStep > 1 ? 'inline-block' : 'none';
  }
}

// ==================== 선택 요약 (Step 6) ====================

function updateSelectionSummary() {
  const el = document.getElementById('selectionSummary');
  if (!el) return;
  el.innerHTML = `
    <span class="summary-item"><b>Part:</b> ${escapeHtml(state.part)}</span>
    <span class="summary-item"><b>부서:</b> ${escapeHtml(state.department)}</span>
    <span class="summary-item"><b>공정:</b> ${escapeHtml(state.process)}</span>
    <span class="summary-item"><b>폐기자:</b> ${escapeHtml(state.person)}</span>
  `;
}

// ==================== 버튼 생성 유틸리티 ====================

function createSelectButton(text, onClick) {
  const btn = document.createElement('button');
  btn.className = 'btn btn-select';
  btn.textContent = text;
  btn.addEventListener('click', onClick);
  return btn;
}

// ==================== 데이터 로드 ====================

async function loadDepartments() {
  if (masterCache.departments) {
    renderDepartmentButtons(masterCache.departments);
    return;
  }
  showLoading();
  const data = await apiCall('/api/departments');
  hideLoading();
  if (data) {
    masterCache.departments = data;
    renderDepartmentButtons(data);
  }
}

function renderDepartmentButtons(departments) {
  const container = document.getElementById('departmentList');
  container.innerHTML = '';
  departments.forEach(dept => {
    container.appendChild(createSelectButton(dept, () => selectDepartment(dept)));
  });
}

async function loadProcesses() {
  showLoading();
  const data = await apiCall(`/api/processes?part=${encodeURIComponent(state.part)}`);
  hideLoading();

  if (data) {
    const container = document.getElementById('processList');
    container.innerHTML = '';
    data.forEach(proc => {
      container.appendChild(createSelectButton(proc, () => selectProcess(proc)));
    });
  }
}

// 폐기입력 폼의 설비 드롭다운 로드
async function loadMachinesForEntry() {
  const select = document.getElementById('newMachine');
  if (!select) return;

  const data = await apiCall(`/api/machines?part=${encodeURIComponent(state.part)}&process=${encodeURIComponent(state.process)}`);

  const prevValue = select.value;
  select.innerHTML = '<option value="">선택 안함</option>';
  if (data) {
    data.forEach(machine => {
      const opt = document.createElement('option');
      opt.value = machine;
      opt.textContent = machine;
      select.appendChild(opt);
    });
  }
  // 이전 선택값 유지
  if (prevValue) select.value = prevValue;
}

// 폐기자 이름 캐시
let cachedPersonNames = [];

async function loadPersons() {
  if (masterCache.persons) {
    cachedPersonNames = masterCache.persons;
    renderPersonButtons(cachedPersonNames);
    document.getElementById('personInput').value = '';
    return;
  }
  showLoading();
  const data = await apiCall('/api/persons');
  hideLoading();
  if (data) {
    masterCache.persons = data;
    cachedPersonNames = data;
    renderPersonButtons(data);
  }
  document.getElementById('personInput').value = '';
}

function renderPersonButtons(names) {
  const container = document.getElementById('personList');
  container.innerHTML = '';
  names.forEach(person => {
    container.appendChild(createSelectButton(person, () => selectPerson(person)));
  });
}

function filterPersons() {
  const keyword = document.getElementById('personInput').value.trim();
  if (!keyword) {
    renderPersonButtons(cachedPersonNames);
    return;
  }
  const filtered = cachedPersonNames.filter(name => name.includes(keyword));
  renderPersonButtons(filtered);
}

// ==================== 선택 함수 ====================

function selectPart(part) {
  state.part = part;
  tmnoCache = null;
  loadDepartments();
  goToStep(2);
}

function selectDepartment(dept) {
  state.department = dept;
  loadPersons();
  goToStep(3);
}

function selectPerson(person) {
  state.person = person;
  loadProcesses();
  goToStep(4);
}

function selectProcess(proc) {
  state.process = proc;
  tmnoCache = null;
  goToStep(5);
  clearNewEntryForm();
}

async function confirmPersonInput() {
  const name = document.getElementById('personInput').value.trim();
  if (!name) {
    showToast('이름을 입력해주세요.', 'error');
    return;
  }
  // 서버에 저장 (이미 있으면 무시됨)
  await apiCall('/api/master_data/person', {
    method: 'POST',
    body: JSON.stringify({ data: [name] })
  });
  // 캐시에 즉시 추가 (서버 재호출 불필요)
  if (masterCache.persons && !masterCache.persons.includes(name)) {
    masterCache.persons.push(name);
    masterCache.persons.sort();
    cachedPersonNames = masterCache.persons;
  }
  selectPerson(name);
}

// ==================== TM-NO 검색 (debounce 적용) ====================

function searchNewTMNO() {
  clearTimeout(tmnoSearchTimer);
  tmnoSearchTimer = setTimeout(doSearchTMNO, 300);
}

async function doSearchTMNO() {
  const keyword = document.getElementById('newTmnoSearch').value.trim();
  const dropdown = document.getElementById('tmnoDropdown');

  if (keyword.length < 1) {
    dropdown.classList.remove('active');
    document.getElementById('newTmno').value = '';
    document.getElementById('newProductName').value = '';
    document.getElementById('newUnitWeight').value = '';
    enableQuantityInput(false);
    return;
  }

  // TM-NO 목록 캐시 (Part+공정 조합당 1회만 서버 호출)
  if (!tmnoCache) {
    tmnoCache = await apiCall(`/api/tmnos?part=${encodeURIComponent(state.part)}&process=${encodeURIComponent(state.process)}`);
  }

  if (tmnoCache) {
    const filtered = tmnoCache.filter(tmno =>
      String(tmno).toUpperCase().includes(keyword.toUpperCase())
    ).slice(0, 10);

    if (filtered.length > 0) {
      dropdown.innerHTML = '';
      filtered.forEach(tmno => {
        const item = document.createElement('div');
        item.className = 'tmno-dropdown-item';
        item.textContent = tmno;
        item.addEventListener('click', () => selectNewTMNO(tmno));
        dropdown.appendChild(item);
      });
      dropdown.classList.add('active');
    } else {
      dropdown.classList.remove('active');
    }
  }
}

async function selectNewTMNO(tmno) {
  document.getElementById('newTmnoSearch').value = tmno;
  document.getElementById('newTmno').value = tmno;
  document.getElementById('tmnoDropdown').classList.remove('active');

  const info = await apiCall(`/api/tmno_info?part=${encodeURIComponent(state.part)}&tmno=${encodeURIComponent(tmno)}`);

  if (info) {
    document.getElementById('newProductName').value = info.productName || '';
    document.getElementById('newUnitWeight').value = info.unitWeight || 0;
    enableQuantityInput(true);
  }
}

// 수량 입력 활성화/비활성화
function enableQuantityInput(enabled) {
  const qtyInput = document.getElementById('newQuantity');
  if (enabled) {
    qtyInput.classList.remove('disabled-input');
    qtyInput.onclick = function() { showNumpad('newQuantity'); };
    qtyInput.style.cursor = 'pointer';
    qtyInput.style.opacity = '1';
  } else {
    qtyInput.classList.add('disabled-input');
    qtyInput.onclick = null;
    qtyInput.style.cursor = 'not-allowed';
    qtyInput.style.opacity = '0.5';
    qtyInput.value = '';
  }
}

// ==================== 폐기사유 선택 ====================

async function showReasonSelector() {
  const reasons = masterCache.scrapReasons;
  if (reasons) {
    scrapReasons = reasons;
    renderReasonButtons(reasons);
    document.getElementById('reasonModal').classList.add('active');
    return;
  }
  showLoading();
  const data = await apiCall('/api/scrap_reasons');
  hideLoading();
  if (data) {
    scrapReasons = data;
    masterCache.scrapReasons = data;
    renderReasonButtons(data);
    document.getElementById('reasonModal').classList.add('active');
  }
}

function renderReasonButtons(reasons) {
  const container = document.getElementById('reasonList');
  container.innerHTML = '';

  const mainReasons = ['공정불량', '셋팅불량'];
  const mainItems = reasons.filter(r => mainReasons.includes(r));
  const etcItems = reasons.filter(r => !mainReasons.includes(r));

  // 공정불량/셋팅불량 그룹 박스
  const mainBox = document.createElement('div');
  mainBox.className = 'reason-group-box reason-group-main';
  const mainGrid = document.createElement('div');
  mainGrid.className = 'button-grid reason-main-grid';
  mainReasons.forEach(name => {
    if (mainItems.includes(name)) {
      mainGrid.appendChild(createSelectButton(name, () => selectReason(name)));
    }
  });
  mainBox.appendChild(mainGrid);
  container.appendChild(mainBox);

  // 기타폐기 그룹 박스
  const etcBox = document.createElement('div');
  etcBox.className = 'reason-group-box reason-group-etc';
  const etcTitle = document.createElement('div');
  etcTitle.className = 'reason-group-title';
  etcTitle.textContent = '기타폐기';
  etcBox.appendChild(etcTitle);
  const etcGrid = document.createElement('div');
  etcGrid.className = 'button-grid reason-etc-grid';
  etcItems.forEach(reason => {
    etcGrid.appendChild(createSelectButton(reason, () => selectReason(reason)));
  });
  etcBox.appendChild(etcGrid);
  // + 폐기사유 추가 버튼을 기타폐기 박스 안에
  const addBtn = document.createElement('button');
  addBtn.className = 'btn btn-add-small';
  addBtn.textContent = '+ 폐기사유 추가';
  addBtn.onclick = () => showAddModal('scrap_name', '폐기사유');
  etcBox.appendChild(addBtn);
  container.appendChild(etcBox);
}

function selectReason(reason) {
  document.getElementById('newReason').value = reason;
  document.getElementById('newReasonBtn').textContent = reason;
  document.getElementById('newReasonBtn').classList.add('selected');
  closeModal('reasonModal');

  // 기타 사유인 경우 비고 입력창 표시
  const remarkContainer = document.getElementById('remarkContainer');
  if (reason.startsWith('기타')) {
    remarkContainer.style.display = 'block';
    document.getElementById('newRemark').value = '';
  } else {
    remarkContainer.style.display = 'none';
    document.getElementById('newRemark').value = '';
  }

}

// ==================== 숫자패드 ====================

function showNumpad(target) {
  state.currentNumpadTarget = target;
  const currentValue = document.getElementById(target).value || '';
  document.getElementById('numpadDisplay').value = currentValue;

  if (target === 'newQuantity') {
    document.getElementById('numpadTitle').textContent = '수량 입력';
  } else {
    document.getElementById('numpadTitle').textContent = '중량 입력 (kg)';
  }

  document.getElementById('numpadModal').classList.add('active');
}

function numpadInput(val) {
  const display = document.getElementById('numpadDisplay');
  const current = display.value;

  // 소수점 중복 방지
  if (val === '.' && current.includes('.')) return;

  // 수량 입력은 정수만 허용
  if (state.currentNumpadTarget === 'newQuantity' && val === '.') return;

  display.value = current + val;
}

function numpadDelete() {
  const display = document.getElementById('numpadDisplay');
  display.value = display.value.slice(0, -1);
}

function numpadClear() {
  document.getElementById('numpadDisplay').value = '';
}

function numpadConfirm() {
  const value = document.getElementById('numpadDisplay').value;
  const target = state.currentNumpadTarget;

  // 유효한 숫자인지 검증
  if (value && isNaN(parseFloat(value))) {
    showToast('올바른 숫자를 입력해주세요.', 'error');
    return;
  }

  document.getElementById(target).value = value;

  // 수량/중량 자동 계산
  const unitWeight = parseFloat(document.getElementById('newUnitWeight').value) || 0;
  const currentReason = document.getElementById('newReason').value;

  if (target === 'newQuantity' && unitWeight > 0) {
    const qty = parseFloat(value) || 0;
    document.getElementById('newWeight').value = (qty * unitWeight).toFixed(3);
  } else if (target === 'newWeight' && unitWeight > 0) {
    const weight = parseFloat(value) || 0;
    document.getElementById('newQuantity').value = Math.round(weight / unitWeight);
  }

  // 룰1: 1Part + 소결로_산화.이물 → 중량 입력 시 수량 = 중량 / 0.0493 (반올림 정수)
  if (target === 'newWeight' && state.part === '1Part' && currentReason === '소결로_산화.이물') {
    const weight = parseFloat(value) || 0;
    if (weight > 0) {
      document.getElementById('newQuantity').value = Math.round(weight / 0.0493);
    }
  }

  closeModal('numpadModal');
}

// ==================== 항목 테이블 ====================

function addEntryToTable() {
  const machine = document.getElementById('newMachine').value;
  const tmno = document.getElementById('newTmno').value;
  const productName = document.getElementById('newProductName').value;
  const reason = document.getElementById('newReason').value;
  const quantity = document.getElementById('newQuantity').value;
  const weight = document.getElementById('newWeight').value;
  const unitWeight = document.getElementById('newUnitWeight').value;
  const remark = document.getElementById('newRemark').value.trim();

  // 폐기사유 필수
  if (!reason) {
    showToast('폐기사유를 선택해주세요.', 'error');
    return;
  }
  // 기타 사유인 경우 비고 필수
  if (reason.startsWith('기타') && !remark) {
    showToast('기타 사유를 입력해주세요.', 'error');
    return;
  }
  // TM-NO 없으면 중량 필수, 있으면 수량 또는 중량
  if (!tmno) {
    if (!weight) {
      showToast('중량을 입력해주세요.', 'error');
      return;
    }
  } else {
    if (!quantity && !weight) {
      showToast('수량 또는 중량을 입력해주세요.', 'error');
      return;
    }
  }

  // 룰2: 기타_불량 → 불량해당=해당, 해당공정=선택한 공정(state.process)
  // 룰3: 소결로_산화.이물 → 불량해당=해당, 해당공정=소결
  let defectCategory = '';
  let defectProcessValue = '';
  if (reason === '기타_불량') {
    defectCategory = '해당';
    defectProcessValue = state.process;
  } else if (reason === '소결로_산화.이물') {
    defectCategory = '해당';
    defectProcessValue = '소결';
  }

  scrapEntries.push({
    machine: machine || '',
    tmno: tmno || '-',
    productName: productName || '-',
    reason: reason,
    quantity: parseFloat(quantity) || 0,
    weight: parseFloat(weight) || 0,
    unitWeight: parseFloat(unitWeight) || 0,
    remark: remark,
    defectCategory: defectCategory,
    defectProcess: defectProcessValue
  });

  updateEntriesTable();
  clearNewEntryForm();
  showToast('항목이 추가되었습니다.', 'success');
}

function updateEntriesTable() {
  const tbody = document.getElementById('entryTableBody');
  tbody.innerHTML = '';

  scrapEntries.forEach((entry, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${entry.machine ? escapeHtml(entry.machine) : '-'}</td>
      <td>${escapeHtml(entry.reason)}</td>
      <td>${entry.remark ? escapeHtml(entry.remark) : '-'}</td>
      <td>${escapeHtml(entry.tmno)}</td>
      <td>${escapeHtml(entry.productName)}</td>
      <td>${entry.quantity}</td>
      <td>${entry.weight}</td>
      <td></td>
    `;
    const delBtn = document.createElement('button');
    delBtn.className = 'btn-row-delete';
    delBtn.textContent = '삭제';
    delBtn.addEventListener('click', () => deleteEntry(idx));
    tr.lastElementChild.appendChild(delBtn);
    tbody.appendChild(tr);
  });

  document.getElementById('entriesCount').textContent = scrapEntries.length;
}

function deleteEntry(idx) {
  scrapEntries.splice(idx, 1);
  updateEntriesTable();
}

function clearNewEntryForm() {
  // 설비 드롭다운은 옵션을 유지하고 선택만 초기화
  const machineSelect = document.getElementById('newMachine');
  if (machineSelect) machineSelect.value = '';
  document.getElementById('newTmnoSearch').value = '';
  document.getElementById('newTmno').value = '';
  document.getElementById('newProductName').value = '';
  document.getElementById('newReason').value = '';
  document.getElementById('newRemark').value = '';
  document.getElementById('remarkContainer').style.display = 'none';
  document.getElementById('newReasonBtn').textContent = '선택';
  document.getElementById('newReasonBtn').classList.remove('selected');
  document.getElementById('newQuantity').value = '';
  document.getElementById('newWeight').value = '';
  document.getElementById('newUnitWeight').value = '';
  enableQuantityInput(false);
}

// ==================== 저장 (확인 팝업 포함) ====================

function confirmSave() {
  if (scrapEntries.length === 0) {
    showToast('저장할 항목이 없습니다.', 'error');
    return;
  }

  // 확인 모달 표시
  const summary = document.getElementById('confirmSummary');
  summary.innerHTML = `
    <p><b>${escapeHtml(state.part)}</b> / <b>${escapeHtml(state.department)}</b> / <b>${escapeHtml(state.process)}</b></p>
    <p>폐기자: <b>${escapeHtml(state.person)}</b></p>
    <p>총 <b>${scrapEntries.length}</b>건을 저장합니다.</p>
  `;
  document.getElementById('confirmModal').classList.add('active');
}

async function saveAllData() {
  closeModal('confirmModal');

  showLoading();
  let successCount = 0;

  for (const entry of scrapEntries) {
    const result = await apiCall('/api/save_scrap', {
      method: 'POST',
      body: JSON.stringify({
        department: state.department,
        part: state.part,
        process: state.process,
        machine: entry.machine || '',
        person: state.person,
        tmno: entry.tmno,
        productName: entry.productName,
        scrapReason: entry.reason,
        quantity: entry.quantity,
        weight: entry.weight,
        remark: entry.remark || '',
        defectCategory: entry.defectCategory || '',
        defectProcess: entry.defectProcess || ''
      })
    });

    if (result && result.success) {
      successCount++;
    }
  }

  hideLoading();

  if (successCount === scrapEntries.length) {
    showToast(`${successCount}건 저장 완료!`, 'success');
    scrapEntries = [];
    updateEntriesTable();
    // 저장 완료 후 처음 화면으로
    setTimeout(() => goToStart(), 1500);
  } else {
    showToast(`${successCount}/${scrapEntries.length}건 저장됨`, 'error');
  }
}

// ==================== 모달 ====================

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove('active');
}

// ==================== 추가 모달 ====================

async function showAddModal(sheetName, label) {
  currentAddSheet = sheetName;
  currentAddLabel = label;

  document.getElementById('addModalTitle').textContent = label + ' 추가';
  const fieldsContainer = document.getElementById('addModalFields');
  fieldsContainer.innerHTML = '<p style="color:#aaa;">로딩 중...</p>';

  document.getElementById('addModal').classList.add('active');

  if (sheetName === 'Depart' || sheetName === 'scrap_name') {
    buildAddForm(sheetName, [], []);
  } else if (sheetName === 'Process') {
    buildAddForm(sheetName, [], []);
  } else if (sheetName === 'machine') {
    cachedProcesses = await apiCall('/api/process_list') || [];
    buildAddForm(sheetName, [], cachedProcesses);
  } else if (sheetName === 'person') {
    buildAddForm(sheetName, [], []);
  } else if (sheetName.includes('TMNO')) {
    buildAddForm(sheetName, [], []);
  }
}

function buildAddForm(sheetName, departments, processes) {
  const fieldsContainer = document.getElementById('addModalFields');

  if (sheetName === 'Depart') {
    fieldsContainer.innerHTML = '<input type="text" class="modal-input" id="addField1" placeholder="부서명">';
  } else if (sheetName === 'Process') {
    fieldsContainer.innerHTML = `
      <select class="modal-input" id="addField1">
        <option value="">Part 선택</option>
        <option value="1Part">1Part</option>
        <option value="2Part">2Part</option>
      </select>
      <input type="text" class="modal-input" id="addField2" placeholder="공정명">
    `;
  } else if (sheetName === 'machine') {
    let processOptions = '<option value="">공정 선택</option>';
    processes.forEach(p => {
      processOptions += `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`;
    });
    fieldsContainer.innerHTML = `
      <select class="modal-input" id="addField1">
        <option value="">Part 선택</option>
        <option value="1Part">1Part</option>
        <option value="2Part">2Part</option>
      </select>
      <select class="modal-input" id="addField2">${processOptions}</select>
      <input type="text" class="modal-input" id="addField3" placeholder="설비명">
    `;
  } else if (sheetName === 'person') {
    fieldsContainer.innerHTML = '<input type="text" class="modal-input" id="addField1" placeholder="폐기자명">';
  } else if (sheetName.includes('TMNO')) {
    fieldsContainer.innerHTML = `
      <input type="text" class="modal-input" id="addField1" placeholder="TM-NO">
      <input type="text" class="modal-input" id="addField2" placeholder="품명">
      <input type="number" step="0.001" class="modal-input" id="addField3" placeholder="단위중량">
      <div class="tmno-process-checks">
        <label class="check-label">성형
          <select class="modal-input modal-select-sm" id="addFieldForming">
            <option value="y" selected>y</option>
            <option value="">빈칸</option>
          </select>
        </label>
        <label class="check-label">소결
          <select class="modal-input modal-select-sm" id="addFieldSintering">
            <option value="y" selected>y</option>
            <option value="">빈칸</option>
          </select>
        </label>
        <label class="check-label">후처리
          <select class="modal-input modal-select-sm" id="addFieldPostProc">
            <option value="y" selected>y</option>
            <option value="">빈칸</option>
          </select>
        </label>
      </div>
    `;
  } else if (sheetName === 'scrap_name') {
    fieldsContainer.innerHTML = '<input type="text" class="modal-input" id="addField1" placeholder="폐기사유">';
  }
}

async function addData() {
  let data = [];

  if (currentAddSheet === 'Depart' || currentAddSheet === 'scrap_name') {
    const val1 = document.getElementById('addField1').value.trim();
    if (!val1) {
      showToast('값을 입력해주세요.', 'error');
      return;
    }
    data = [val1];
  } else if (currentAddSheet === 'Process') {
    const val1 = document.getElementById('addField1').value;
    const val2 = document.getElementById('addField2').value.trim();
    if (!val1 || !val2) {
      showToast('모든 값을 입력해주세요.', 'error');
      return;
    }
    data = [val1, val2];
  } else if (currentAddSheet === 'machine') {
    const val1 = document.getElementById('addField1').value;
    const val2 = document.getElementById('addField2').value;
    const val3 = document.getElementById('addField3').value.trim();
    if (!val1 || !val2 || !val3) {
      showToast('모든 값을 입력해주세요.', 'error');
      return;
    }
    data = [val1, val2, val3];
  } else if (currentAddSheet === 'person') {
    const val1 = document.getElementById('addField1').value.trim();
    if (!val1) {
      showToast('이름을 입력해주세요.', 'error');
      return;
    }
    data = [val1];
  } else if (currentAddSheet.includes('TMNO')) {
    const val1 = document.getElementById('addField1').value.trim();
    const val2 = document.getElementById('addField2').value.trim();
    const val3 = document.getElementById('addField3').value;
    if (!val1 || !val2) {
      showToast('TM-NO와 품명을 입력해주세요.', 'error');
      return;
    }
    const forming = document.getElementById('addFieldForming').value;
    const sintering = document.getElementById('addFieldSintering').value;
    const postProc = document.getElementById('addFieldPostProc').value;
    data = [val1, val2, parseFloat(val3) || 0, forming, sintering, postProc];
  }

  showLoading();
  const result = await apiCall(`/api/master_data/${currentAddSheet}`, {
    method: 'POST',
    body: JSON.stringify({ data: data })
  });
  hideLoading();

  closeModal('addModal');
  if (result && result.success) {
    showToast(result.message, 'success');
    tmnoCache = null;
    invalidateMasterCache(currentAddSheet);
    // 관리자 모드면 테이블 갱신
    if (state.isAdmin && document.getElementById('adminScreen').classList.contains('active')) {
      showMasterDataManager(currentAddSheet);
    }
    // 입력 화면: 추가한 시트에 해당하는 목록 직접 갱신
    refreshBySheet(currentAddSheet);
  } else {
    showToast(result?.message || '추가 실패', 'error');
  }
}

function invalidateMasterCache(sheetName) {
  if (sheetName === 'Depart') masterCache.departments = null;
  else if (sheetName === 'person') masterCache.persons = null;
  else if (sheetName === 'scrap_name') masterCache.scrapReasons = null;
}

function refreshBySheet(sheetName) {
  if (sheetName === 'Depart') loadDepartments();
  else if (sheetName === 'Process') loadProcesses();
  else if (sheetName === 'machine') loadMachinesForEntry();
  else if (sheetName === 'person') loadPersons();
  else if (sheetName === 'scrap_name') {
    // 폐기사유 모달이 열려있으면 목록 즉시 갱신
    if (document.getElementById('reasonModal').classList.contains('active')) {
      showReasonSelector();
    }
  }
}

// ==================== 관리자 모드 ====================

function showAdminLogin() {
  document.getElementById('adminPassword').value = '';
  document.getElementById('adminLoginModal').classList.add('active');
}

async function adminLogin() {
  const password = document.getElementById('adminPassword').value;

  showLoading();
  const isValid = await apiCall('/api/verify_password', {
    method: 'POST',
    body: JSON.stringify({ password: password })
  });
  hideLoading();

  if (isValid) {
    state.isAdmin = true;
    state.adminPassword = password;
    closeModal('adminLoginModal');
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById('adminScreen').classList.add('active');
    history.pushState({ screen: 'admin' }, '');
  } else {
    showToast('비밀번호가 올바르지 않습니다.', 'error');
  }
}

async function showMasterDataManager(sheetName, btn) {
  currentMasterSheet = sheetName;

  document.querySelectorAll('.btn-admin').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');

  showLoading();
  const result = await apiCall(`/api/master_data/${sheetName}`);
  hideLoading();

  if (result && result.success) {
    displayMasterDataTable(result.headers, result.rows, sheetName);
  } else {
    document.getElementById('adminContent').innerHTML = '<p>데이터를 불러올 수 없습니다.</p>';
  }
}

function displayMasterDataTable(headers, rows, sheetName) {
  currentTableHeaders = headers;
  currentTableRows = rows;

  const container = document.getElementById('adminContent');
  container.innerHTML = '';

  // 추가 버튼 (마스터 데이터 시트일 때)
  if (sheetName && sheetName !== 'Data') {
    const addBtnLabel = getSheetLabel(sheetName);
    const addBtn = document.createElement('button');
    addBtn.className = 'btn btn-add-small';
    addBtn.textContent = `+ ${addBtnLabel} 추가`;
    addBtn.style.marginBottom = '15px';
    addBtn.addEventListener('click', () => showAddModal(sheetName, addBtnLabel));
    container.appendChild(addBtn);
  }

  const table = document.createElement('table');
  table.className = 'data-table';

  // 헤더
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });
  const thAction = document.createElement('th');
  thAction.textContent = '관리';
  headerRow.appendChild(thAction);
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // 바디
  const tbody = document.createElement('tbody');
  rows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    row.data.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell != null ? String(cell) : '-';
      tr.appendChild(td);
    });
    const tdAction = document.createElement('td');

    const editBtn = document.createElement('button');
    editBtn.className = 'btn btn-edit';
    editBtn.textContent = '수정';
    editBtn.addEventListener('click', () => editMasterData(idx, row.rowIndex));

    const delBtn = document.createElement('button');
    delBtn.className = 'btn btn-delete';
    delBtn.textContent = '삭제';
    delBtn.addEventListener('click', () => deleteMasterData(row.rowIndex));

    tdAction.appendChild(editBtn);
    tdAction.appendChild(delBtn);
    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  container.appendChild(table);
}

function getSheetLabel(sheetName) {
  const labels = {
    'Depart': '부서', 'Process': '공정', 'machine': '설비',
    'person': '폐기자', '1Part_TMNO': '1Part TM-NO', '2Part_TMNO': '2Part TM-NO',
    'scrap_name': '폐기사유'
  };
  return labels[sheetName] || sheetName;
}

function editMasterData(dataIndex, rowIndex) {
  const row = currentTableRows[dataIndex];
  const headers = currentTableHeaders;

  if (!row) {
    showToast('데이터를 찾을 수 없습니다.', 'error');
    return;
  }

  const fieldsContainer = document.getElementById('editModalFields');
  fieldsContainer.innerHTML = '';

  headers.forEach((header, index) => {
    const value = row.data[index] || '';
    const div = document.createElement('div');
    div.style.marginBottom = '15px';

    const label = document.createElement('label');
    label.style.cssText = 'display:block; margin-bottom:5px; color:#aaa; font-size:14px;';
    label.textContent = header;

    let input;
    if (header === '불량해당') {
      // 불량해당: select (해당 / 빈칸)
      input = document.createElement('select');
      input.className = 'modal-input edit-field';
      input.dataset.index = index;
      const optEmpty = document.createElement('option');
      optEmpty.value = '';
      optEmpty.textContent = '빈칸';
      const optYes = document.createElement('option');
      optYes.value = '해당';
      optYes.textContent = '해당';
      input.appendChild(optEmpty);
      input.appendChild(optYes);
      input.value = String(value);
    } else {
      input = document.createElement('input');
      input.type = 'text';
      input.className = 'modal-input edit-field';
      input.dataset.index = index;
      input.value = String(value);
    }

    div.appendChild(label);
    div.appendChild(input);
    fieldsContainer.appendChild(div);
  });

  document.getElementById('editRowIndex').value = rowIndex;
  document.getElementById('editModal').classList.add('active');
}

async function saveEditData() {
  const rowIndex = parseInt(document.getElementById('editRowIndex').value);
  const fields = document.querySelectorAll('#editModalFields .edit-field');
  const newData = [];

  fields.forEach(field => {
    newData.push(field.value);
  });

  showLoading();
  const result = await apiCall(`/api/master_data/${currentMasterSheet}/${rowIndex}`, {
    method: 'PUT',
    body: JSON.stringify({
      data: newData,
      password: state.adminPassword
    })
  });
  hideLoading();

  if (result && result.success) {
    showToast(result.message, 'success');
    closeModal('editModal');
    invalidateMasterCache(currentMasterSheet);
    if (currentMasterSheet === 'Data') {
      showScrapRecords(null, scrapRecordPage);
    } else {
      showMasterDataManager(currentMasterSheet);
    }
  } else {
    showToast(result?.message || '수정 실패', 'error');
  }
}

async function deleteMasterData(rowIndex) {
  // 커스텀 확인 모달 사용
  showDeleteConfirm(() => doDeleteMasterData(rowIndex));
}

function showDeleteConfirm(onConfirm) {
  const modal = document.getElementById('deleteConfirmModal');
  modal.classList.add('active');

  const confirmBtn = document.getElementById('deleteConfirmBtn');
  const newBtn = confirmBtn.cloneNode(true);
  confirmBtn.parentNode.replaceChild(newBtn, confirmBtn);
  newBtn.id = 'deleteConfirmBtn';
  newBtn.addEventListener('click', () => {
    closeModal('deleteConfirmModal');
    onConfirm();
  });
}

async function doDeleteMasterData(rowIndex) {
  showLoading();
  const result = await apiCall(`/api/master_data/${currentMasterSheet}/${rowIndex}`, {
    method: 'DELETE',
    body: JSON.stringify({ password: state.adminPassword })
  });
  hideLoading();

  if (result && result.success) {
    showToast(result.message, 'success');
    invalidateMasterCache(currentMasterSheet);
    if (currentMasterSheet === 'Data') {
      showScrapRecords(null, scrapRecordPage);
    } else {
      showMasterDataManager(currentMasterSheet);
    }
  } else {
    showToast(result?.message || '삭제 실패', 'error');
  }
}

let scrapRecordPage = 1;

// 기록 조회 필터 상태
let scrapFilter = {
  startDate: '',
  endDate: '',
  part: '',
  reasonType: '',
  defect: ''
};

async function showScrapRecords(btn, page) {
  currentMasterSheet = 'Data';
  scrapRecordPage = page || 1;

  document.querySelectorAll('.btn-admin').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');

  const container = document.getElementById('adminContent');

  // 필터 UI가 없으면 생성
  if (!document.getElementById('scrapFilterPanel')) {
    const today = new Date().toISOString().slice(0, 10);
    const filterHtml = `
      <div id="scrapFilterPanel" class="scrap-filter-panel">
        <div class="scrap-filter-row">
          <div class="scrap-filter-group">
            <label>시작일</label>
            <input type="date" id="scrapStartDate" class="scrap-filter-input" value="${today}">
          </div>
          <span class="scrap-filter-sep">~</span>
          <div class="scrap-filter-group">
            <label>종료일</label>
            <input type="date" id="scrapEndDate" class="scrap-filter-input" value="${today}">
          </div>
          <div class="scrap-filter-group">
            <label>Part</label>
            <select id="scrapPartFilter" class="scrap-filter-input">
              <option value="">전체</option>
              <option value="1Part">1Part</option>
              <option value="2Part">2Part</option>
            </select>
          </div>
          <div class="scrap-filter-group">
            <label>폐기사유</label>
            <select id="scrapReasonFilter" class="scrap-filter-input">
              <option value="">전체</option>
              <option value="main">공정·셋팅불량</option>
              <option value="etc">기타폐기</option>
            </select>
          </div>
          <div class="scrap-filter-group">
            <label>불량해당</label>
            <select id="scrapDefectFilter" class="scrap-filter-input">
              <option value="">전체</option>
              <option value="해당">해당</option>
              <option value="empty">빈칸</option>
            </select>
          </div>
          <button class="btn btn-filter-search" onclick="applyScrapFilter()">검색</button>
        </div>
      </div>
      <div id="scrapTableArea"></div>
    `;
    container.innerHTML = filterHtml;

    // 필터 상태 복원
    if (scrapFilter.startDate) document.getElementById('scrapStartDate').value = scrapFilter.startDate;
    if (scrapFilter.endDate) document.getElementById('scrapEndDate').value = scrapFilter.endDate;
    if (scrapFilter.part) document.getElementById('scrapPartFilter').value = scrapFilter.part;
    if (scrapFilter.reasonType) document.getElementById('scrapReasonFilter').value = scrapFilter.reasonType;
    if (scrapFilter.defect) document.getElementById('scrapDefectFilter').value = scrapFilter.defect;
  }

  // 필터 파라미터 조립
  let params = `page=${scrapRecordPage}&per_page=100`;
  if (scrapFilter.startDate) params += `&start_date=${scrapFilter.startDate}`;
  if (scrapFilter.endDate) params += `&end_date=${scrapFilter.endDate}`;
  if (scrapFilter.part) params += `&part=${encodeURIComponent(scrapFilter.part)}`;
  if (scrapFilter.reasonType) params += `&reason_type=${scrapFilter.reasonType}`;
  if (scrapFilter.defect) params += `&defect=${encodeURIComponent(scrapFilter.defect)}`;

  showLoading();
  const result = await apiCall(`/api/scrap_records?${params}`);
  hideLoading();

  if (result && result.success) {
    displayScrapRecordTable(result.headers, result.rows);
    if (result.pagination) {
      displayScrapPagination(result.pagination);
    }
  } else {
    document.getElementById('scrapTableArea').innerHTML = '<p>데이터를 불러올 수 없습니다.</p>';
  }
}

function applyScrapFilter() {
  scrapFilter.startDate = document.getElementById('scrapStartDate').value;
  scrapFilter.endDate = document.getElementById('scrapEndDate').value;
  scrapFilter.part = document.getElementById('scrapPartFilter').value;
  scrapFilter.reasonType = document.getElementById('scrapReasonFilter').value;
  scrapFilter.defect = document.getElementById('scrapDefectFilter').value;
  showScrapRecords(null, 1);
}

function displayScrapRecordTable(headers, rows) {
  currentTableHeaders = headers;
  currentTableRows = rows;

  const container = document.getElementById('scrapTableArea');
  container.innerHTML = '';

  const tableWrap = document.createElement('div');
  tableWrap.style.overflowX = 'auto';

  const table = document.createElement('table');
  table.className = 'data-table';

  // 헤더
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });
  const thAction = document.createElement('th');
  thAction.textContent = '관리';
  headerRow.appendChild(thAction);
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // 바디
  const tbody = document.createElement('tbody');
  rows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    row.data.forEach((cell, colIdx) => {
      const td = document.createElement('td');
      td.textContent = cell != null ? String(cell) : '-';
      tr.appendChild(td);
    });
    const tdAction = document.createElement('td');

    const editBtn = document.createElement('button');
    editBtn.className = 'btn btn-edit';
    editBtn.textContent = '수정';
    editBtn.addEventListener('click', () => editMasterData(idx, row.rowIndex));

    const delBtn = document.createElement('button');
    delBtn.className = 'btn btn-delete';
    delBtn.textContent = '삭제';
    delBtn.addEventListener('click', () => deleteMasterData(row.rowIndex));

    tdAction.appendChild(editBtn);
    tdAction.appendChild(delBtn);
    tr.appendChild(tdAction);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  tableWrap.appendChild(table);
  container.appendChild(tableWrap);
}

function displayScrapPagination(pg) {
  const container = document.getElementById('scrapTableArea');
  const div = document.createElement('div');
  div.style.cssText = 'display:flex; justify-content:center; align-items:center; gap:12px; margin-top:15px; padding:10px;';

  const prevBtn = document.createElement('button');
  prevBtn.className = 'btn btn-select';
  prevBtn.textContent = '◀ 이전';
  prevBtn.disabled = pg.page <= 1;
  prevBtn.addEventListener('click', () => showScrapRecords(null, pg.page - 1));

  const info = document.createElement('span');
  info.style.cssText = 'color:#aaa; font-size:14px;';
  info.textContent = `${pg.page} / ${pg.total_pages} 페이지 (총 ${pg.total}건)`;

  const nextBtn = document.createElement('button');
  nextBtn.className = 'btn btn-select';
  nextBtn.textContent = '다음 ▶';
  nextBtn.disabled = pg.page >= pg.total_pages;
  nextBtn.addEventListener('click', () => showScrapRecords(null, pg.page + 1));

  div.appendChild(prevBtn);
  div.appendChild(info);
  div.appendChild(nextBtn);
  container.appendChild(div);
}

function displayPagination(pg) {
  const container = document.getElementById('adminContent');
  const div = document.createElement('div');
  div.style.cssText = 'display:flex; justify-content:center; align-items:center; gap:12px; margin-top:15px; padding:10px;';

  const prevBtn = document.createElement('button');
  prevBtn.className = 'btn btn-select';
  prevBtn.textContent = '◀ 이전';
  prevBtn.disabled = pg.page <= 1;
  prevBtn.addEventListener('click', () => showScrapRecords(null, pg.page - 1));

  const info = document.createElement('span');
  info.style.cssText = 'color:#aaa; font-size:14px;';
  info.textContent = `${pg.page} / ${pg.total_pages} 페이지 (총 ${pg.total}건)`;

  const nextBtn = document.createElement('button');
  nextBtn.className = 'btn btn-select';
  nextBtn.textContent = '다음 ▶';
  nextBtn.disabled = pg.page >= pg.total_pages;
  nextBtn.addEventListener('click', () => showScrapRecords(null, pg.page + 1));

  div.appendChild(prevBtn);
  div.appendChild(info);
  div.appendChild(nextBtn);
  container.appendChild(div);
}

// ==================== 유틸리티 ====================

function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function showLoading() {
  document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('active');
}

function showToast(message, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = 'toast active ' + type;

  setTimeout(() => {
    toast.classList.remove('active');
  }, 3000);
}

// ==================== 브라우저 뒤로가기 처리 ====================

window.addEventListener('popstate', function(e) {
  const s = e.state;
  if (!s || s.screen === 'start') {
    resetState();
    document.querySelectorAll('.screen').forEach(sc => sc.classList.remove('active'));
    document.getElementById('startScreen').classList.add('active');
  } else if (s.screen === 'input') {
    document.querySelectorAll('.screen').forEach(sc => sc.classList.remove('active'));
    document.getElementById('inputScreen').classList.add('active');
    if (s.step && s.step > 1) {
      goToStep(s.step - 1);
    }
  } else if (s.screen === 'admin') {
    document.querySelectorAll('.screen').forEach(sc => sc.classList.remove('active'));
    document.getElementById('startScreen').classList.add('active');
  }
});

// ==================== Excel 다운로드 ====================

function showExcelDownload(btn) {
  currentMasterSheet = '';
  document.querySelectorAll('.btn-admin').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');

  const container = document.getElementById('adminContent');
  const today = new Date().toISOString().slice(0, 10);

  container.innerHTML = `
    <div class="excel-download-panel">
      <h3 class="excel-panel-title">Excel 다운로드</h3>
      <p class="excel-panel-desc">기간을 선택하면 해당 기간의 폐기불량 기록을 Excel 파일로 다운로드합니다.</p>

      <div class="excel-quick-btns">
        <button class="btn-quick" onclick="setExcelRange('today')">오늘</button>
        <button class="btn-quick" onclick="setExcelRange('week')">최근 1주</button>
        <button class="btn-quick" onclick="setExcelRange('month')">최근 1개월</button>
        <button class="btn-quick" onclick="setExcelRange('3month')">최근 3개월</button>
        <button class="btn-quick" onclick="setExcelRange('year')">최근 1년</button>
      </div>

      <div class="excel-date-row">
        <div class="excel-date-group">
          <label>시작일</label>
          <input type="date" id="excelStartDate" class="excel-date-input" value="${today}">
        </div>
        <span class="excel-date-separator">~</span>
        <div class="excel-date-group">
          <label>종료일</label>
          <input type="date" id="excelEndDate" class="excel-date-input" value="${today}">
        </div>
      </div>

      <button class="btn btn-excel" onclick="downloadExcel()">Excel 다운로드</button>
    </div>
  `;
}

function setExcelRange(range) {
  const today = new Date();
  let start = new Date();

  if (range === 'today') {
    start = today;
  } else if (range === 'week') {
    start.setDate(today.getDate() - 7);
  } else if (range === 'month') {
    start.setMonth(today.getMonth() - 1);
  } else if (range === '3month') {
    start.setMonth(today.getMonth() - 3);
  } else if (range === 'year') {
    start.setFullYear(today.getFullYear() - 1);
  }

  document.getElementById('excelStartDate').value = start.toISOString().slice(0, 10);
  document.getElementById('excelEndDate').value = today.toISOString().slice(0, 10);

  // 선택된 버튼 하이라이트
  document.querySelectorAll('.excel-quick-btns .btn-quick').forEach(b => b.classList.remove('selected'));
  event.target.classList.add('selected');
}

function downloadExcel() {
  const startDate = document.getElementById('excelStartDate').value;
  const endDate = document.getElementById('excelEndDate').value;

  if (!startDate || !endDate) {
    showToast('시작일과 종료일을 선택해주세요.', 'error');
    return;
  }

  if (startDate > endDate) {
    showToast('시작일이 종료일보다 늦습니다.', 'error');
    return;
  }

  showToast('Excel 파일 생성 중...', 'info');
  window.location.href = `/api/export_excel?start_date=${startDate}&end_date=${endDate}`;
}


// ==================== 초기화 ====================

document.addEventListener('DOMContentLoaded', function() {
  history.replaceState({ screen: 'start' }, '');
  // 마스터 데이터 미리 로드 (부서/인원/불량유형 - 1회 호출)
  loadInitData();
  console.log('폐기불량 관리시스템 초기화 완료');
});

async function loadInitData() {
  const data = await apiCall('/api/init_data');
  if (data) {
    masterCache.departments = data.departments;
    masterCache.persons = data.persons;
    masterCache.scrapReasons = data.scrapReasons;
    cachedPersonNames = data.persons;
  }
}
