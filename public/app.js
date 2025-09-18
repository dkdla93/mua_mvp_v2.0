// =========================
// Global State
// =========================
var appState = {
  excelData: null,
  allSheets: {},
  currentSheet: null,
  minimapImage: null,
  sceneImages: [],
  materials: [],                 // [{ id, tabName, material, area, item, remarks, brand, imageUrl, image, category, ... }]
  sceneMaterialMapping: {},      // sceneIndex -> [materialId]
  currentSelectedScene: 0,
  minimapBoxes: {}               // sceneIndex -> {x,y,w,h} in [0..1]
};

// Canvas (minimap)
var canvas, ctx, isDrawing = false, startX = 0, startY = 0, currentRect = null, minimapImgObj = null;

document.addEventListener('DOMContentLoaded', function () {
  initializeEventListeners();
});

// =========================
// Event bindings
// =========================
function initializeEventListeners() {
  document.getElementById('excelFile').addEventListener('change', handleExcelUpload);
  document.getElementById('minimapFile').addEventListener('change', handleMinimapUpload);
  document.getElementById('sceneFiles').addEventListener('change', handleSceneUpload);
  document.getElementById('generateBtn').addEventListener('click', generatePPT);
}

// =========================
// Excel handling (SheetJS)
// =========================
function handleExcelUpload(e) {
  var file = e.target.files[0];
  if (!file) return;

  var reader = new FileReader();
  reader.onload = function (ev) {
    try {
      var data = new Uint8Array(ev.target.result);
      var workbook = XLSX.read(data, {
        type: 'array',
        cellStyles: true, cellFormulas: true, cellDates: true, cellNF: true, sheetStubs: true
      });

      appState.allSheets = {};
      for (var i = 0; i < workbook.SheetNames.length; i++) {
        var sheetName = workbook.SheetNames[i];
        var sheet = workbook.Sheets[sheetName];
        appState.allSheets[sheetName] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      }

      var infoDiv = document.getElementById('excelInfo');
      infoDiv.innerHTML = '<strong>업로드 완료:</strong> ' + file.name + ' (' + workbook.SheetNames.length + '개 시트)';
      infoDiv.style.display = 'block';

      autoSelectFirstValidSheet();
      checkAllFilesUploaded();
    } catch (err) {
      showStatus('엑셀 파일 읽기 실패: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function autoSelectFirstValidSheet() {
  var validSheets = [];
  var sheetNames = Object.keys(appState.allSheets);

  for (var i = 0; i < sheetNames.length; i++) {
    var s = sheetNames[i];
    if (s.match(/^\d+\./) || s.indexOf('1.') !== -1) validSheets.push(s);
  }
  if (validSheets.length === 0) {
    for (var j = 1; j < sheetNames.length; j++) validSheets.push(sheetNames[j]);
  }

  if (validSheets.length > 0) {
    appState.currentSheet = validSheets[0];
    appState.excelData = appState.allSheets[validSheets[0]];
    parseExcelData();
  }
}

/**
 * 스펙 엑셀의 행 구조 예시(시트별):
 *  - [?, 'AREA', <값>]
 *  - [?, 'MATERIAL', <값>]
 *  - [?, 'ITEM', <값>]
 *  - [?, 'REMARKS' or 'REMARK', <값>]
 *  - [?, 'IMAGE', <URL/파일명/base64>]
 *  - ...
 *  영역 header가 'AREA' 나오면 새 레코드 시작으로 간주.
 */
function parseExcelData() {
  appState.materials = [];
  var currentMaterial = null;
  var currentCategory = '';
  var allSheetNames = Object.keys(appState.allSheets);

  for (var si = 0; si < allSheetNames.length; si++) {
    var sheetName = allSheetNames[si];
    if (sheetName.match(/^A\./)) continue; // 표지/설명 시트 무시(예: A.*)

    var sheetData = appState.allSheets[sheetName];
    currentCategory = '';

    for (var i = 1; i < sheetData.length; i++) {
      var row = sheetData[i];
      if (!row || row.length < 2) continue;

      // 상단 카테고리 영역(예: MATERIAL / SWITCH / LIGHT 등)
      if (
        row[0] && row[0].toString().trim() !== '' &&
        (row[0].toString().indexOf('MATERIAL') !== -1 ||
         row[0].toString().indexOf('SWITCH')   !== -1 ||
         row[0].toString().indexOf('LIGHT')    !== -1)
      ) {
        currentCategory = row[0].toString().trim();
      }

      var key = (row[1] || '').toString().toUpperCase();

      if (key.indexOf('AREA') !== -1) {
        // 새 레코드 시작
        if (currentMaterial) appState.materials.push(currentMaterial);
        currentMaterial = {
          id: appState.materials.length + 1,
          tabName: sheetName,
          displayId: '#' + sheetName,
          category: currentCategory || 'MATERIAL',

          material: '',
          area: (row[2] || ''),
          item: '',
          remarks: '',
          brand: '',

          imageUrl: '',
          image: null
        };
      } else if (key.indexOf('MATERIAL') !== -1 && currentMaterial) {
        currentMaterial.material = row[2] || '';
      } else if (key.indexOf('ITEM') !== -1 && currentMaterial) {
        currentMaterial.item = row[2] || '';
      } else if ((key.indexOf('REMARKS') !== -1 || key.indexOf('REMARK') !== -1) && currentMaterial) {
        currentMaterial.remarks = row[2] || '';
      } else if (key.indexOf('IMAGE') !== -1 && currentMaterial) {
        // 이미지 셀에 URL/파일명/base64 중 택1 기입해주면 사용
        currentMaterial.imageUrl = row[2] || '';
        currentMaterial.image    = row[2] || null;
      } else if (currentMaterial) {
        // 보조 필드(브랜드 등)
        if (typeof row[4] !== 'undefined' && row[4]) currentMaterial.brand = row[4];
      }
    }
    if (currentMaterial) {
      appState.materials.push(currentMaterial);
      currentMaterial = null;
    }
  }

  setTimeout(checkAllFilesUploaded, 100);
}

// =========================
// Minimap image handling
// =========================
function handleMinimapUpload(e) {
  var file = e.target.files[0];
  if (!file) return;

  var reader = new FileReader();
  reader.onload = function (ev) {
    appState.minimapImage = ev.target.result;

    var infoDiv = document.getElementById('minimapInfo');
    infoDiv.innerHTML = '<strong>업로드 완료:</strong> ' + file.name +
      '<br><img src="' + appState.minimapImage + '" style="max-width:200px;margin-top:10px;border-radius:5px;">';
    infoDiv.style.display = 'block';

    minimapImgObj = new Image();
    minimapImgObj.onload = function () {
      setupMinimapCanvas();
    };
    minimapImgObj.src = appState.minimapImage;

    setTimeout(checkAllFilesUploaded, 100);
  };
  reader.readAsDataURL(file);
}

// =========================
// Scene images
// =========================
function handleSceneUpload(e) {
  var files = Array.from(e.target.files);
  if (files.length === 0) return;

  appState.sceneImages = [];
  var loaded = 0;

  for (var i = 0; i < files.length; i++) {
    (function (idx, f) {
      var r = new FileReader();
      r.onload = function (ev) {
        appState.sceneImages.push({ name: f.name, data: ev.target.result, index: idx });
        loaded++;
        if (loaded === files.length) {
          displaySceneInfo();
          checkAllFilesUploaded();
        }
      };
      r.readAsDataURL(f);
    })(i, files[i]);
  }
}

function displaySceneInfo() {
  var html = '<strong>업로드 완료:</strong> ' + appState.sceneImages.length + '개 장면 이미지<br>';
  html += '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:10px;">';
  for (var i = 0; i < appState.sceneImages.length; i++) {
    var s = appState.sceneImages[i];
    html += '<div style="text-align:center;"><img src="' + s.data + '" style="width:80px;height:60px;object-fit:cover;border-radius:3px;"><div style="font-size:0.8em;margin-top:5px;">' + s.name + '</div></div>';
  }
  html += '</div>';

  var el = document.getElementById('sceneInfo');
  el.innerHTML = html;
  el.style.display = 'block';
  setTimeout(checkAllFilesUploaded, 100);
}

// =========================
// UI build
// =========================
function checkAllFilesUploaded() {
  var hasExcel = appState.currentSheet !== null && appState.materials.length > 0;
  var hasMinimap = appState.minimapImage !== null;
  var hasScenes = appState.sceneImages.length > 0;

  if (hasExcel && hasMinimap && hasScenes) {
    try {
      createMaterialInterface();
      document.getElementById('matchingStep').style.display = 'block';
      document.getElementById('generateStep').style.display = 'block';
      document.getElementById('minimapDrawWrap').style.display = 'block';
      showStatus('모든 파일이 업로드되었습니다! 장면별 자재와 미니맵 위치를 지정해주세요.', 'success');
    } catch (e) {
      showStatus('인터페이스 생성 중 오류: ' + e.message, 'error');
    }
  } else {
    var missing = [];
    if (!hasExcel) missing.push('엑셀 파일');
    if (!hasMinimap) missing.push('미니맵 이미지');
    if (!hasScenes) missing.push('장면 이미지들');
    if (missing.length > 0) showStatus('누락된 항목: ' + missing.join(', '), 'error');
  }
}

function createMaterialInterface() {
  createSceneSelector();
  createMaterialTabFilter();
  createMaterialTable();
  selectScene(0);
}

function createSceneSelector() {
  var c = document.getElementById('sceneSelector');
  c.innerHTML = '';
  for (var i = 0; i < appState.sceneImages.length; i++) {
    var s = appState.sceneImages[i];
    var div = document.createElement('div');
    div.className = 'scene-item-selector';
    if (i === 0) div.classList.add('active');
    div.innerHTML =
      '<img src="' + s.data + '" alt="' + s.name + '" class="scene-thumb">' +
      '<div><div style="font-weight:bold;">' + s.name + '</div>' +
      '<div style="font-size:0.8em;color:#7f8c8d;">장면 ' + (i + 1) + '</div>' +
      '<div style="font-size:0.8em;color:#27ae60;" id="scene-count-' + i + '">자재 0개 선택됨</div></div>';
    (function (idx) { div.onclick = function () { selectScene(idx); }; })(i);
    c.appendChild(div);
  }
}

function createMaterialTabFilter() {
  var c = document.getElementById('materialTabButtons');
  c.innerHTML = '';
  var allTabs = {};
  for (var i = 0; i < appState.materials.length; i++) {
    var m = appState.materials[i];
    if (m.tabName) allTabs[m.tabName] = true;
  }
  var tabNames = Object.keys(allTabs);

  var allBtn = document.createElement('button');
  allBtn.className = 'material-tab-filter-btn active';
  allBtn.textContent = '전체 보기';
  allBtn.onclick = function () { filterMaterialsByTab(null); updateActiveFilterButton(allBtn); };
  c.appendChild(allBtn);

  for (var j = 0; j < tabNames.length; j++) {
    (function (tabName) {
      var btn = document.createElement('button');
      btn.className = 'material-tab-filter-btn';
      btn.textContent = tabName;
      btn.onclick = function () { filterMaterialsByTab(tabName); updateActiveFilterButton(btn); };
      c.appendChild(btn);
    })(tabNames[j]);
  }
}

function updateActiveFilterButton(activeBtn) {
  var buttons = document.querySelectorAll('.material-tab-filter-btn');
  for (var i = 0; i < buttons.length; i++) buttons[i].classList.remove('active');
  activeBtn.classList.add('active');
}

function createMaterialTable() {
  var c = document.getElementById('materialTableList');
  c.innerHTML = '';
  if (appState.materials.length === 0) {
    c.innerHTML = '<tr><td colspan="9" style="text-align:center;color:#999;padding:20px;">자재가 없습니다.</td></tr>';
    return;
  }

  for (var i = 0; i < appState.materials.length; i++) {
    var m = appState.materials[i];
    var tr = document.createElement('tr');
    tr.className = 'material-row';
    tr.setAttribute('data-material-id', m.id);
    tr.style.height = '35px';

    var imageHtml = (m.image || m.imageUrl) ?
      '<img src="' + (m.image || m.imageUrl) + '" style="width:40px;height:30px;object-fit:cover;border-radius:3px;" alt="자재">' :
      '<div style="width:40px;height:30px;background:#f0f0f0;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:8px;color:#999;">없음</div>';

    var remarksValue = '';
    if (m.remarks && m.remarks.toString().trim() !== '') remarksValue = m.remarks.toString().trim();
    else if (m.brand && m.brand.toString().trim() !== '') remarksValue = m.brand.toString().trim();

    tr.innerHTML =
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;font-size:0.8em;">' + (i + 1) + '</td>' +
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (m.tabName || '') + '</td>' +
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;font-weight:bold;">' + (m.material || m.category || '') + '</td>' +
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (m.area || '') + '</td>' +
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (m.item || '') + '</td>' +
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (remarksValue || '') + '</td>' +
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;">' + imageHtml + '</td>' +
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;"><input type="checkbox" id="checkbox_' + m.id + '" style="width:16px;height:16px;"></td>';

    (function (mid) {
      var checkbox = tr.querySelector('input[type="checkbox"]');
      checkbox.addEventListener('change', function (e) { toggleMaterial(mid, e.target.checked); });
      tr.addEventListener('click', function (e) {
        if (e.target.type !== 'checkbox') {
          var cb = document.getElementById('checkbox_' + mid);
          if (cb) { cb.checked = !cb.checked; toggleMaterial(mid, cb.checked); }
        }
      });
    })(m.id);

    c.appendChild(tr);
  }
}

function selectScene(sceneIndex) {
  var items = document.querySelectorAll('.scene-item-selector');
  for (var i = 0; i < items.length; i++) items[i].classList.remove('active');
  if (items[sceneIndex]) items[sceneIndex].classList.add('active');

  appState.currentSelectedScene = sceneIndex;

  var sceneName = appState.sceneImages[sceneIndex].name;
  var displayName = sceneName.replace(/\.[^/.]+$/, '');
  document.getElementById('currentSceneTitle').textContent = displayName;

  updateCheckboxStates();
  drawMinimapForScene();
}

function toggleMaterial(materialId, checked) {
  if (!appState.sceneMaterialMapping[appState.currentSelectedScene]) appState.sceneMaterialMapping[appState.currentSelectedScene] = [];
  var mats = appState.sceneMaterialMapping[appState.currentSelectedScene];
  var idx = mats.indexOf(materialId);
  if (checked && idx === -1) mats.push(materialId);
  else if (!checked && idx > -1) mats.splice(idx, 1);
  updateCheckboxStates();
}

function filterMaterialsByTab(selectedTab) {
  var rows = document.querySelectorAll('.material-row');
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var mid = parseInt(row.getAttribute('data-material-id'), 10);
    var m = null;
    for (var j = 0; j < appState.materials.length; j++) if (appState.materials[j].id === mid) { m = appState.materials[j]; break; }
    if (m) row.style.display = (selectedTab === null || m.tabName === selectedTab) ? 'table-row' : 'none';
  }
}

function updateCheckboxStates() {
  var selectedIds = appState.sceneMaterialMapping[appState.currentSelectedScene] || [];
  for (var i = 0; i < appState.materials.length; i++) {
    var m = appState.materials[i];
    var cb = document.getElementById('checkbox_' + m.id);
    if (cb) {
      var row = cb.closest('tr');
      if (selectedIds.indexOf(m.id) > -1) { cb.checked = true; if (row) row.style.backgroundColor = '#e8f5e8'; }
      else { cb.checked = false; if (row) row.style.backgroundColor = ''; }
    }
  }
  var countEl = document.getElementById('scene-count-' + appState.currentSelectedScene);
  if (countEl) countEl.textContent = '자재 ' + selectedIds.length + '개 선택됨';
}

// =========================
/** Minimap drawing */
// =========================
function setupMinimapCanvas() {
  canvas = document.getElementById('minimapCanvas');
  ctx = canvas.getContext('2d');

  canvas.onmousedown = function (e) {
    isDrawing = true;
    var rect = canvas.getBoundingClientRect();
    startX = e.clientX - rect.left;
    startY = e.clientY - rect.top;
    currentRect = { x: startX, y: startY, w: 0, h: 0 };
  };
  canvas.onmousemove = function (e) {
    if (!isDrawing) return;
    var rect = canvas.getBoundingClientRect();
    var x = e.clientX - rect.left;
    var y = e.clientY - rect.top;
    currentRect.w = x - startX;
    currentRect.h = y - startY;
    drawMinimapForScene(currentRect);
  };
  canvas.onmouseup = function () {
    if (!isDrawing) return;
    isDrawing = false;
    if (currentRect) {
      var nx = clamp(Math.min(currentRect.x, currentRect.x + currentRect.w) / canvas.width, 0, 1);
      var ny = clamp(Math.min(currentRect.y, currentRect.y + currentRect.h) / canvas.height, 0, 1);
      var nw = clamp(Math.abs(currentRect.w) / canvas.width, 0, 1);
      var nh = clamp(Math.abs(currentRect.h) / canvas.height, 0, 1);
      appState.minimapBoxes[appState.currentSelectedScene] = { x: nx, y: ny, w: nw, h: nh };
      drawMinimapForScene();
    }
  };

  document.getElementById('resetBoxBtn').onclick = function () {
    delete appState.minimapBoxes[appState.currentSelectedScene];
    drawMinimapForScene();
  };

  drawMinimapForScene();
}

function drawMinimapForScene(liveRect) {
  if (!canvas || !ctx) return;

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  if (minimapImgObj) {
    var iw = minimapImgObj.naturalWidth;
    var ih = minimapImgObj.naturalHeight;
    var scale = Math.min(canvas.width / iw, canvas.height / ih);
    var dw = iw * scale;
    var dh = ih * scale;
    var dx = (canvas.width - dw) / 2;
    var dy = (canvas.height - dh) / 2;
    ctx.drawImage(minimapImgObj, dx, dy, dw, dh);
  }

  var saved = appState.minimapBoxes[appState.currentSelectedScene];
  if (saved) {
    ctx.save();
    ctx.lineWidth = 3;
    ctx.strokeStyle = 'red';
    ctx.fillStyle = 'rgba(255,0,0,0.12)';
    var rx = saved.x * canvas.width;
    var ry = saved.y * canvas.height;
    var rw = saved.w * canvas.width;
    var rh = saved.h * canvas.height;
    ctx.fillRect(rx, ry, rw, rh);
    ctx.strokeRect(rx, ry, rw, rh);
    ctx.restore();
  }

  if (liveRect) {
    ctx.save();
    ctx.lineWidth = 2;
    ctx.strokeStyle = 'red';
    ctx.setLineDash([6, 4]);
    var lx = Math.min(liveRect.x, liveRect.x + liveRect.w);
    var ly = Math.min(liveRect.y, liveRect.y + liveRect.h);
    var lw = Math.abs(liveRect.w);
    var lh = Math.abs(liveRect.h);
    ctx.strokeRect(lx, ly, lw, lh);
    ctx.restore();
  }
}

function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }

// =========================
// PPT Generation (PptxGenJS)
// =========================
function generatePPT() {
  if (Object.keys(appState.sceneMaterialMapping).length === 0) {
    showStatus('장면별 자재를 선택해주세요.', 'error');
    return;
  }
  if (!appState.minimapImage) {
    showStatus('미니맵 이미지를 업로드해주세요.', 'error');
    return;
  }

  document.getElementById('progress').style.display = 'block';
  document.getElementById('generateBtn').disabled = true;

  try {
    var pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 5.625 });
    pptx.layout = 'LAYOUT_16x9';

    var total = appState.sceneImages.length;
    var progress = 0;

    var previewContainer = document.getElementById('slidePreview');
    previewContainer.innerHTML = '';

    for (var i = 0; i < appState.sceneImages.length; i++) {
      var scene = appState.sceneImages[i];
      var selectedIds = appState.sceneMaterialMapping[i] || [];

      var selectedMaterials = [];
      for (var j = 0; j < selectedIds.length; j++) {
        var mid = selectedIds[j];
        for (var k = 0; k < appState.materials.length; k++) {
          if (appState.materials[k].id === mid) { selectedMaterials.push(appState.materials[k]); break; }
        }
      }

      var slide = pptx.addSlide();

      // Title (Korean-safe)
      var title = scene.name.replace(/\.[^/.]+$/, '');
      slide.addText(decodeURIComponent(title), {
        x: 0.5, y: 0.3, fontSize: 20, bold: true, color: '363636', fontFace: 'Apple SD Gothic Neo'
      });

      // Adaptive scene image height based on rows
      var topY = 0.8;
      var sceneH = (selectedMaterials.length > 6 ? 3.0 : 3.5);
      var sceneW = 6.2;

      slide.addImage({ data: scene.data, x: 0.5, y: topY, w: sceneW, h: sceneH, rounding: 6 });

      // Minimap (bigger)
      var miniImg = renderMinimapForExport(i, 3.6, 2.6);
      slide.addText('미니맵', { x: 7.0, y: topY, fontSize: 12, bold: true, color: '555555', fontFace: 'Apple SD Gothic Neo' });
      slide.addImage({ data: miniImg, x: 6.9, y: topY + 0.3, w: 3.3, h: 2.4, rounding: 4 });

      // Table placement (avoid overflow)
      var space = 0.3;
      var tableY = topY + sceneH + space;
      var available = 5.2 - tableY;                      // safe bottom margin ~0.4
      var rowCount = (selectedMaterials.length + 1);     // header + rows
      var rowH = Math.max(0.18, Math.min(0.30, available / rowCount));
      var tableX = 0.5;

      // Rows
      var rows = [['No.', '탭명', 'MATERIAL', 'AREA', 'ITEM', 'REMARKS', 'IMAGE']];
      for (var r = 0; r < selectedMaterials.length; r++) {
        var m = selectedMaterials[r];
        var remarks =
          (m.remarks && m.remarks.toString().trim() !== '' && m.remarks.toString().trim().toUpperCase() !== 'REMARKS')
            ? m.remarks.toString().trim()
            : (m.brand && m.brand.toString().trim() !== '' ? m.brand.toString().trim() : '');
        rows.push([
          String(r + 1),
          m.tabName || '',
          (m.material || m.category || ''),
          m.area || '',
          m.item || '',
          (m.remarks || remarks || ''),
          (m.imageUrl ? '있음' : '')
        ]);
      }

      var tableOpts = {
        x: tableX, y: tableY, w: 9.4,
        colW: [0.6, 1.3, 1.3, 1.2, 3.0, 1.4, 0.8],
        fontSize: 10,
        border: { type: 'solid', color: 'CCCCCC', pt: 1 },
        fill: 'FFFFFF', valign: 'middle', margin: 2, color: '333333',
        rowH: rowH
      };
      slide.addTable(rows, tableOpts);

      // ---------- HTML Preview ----------
      var slideDiv = document.createElement('div');
      slideDiv.style.cssText = 'margin-bottom:30px;padding:20px;border:1px solid #ddd;border-radius:8px;background:#fafafa;';

      var tableHtml = '';
      if (selectedMaterials.length > 0) {
        var trows = '';
        for (var rr = 0; rr < selectedMaterials.length; rr++) {
          var mm = selectedMaterials[rr];
          var rv =
            (mm.remarks && mm.remarks.toString().trim() !== '' && mm.remarks.toString().trim().toUpperCase() !== 'REMARKS')
              ? mm.remarks.toString().trim()
              : (mm.brand && mm.brand.toString().trim() !== '' ? mm.brand.toString().trim() : '');
          trows +=
            '<tr style="height:35px;">' +
            '<td style="border:1px solid #ddd;padding:4px;text-align:center;font-size:0.8em;">' + (rr + 1) + '</td>' +
            '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (mm.tabName || '') + '</td>' +
            '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;font-weight:bold;">' + (mm.material || mm.category || '') + '</td>' +
            '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (mm.area || '') + '</td>' +
            '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (mm.item || '') + '</td>' +
            '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (rv || mm.remarks || '') + '</td>' +
            '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">' + (mm.imageUrl ? '있음' : '') + '</td>' +
            '</tr>';
        }
        tableHtml =
          '<table style="width:100%;margin-top:15px;border-collapse:collapse;font-size:0.85em;">' +
          '<thead><tr style="background:#f8f9fa;height:30px;">' +
          '<th style="border:1px solid #ddd;padding:5px;width:40px;font-size:0.8em;">No.</th>' +
          '<th style="border:1px solid #ddd;padding:5px;width:120px;font-size:0.8em;">탭명</th>' +
          '<th style="border:1px solid #ddd;padding:5px;width:120px;font-size:0.8em;">MATERIAL</th>' +
          '<th style="border:1px solid #ddd;padding:5px;width:60px;font-size:0.8em;">AREA</th>' +
          '<th style="border:1px solid #ddd;padding:5px;font-size:0.8em;">ITEM</th>' +
          '<th style="border:1px solid #ddd;padding:5px;font-size:0.8em;">REMARKS</th>' +
          '<th style="border:1px solid #ddd;padding:5px;font-size:0.8em;">IMAGE</th>' +
          '</tr></thead><tbody>' + trows + '</tbody></table>';
      } else {
        tableHtml = '<p style="text-align:center;color:#7f8c8d;margin-top:15px;">선택된 자재가 없습니다.</p>';
      }

      slideDiv.innerHTML =
        '<h4 style="margin-bottom:15px;color:#2c3e50;">' + title + '</h4>' +
        '<div style="display:flex;gap:20px;align-items:flex-start;">' +
        '<div style="flex:2;"><img src="' + scene.data + '" style="max-width:100%;height:200px;object-fit:cover;border-radius:5px;"></div>' +
        '<div style="flex:1;"><h5>미니맵</h5><img src="' + miniImg + '" style="max-width:100%;height:120px;object-fit:cover;border-radius:5px;"></div>' +
        '</div>' + tableHtml;

      previewContainer.appendChild(slideDiv);

      progress++; updateProgress((progress / total) * 100);
    }

    pptx.writeFile({ fileName: '착공도서.pptx' }).then(function () {
      updateProgress(100);
      showStatus('PPT가 성공적으로 생성되었습니다! 파일이 다운로드됩니다.', 'success');
      document.getElementById('progress').style.display = 'none';
      document.getElementById('generateBtn').disabled = false;
      document.getElementById('previewArea').style.display = 'block';
    });
  } catch (err) {
    showStatus('PPT 생성 중 오류: ' + err.message, 'error');
    document.getElementById('progress').style.display = 'none';
    document.getElementById('generateBtn').disabled = false;
  }
}

// 미니맵 이미지를 지정 크기로 렌더하여 dataURL 반환
function renderMinimapForExport(sceneIndex, outW, outH) {
  var off = document.createElement('canvas');
  var W = (outW || 3) * 96; // approx 96dpi
  var H = (outH || 2) * 96;
  off.width = W; off.height = H;
  var c2 = off.getContext('2d');

  try {
    var img = minimapImgObj;
    if (img) {
      var iw = img.naturalWidth, ih = img.naturalHeight;
      var scale = Math.min(W / iw, H / ih);
      var dw = iw * scale, dh = ih * scale;
      var dx = (W - dw) / 2, dy = (H - dh) / 2;

      c2.fillStyle = '#fff'; c2.fillRect(0, 0, W, H);
      c2.drawImage(img, dx, dy, dw, dh);

      var box = appState.minimapBoxes[sceneIndex];
      if (box) {
        c2.save();
        c2.strokeStyle = 'red'; c2.lineWidth = 3;
        c2.fillStyle = 'rgba(255,0,0,0.12)';
        var rx = box.x * W, ry = box.y * H, rw = box.w * W, rh = box.h * H;
        c2.fillRect(rx, ry, rw, rh);
        c2.strokeRect(rx, ry, rw, rh);
        c2.restore();
      }
    }
  } catch (e) {}
  return off.toDataURL('image/png');
}

// =========================
// UI helpers
// =========================
function updateProgress(pct) {
  document.getElementById('progressFill').style.width = pct + '%';
}
function showStatus(msg, type) {
  var el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status ' + type;
  el.style.display = 'block';
  if (type === 'success') setTimeout(function () { el.style.display = 'none'; }, 5000);
}
