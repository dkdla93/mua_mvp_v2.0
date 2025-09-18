/* =======================
   Global State & Helpers
   ======================= */
var appState = {
  excelData: null,
  allSheets: {},
  currentSheet: null,
  minimapImage: null,
  sceneImages: [],
  materials: [],                 // {id, tabName, category, material, area, item, remarks, brand, imageUrl, image}
  sceneMaterialMapping: {},      // sceneIdx -> [materialId]
  currentSelectedScene: 0,
  minimapBoxes: {}               // sceneIdx -> {x,y,w,h} in [0..1]
};

var canvas, ctx, isDrawing=false, startX=0, startY=0, currentRect=null, minimapImgObj=null;

// === 강력한 헤더/열 탐지 ===
// 시트 상단(최대 40행)을 훑으면서 같은 "행" 안에 AREA/ITEM/REMARKS/IMAGE 키들이
// 함께 나타나는 행을 헤더 행으로 간주. 그 행의 열 인덱스를 채택합니다.
// 없으면 fallback으로 "가장 오른쪽"에 있는 REMARKS/IMAGE 열을 택합니다.
function detectSheetLayout(data){
  function idxInRow(row, kw){
    for (let c=0;c<row.length;c++){
      const v = trim(row[c]).toUpperCase();
      if (v && v.indexOf(kw)!==-1) return c;
    }
    return -1;
  }

  let best=null;
  const scan = Math.min(40, data.length);
  for (let r=0; r<scan; r++){
    const row = data[r] || [];
    const area    = idxInRow(row, 'AREA');
    const item    = idxInRow(row, 'ITEM');
    let   remarks = idxInRow(row, 'REMARKS'); if (remarks<0) remarks = idxInRow(row, 'REMARK');
    const image   = idxInRow(row, 'IMAGE');

    const score = (area>=0) + (item>=0) + (remarks>=0) + (image>=0);
    if (score >= 2){
      const width = Math.max(area, item, remarks, image); // 오른쪽일수록 헤더일 확률 ↑
      if (!best || score>best.score || (score===best.score && width>best.width)){
        best = { row:r, areaCol:area, itemCol:item, remarksCol:remarks, imageCol:image, score, width };
      }
    }
  }

  if (!best){
    // fallback: 첫 40행에서 가장 "오른쪽"에 있는 REMARKS/IMAGE를 채택
    let remarksCol=-1, imageCol=-1;
    for (let r=0; r<scan; r++){
      const row = data[r] || [];
      for (let c=0; c<row.length; c++){
        const v = trim(row[c]).toUpperCase();
        if (v.indexOf('REMARKS')!==-1 || v==='REMARK') remarksCol = Math.max(remarksCol, c);
        if (v.indexOf('IMAGE')  !==-1)                    imageCol   = Math.max(imageCol,   c);
      }
    }
    best = { row:0, areaCol:-1, itemCol:-1, remarksCol, imageCol, score:0, width:Math.max(remarksCol, imageCol) };
  }

  return best;
}

// REMARKS 쓰레기값(라벨/안내문) 걸러내기
function cleanRemarks(val){
  const t = trim(val);
  if (!t) return '';
  if (/^\s*(REMARKS|비고)\s*$/i.test(t)) return '';           // 라벨 그 자체
  if (/이미지\s*첨부|도어\s*모델링/i.test(t)) return '';      // 안내문
  return t;
}


// 시트(2차원 배열) 상단 몇 줄에서 "IMAGE"가 들어간 셀의 열 인덱스를 찾음
function getImageColIndex(sheetData) {
  if (!sheetData || !sheetData.length) return -1;
  var maxScanRows = Math.min(20, sheetData.length);
  for (var r = 0; r < maxScanRows; r++) {
    var row = sheetData[r] || [];
    for (var c = 0; c < row.length; c++) {
      var v = trim(row[c]).toUpperCase();
      if (v.indexOf('IMAGE') !== -1) return c;
    }
  }
  return -1;
}

function isMac(){ return /Macintosh|Mac OS X/.test(navigator.userAgent); }
function koFont(){ return isMac() ? "Apple SD Gothic Neo" : "Malgun Gothic"; }
function koNormalize(s){ try{ return String(s||"").normalize("NFC"); }catch(_){ return String(s||""); } }
function trim(v){ return (v==null)?"":String(v).trim(); }

function stripQuotes(s) {
  // 앞뒤 ' " ‘ ’ “ ” < > ( ) 제거
  return String(s||'').trim()
    .replace(/^[\s'"‘’“”<\(]+/, '')
    .replace(/[\s'"‘’“”>\)]+$/, '');
}
function extractImageUrl(s) {
  if (!s) return '';
  let raw = stripQuotes(s);

  // 1) data URL 우선
  if (/^data:image\//i.test(raw)) return raw;

  // 2) 문자열 내부에서 가장 그럴듯한 http(s) 이미지 URL 뽑기
  //    - 확장자 있는 것
  let m = raw.match(/https?:\/\/[^\s"')]+?\.(?:png|jpe?g|gif|webp|svg)(\?[^\s"')]*)?/i);
  if (m) return m[0];

  // 3) 확장자 없어도 바로 이미지 응답하는 호스트(예: picsum, placeholder 등)
  m = raw.match(/https?:\/\/[^\s"')]+/i);
  if (m) return m[0];

  return '';
}

function isLikelyImage(v) {
  const u = extractImageUrl(v);
  if (!u) return false;
  return /^data:image\//i.test(u) || /^https?:\/\//i.test(u);
}

function thumbHtml(url){
  return '<img src="'+url+'" style="width:40px;height:30px;object-fit:cover;border-radius:3px;" alt="이미지">';
}

document.addEventListener("DOMContentLoaded", function(){
  document.getElementById('excelFile').addEventListener('change', handleExcelUpload);
  document.getElementById('minimapFile').addEventListener('change', handleMinimapUpload);
  document.getElementById('sceneFiles').addEventListener('change', handleSceneUpload);
  document.getElementById('generateBtn').addEventListener('click', generatePPT);
});

/* =======================
   Excel Loading & Parsing
   ======================= */
function handleExcelUpload(e){
  var file = e.target.files[0]; if(!file) return;
  var reader = new FileReader();
  reader.onload = function(ev){
    try{
      var wb = XLSX.read(new Uint8Array(ev.target.result), {
        type:'array', cellStyles:true, cellFormulas:true, cellDates:true, cellNF:true, sheetStubs:true
      });
      appState.allSheets = {};
      wb.SheetNames.forEach(function(sn){
        appState.allSheets[sn] = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1});
      });
      var info = document.getElementById('excelInfo');
      info.innerHTML = '<strong>업로드 완료:</strong> '+file.name+' ('+wb.SheetNames.length+'개 시트)';
      info.style.display = 'block';
      autoSelectFirstValidSheet();
      checkAllFilesUploaded();
    }catch(err){
      showStatus('엑셀 파일 읽기 실패: '+err.message,'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function autoSelectFirstValidSheet(){
  var names = Object.keys(appState.allSheets);
  var valid = names.filter(function(s){ return /^\d+\./.test(s) || s.indexOf('1.')!==-1; });
  if(valid.length===0) valid = names.slice(1);
  if(valid.length){
    appState.currentSheet = valid[0];
    appState.excelData = appState.allSheets[valid[0]];
    parseExcelData();
  }
}

// 키워드를 행 전체에서 찾고, 그 **오른쪽 최대 6칸** 내 첫 non-empty 값을 반환
function findKeyInRow(row, keywords){
  for(var c=0;c<row.length;c++){
    var v = trim(row[c]).toUpperCase();
    for(var k=0;k<keywords.length;k++){
      if(v.indexOf(keywords[k])!==-1) return { idx:c, key:keywords[k] };
    }
  }
  return null;
}
function valueRightOf(row, keyIdx){
  for(var c=keyIdx+1; c<Math.min(row.length, keyIdx+7); c++){
    var vv = row[c];
    if(vv!=null && trim(vv)!=="") return vv;
  }
  // fallback
  if(row[2] && trim(row[2])!=="") return row[2];
  if(row[3] && trim(row[3])!=="") return row[3];
  if(row[4] && trim(row[4])!=="") return row[4];
  return '';
}

function parseExcelData(){
  // 내부 유틸 (이 함수 안에서만 사용)
  function isHeaderNoise(v){
    return /^(REMARKS|REMARK|MANUFAC|ORIGIN|IMAGE|EA|QTY|UNIT|DESCRIPTION)$/i.test(trim(v));
  }
  // 현재 행에서 라벨을 찾아 오른쪽 첫 값셀
  function valueRightByLabel(row, labels){
    var hit = findKeyInRow(row, labels);
    return hit ? trim(valueRightOf(row, hit.idx) || '') : '';
  }
  // 이미지 추출(문자열에서 data:나 http(s) URL만 살림)
  function pickImageFromCell(v){
    var url = extractImageUrl(v);
    return isLikelyImage(url) ? url : '';
  }
  // 같은 행에서 못 찾으면, 고정 라벨열 +1 → 다음 1~2행을 더 탐색
  function robustPickImage(data, r, imageCol){
    var row = data[r] || [];

    // 1) 같은 행에서 "IMAGE" 라벨 오른쪽 값
    var val = valueRightByLabel(row, ['IMAGE']);
    var u = pickImageFromCell(val);
    if (u) return u;

    // 2) 보조: 시트 상단에서 잡은 IMAGE 열이 라벨열인 경우 그 오른쪽 값
    if (imageCol > -1 && row[imageCol+1] != null){
      u = pickImageFromCell(row[imageCol+1]);
      if (u) return u;
    }
    // 3) 그래도 없으면 아래로 1~2행 더 보며 "IMAGE" 라벨 오른쪽 값 시도
    for (var k=1; k<=2; k++){
      var r2 = r + k;
      if (r2 >= data.length) break;
      var row2 = data[r2] || [];
      var val2 = valueRightByLabel(row2, ['IMAGE']);
      u = pickImageFromCell(val2);
      if (u) return u;
      // 보조 2: imageCol+1도 한 번 더 시도
      if (imageCol > -1 && row2[imageCol+1] != null){
        u = pickImageFromCell(row2[imageCol+1]);
        if (u) return u;
      }
    }
    return '';
  }

  appState.materials = [];

  var sheets = Object.keys(appState.allSheets);
  for (var s = 0; s < sheets.length; s++) {
    var sheetName = sheets[s];
    // 표지/서문 스킵
    if (/^A\./.test(sheetName)) continue;

    var data = appState.allSheets[sheetName];
    if (!data || !data.length) continue;

    // 시트 상단에서 라벨 열(헤더) 대략 잡기 (보조 용도)
    var remarksCol = getColIndexByHeader(data, ['REMARKS','REMARK']);
    var imageCol   = getColIndexByHeader(data, ['IMAGE']);

    var current = null;
    var currentCategory   = ''; // MATERIAL/SWITCH/LIGHT 등 큰 카테고리
    var currentGroupLabel = ''; // A열의 구역 라벨 (예: WALL COVERING [도배])

    for (var r = 1; r < data.length; r++) {
      var row = data[r]; if (!row || row.length < 2) continue;

      // 큰 카테고리 갱신
      var left0Upper = trim(row[0]).toUpperCase();
      if (left0Upper &&
          (left0Upper.indexOf('MATERIAL') !== -1 ||
           left0Upper.indexOf('SWITCH')   !== -1 ||
           left0Upper.indexOf('LIGHT')    !== -1)) {
        currentCategory = trim(row[0]);
      }

      // A열 라벨(헤더성 단어는 제외)
      var a0 = trim(row[0]);
      if (a0 && !/^(MATERIAL|DESCRIPTION)$/i.test(a0)) {
        currentGroupLabel = a0;
      }

      // 키 감지 (DESCRIPTION은 제외)
      var hit = findKeyInRow(row, ['AREA','MATERIAL','ITEM','REMARKS','REMARK','IMAGE']);
      if (!hit) continue;

      if (hit.key === 'AREA') {
        // 새 항목 시작: 이전 항목 푸시
        if (current) appState.materials.push(current);

        current = {
          id: appState.materials.length + 1,
          tabName: sheetName,
          displayId: '#' + sheetName,

          // MATERIAL 기본값: 그룹 라벨 → 카테고리 → 시트명
          material: currentGroupLabel || currentCategory || sheetName || '',
          category: currentCategory || 'MATERIAL',

          area:    trim(valueRightOf(row, hit.idx) || ''),
          item:    '',
          remarks: '',
          brand:   '',
          imageUrl: '',
          image:    null
        };

        // AREA 행에서는 remarks/image 확정하지 않음

      } else if (hit.key === 'MATERIAL' && current) {
        // MATERIAL 값 (오염 라벨은 무시)
        var mv = trim(valueRightOf(row, hit.idx) || '');
        if (mv && !isHeaderNoise(mv)) {
          current.material = mv;
        }

        // 보조로 같은 행의 REMARKS/IMAGE도 시도(ITEM에서 다시 덮어씀)
        var remTry = valueRightByLabel(row, ['REMARKS','REMARK']);
        if (remTry && !isHeaderNoise(remTry)) current.remarks = remTry;

        var imgTry = robustPickImage(data, r, imageCol);
        if (imgTry){ current.imageUrl = imgTry; current.image = imgTry; }

      } else if (hit.key === 'ITEM' && current) {
        // ITEM 값
        current.item = trim(valueRightOf(row, hit.idx) || '');

        // ✅ REMARKS: 같은 행에서 라벨 찾아 오른쪽 값
        var remVal = valueRightByLabel(row, ['REMARKS','REMARK']);
        if (remVal && !isHeaderNoise(remVal)) {
          current.remarks = remVal;
        } else if (remarksCol > -1 && row[remarksCol+1] != null) {
          // 보조: 상단에서 잡은 remarksCol은 라벨열일 수 있으므로 +1을 값으로 시도
          var remVal2 = trim(row[remarksCol+1]);
          if (remVal2 && !isHeaderNoise(remVal2)) current.remarks = remVal2;
        }

        // ✅ IMAGE: 같은/인접 행에서 안정적으로 추출
        var imgUrl = robustPickImage(data, r, imageCol);
        if (imgUrl){ current.imageUrl = imgUrl; current.image = imgUrl; }

      } else if ((hit.key === 'REMARKS' || hit.key === 'REMARK') && current) {
        // 드물게 REMARKS가 행 키로 따로 오는 시트용 안전망
        var rv = trim(valueRightOf(row, hit.idx) || '');
        if (rv && !isHeaderNoise(rv)) current.remarks = rv;

      } else if (hit.key === 'IMAGE' && current) {
        // 드물게 IMAGE가 행 키로 따로 오는 시트용 안전망
        var iv = trim(valueRightOf(row, hit.idx) || '');
        var iu = pickImageFromCell(iv);
        if (iu){ current.imageUrl = iu; current.image = iu; }
      }
    }

    // 마지막 항목 정리
    if (current) {
      if (!trim(current.material)) {
        current.material = currentGroupLabel || currentCategory || sheetName || '';
      }
      appState.materials.push(current);
      current = null;
    }
  }

  // 파싱 후 다음 단계로
  setTimeout(checkAllFilesUploaded, 100);
}




/* =======================
   Minimap
   ======================= */
function handleMinimapUpload(e){
  var file = e.target.files[0]; if(!file) return;
  var reader = new FileReader();
  reader.onload = function(ev){
    appState.minimapImage = ev.target.result;
    var info = document.getElementById('minimapInfo');
    info.innerHTML = '<strong>업로드 완료:</strong> '+file.name+'<br><img src="'+appState.minimapImage+'" style="max-width:200px;margin-top:10px;border-radius:5px;">';
    info.style.display = 'block';
    minimapImgObj = new Image();
    minimapImgObj.onload = function(){ setupMinimapCanvas(); };
    minimapImgObj.src = appState.minimapImage;
    setTimeout(checkAllFilesUploaded,100);
  };
  reader.readAsDataURL(file);
}

/* =======================
   Scenes
   ======================= */
function handleSceneUpload(e){
  var files = Array.from(e.target.files); if(files.length===0) return;
  appState.sceneImages = [];
  var loaded=0;
  files.forEach(function(f,idx){
    var r=new FileReader();
    r.onload=function(ev){
      appState.sceneImages.push({name:f.name, data:ev.target.result, index:idx});
      loaded++; if(loaded===files.length){ displaySceneInfo(); checkAllFilesUploaded(); }
    };
    r.readAsDataURL(f);
  });
}
function displaySceneInfo(){
  var html = '<strong>업로드 완료:</strong> '+appState.sceneImages.length+'개 장면 이미지<br>';
  html += '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:10px;">';
  appState.sceneImages.forEach(function(s){
    html += '<div style="text-align:center;"><img src="'+s.data+'" style="width:80px;height:60px;object-fit:cover;border-radius:3px;"><div style="font-size:0.8em;margin-top:5px;">'+s.name+'</div></div>';
  });
  html+='</div>';
  var el=document.getElementById('sceneInfo'); el.innerHTML=html; el.style.display='block';
  setTimeout(checkAllFilesUploaded,100);
}

/* =======================
   UI build (Step 2)
   ======================= */
function checkAllFilesUploaded(){
  var hasExcel=appState.currentSheet!==null && appState.materials.length>0;
  var hasMinimap=!!appState.minimapImage;
  var hasScenes=appState.sceneImages.length>0;

  if(hasExcel && hasMinimap && hasScenes){
    try{
      createMaterialInterface();
      document.getElementById('matchingStep').style.display='block';
      document.getElementById('generateStep').style.display='block';
      document.getElementById('minimapDrawWrap').style.display='block';
      showStatus('모든 파일이 업로드되었습니다! 장면별 자재와 미니맵 위치를 지정해주세요.','success');
    }catch(e){ showStatus('인터페이스 생성 중 오류: '+e.message,'error'); }
  }else{
    var miss=[]; if(!hasExcel) miss.push('엑셀 파일'); if(!hasMinimap) miss.push('미니맵 이미지'); if(!hasScenes) miss.push('장면 이미지들');
    if(miss.length) showStatus('누락된 항목: '+miss.join(', '),'error');
  }
}

function createMaterialInterface(){
  createSceneSelector();
  createMaterialTabFilter();
  createMaterialTable();
  selectScene(0);
}

function createSceneSelector(){
  var c=document.getElementById('sceneSelector'); c.innerHTML='';
  appState.sceneImages.forEach(function(s, i){
    var div=document.createElement('div');
    div.className='scene-item-selector';
    if(i===0) div.classList.add('active');
    div.innerHTML='<img src="'+s.data+'" alt="'+s.name+'" class="scene-thumb">'+
      '<div><div style="font-weight:bold;">'+s.name+'</div>'+
      '<div style="font-size:0.8em;color:#7f8c8d;">장면 '+(i+1)+'</div>'+
      '<div style="font-size:0.8em;color:#27ae60;" id="scene-count-'+i+'">자재 0개 선택됨</div></div>';
    div.onclick=function(){ selectScene(i); };
    c.appendChild(div);
  });
}

function createMaterialTabFilter(){
  var c=document.getElementById('materialTabButtons'); c.innerHTML='';
  var tabs={}; appState.materials.forEach(function(m){ if(m.tabName) tabs[m.tabName]=true; });
  var allBtn=document.createElement('button'); allBtn.className='material-tab-filter-btn active';
  allBtn.textContent='전체 보기'; allBtn.onclick=function(){ filterMaterialsByTab(null); updateActiveFilterButton(allBtn); };
  c.appendChild(allBtn);
  Object.keys(tabs).forEach(function(name){
    var btn=document.createElement('button'); btn.className='material-tab-filter-btn'; btn.textContent=name;
    btn.onclick=function(){ filterMaterialsByTab(name); updateActiveFilterButton(btn); };
    c.appendChild(btn);
  });
}
function updateActiveFilterButton(activeBtn){
  document.querySelectorAll('.material-tab-filter-btn').forEach(function(b){ b.classList.remove('active'); });
  activeBtn.classList.add('active');
}

function createMaterialTable(){
  var c=document.getElementById('materialTableList'); c.innerHTML='';
  if(appState.materials.length===0){
    c.innerHTML='<tr><td colspan="9" style="text-align:center;color:#999;padding:20px;">자재가 없습니다.</td></tr>'; return;
  }
  appState.materials.forEach(function(m, i){
    var tr=document.createElement('tr'); tr.className='material-row'; tr.dataset.materialId=m.id; tr.style.height='35px';

    var imgSrc = (m.image || m.imageUrl || "");
    var imageHtml = isLikelyImage(imgSrc)
      ? thumbHtml(imgSrc)
      : '<div style="width:40px;height:30px;background:#f0f0f0;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:8px;color:#999;">없음</div>';

    var remarksValue = trim(m.remarks) || trim(m.brand);

    tr.innerHTML =
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;font-size:0.8em;">'+(i+1)+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(m.tabName||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;font-weight:bold;">'+(m.material||m.category||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(m.area||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(m.item||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(remarksValue||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;">'+imageHtml+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;"><input type="checkbox" id="checkbox_'+m.id+'" style="width:16px;height:16px;"></td>';

    var cb = tr.querySelector('input[type="checkbox"]');
    cb.addEventListener('change', function(e){ toggleMaterial(m.id, e.target.checked); });
    tr.addEventListener('click', function(e){
      if(e.target.type!=='checkbox'){ cb.checked=!cb.checked; toggleMaterial(m.id, cb.checked); }
    });

    c.appendChild(tr);
  });
}

function selectScene(sceneIndex){
  document.querySelectorAll('.scene-item-selector').forEach(function(n){ n.classList.remove('active'); });
  var list=document.querySelectorAll('.scene-item-selector'); if(list[sceneIndex]) list[sceneIndex].classList.add('active');
  appState.currentSelectedScene=sceneIndex;
  var sceneName=appState.sceneImages[sceneIndex].name;
  document.getElementById('currentSceneTitle').textContent = sceneName.replace(/\.[^/.]+$/,'');
  updateCheckboxStates(); drawMinimapForScene();
}

function toggleMaterial(materialId, checked){
  var map = appState.sceneMaterialMapping[appState.currentSelectedScene] || (appState.sceneMaterialMapping[appState.currentSelectedScene]=[]);
  var idx = map.indexOf(materialId);
  if(checked && idx===-1) map.push(materialId);
  else if(!checked && idx>-1) map.splice(idx,1);
  updateCheckboxStates();
}

function filterMaterialsByTab(tab){
  document.querySelectorAll('.material-row').forEach(function(row){
    var mid = parseInt(row.dataset.materialId,10);
    var m = appState.materials.find(function(x){ return x.id===mid; });
    row.style.display = (!m || tab===null || m.tabName===tab) ? 'table-row' : 'none';
  });
}
function updateCheckboxStates(){
  var ids = appState.sceneMaterialMapping[appState.currentSelectedScene] || [];
  appState.materials.forEach(function(m){
    var cb = document.getElementById('checkbox_'+m.id);
    if(cb){
      var tr = cb.closest('tr');
      if(ids.indexOf(m.id)>-1){ cb.checked=true; if(tr) tr.style.backgroundColor='#e8f5e8'; }
      else{ cb.checked=false; if(tr) tr.style.backgroundColor=''; }
    }
  });
  var el=document.getElementById('scene-count-'+appState.currentSelectedScene);
  if(el) el.textContent='자재 '+ids.length+'개 선택됨';
}

/* =======================
   Minimap Drawing
   ======================= */
function setupMinimapCanvas(){
  canvas=document.getElementById('minimapCanvas'); ctx=canvas.getContext('2d');
  canvas.onmousedown=function(e){ isDrawing=true; var r=canvas.getBoundingClientRect(); startX=e.clientX-r.left; startY=e.clientY-r.top; currentRect={x:startX,y:startY,w:0,h:0}; };
  canvas.onmousemove=function(e){ if(!isDrawing)return; var r=canvas.getBoundingClientRect(); var x=e.clientX-r.left; var y=e.clientY-r.top; currentRect.w=x-startX; currentRect.h=y-startY; drawMinimapForScene(currentRect); };
  canvas.onmouseup=function(){ if(!isDrawing)return; isDrawing=false; if(currentRect){ var nx=Math.min(currentRect.x,currentRect.x+currentRect.w)/canvas.width; var ny=Math.min(currentRect.y,currentRect.y+currentRect.h)/canvas.height; var nw=Math.abs(currentRect.w)/canvas.width; var nh=Math.abs(currentRect.h)/canvas.height; appState.minimapBoxes[appState.currentSelectedScene]={x:Math.max(0,Math.min(1,nx)),y:Math.max(0,Math.min(1,ny)),w:Math.max(0,Math.min(1,nw)),h:Math.max(0,Math.min(1,nh))}; drawMinimapForScene(); } };
  document.getElementById('resetBoxBtn').onclick=function(){ delete appState.minimapBoxes[appState.currentSelectedScene]; drawMinimapForScene(); };
  drawMinimapForScene();
}
function drawMinimapForScene(liveRect){
  if(!canvas||!ctx) return;
  ctx.clearRect(0,0,canvas.width,canvas.height);
  if(minimapImgObj){
    var iw=minimapImgObj.naturalWidth, ih=minimapImgObj.naturalHeight;
    var scale=Math.min(canvas.width/iw, canvas.height/ih);
    var dw=iw*scale, dh=ih*scale, dx=(canvas.width-dw)/2, dy=(canvas.height-dh)/2;
    ctx.drawImage(minimapImgObj,dx,dy,dw,dh);
  }
  var saved=appState.minimapBoxes[appState.currentSelectedScene];
  if(saved){
    ctx.save(); ctx.lineWidth=3; ctx.strokeStyle='red'; ctx.fillStyle='rgba(255,0,0,0.12)';
    var rx=saved.x*canvas.width, ry=saved.y*canvas.height, rw=saved.w*canvas.width, rh=saved.h*canvas.height;
    ctx.fillRect(rx,ry,rw,rh); ctx.strokeRect(rx,ry,rw,rh); ctx.restore();
  }
  if(liveRect){
    ctx.save(); ctx.lineWidth=2; ctx.strokeStyle='red'; ctx.setLineDash([6,4]);
    var lx=Math.min(liveRect.x,liveRect.x+liveRect.w), ly=Math.min(liveRect.y,liveRect.y+liveRect.h), lw=Math.abs(liveRect.w), lh=Math.abs(liveRect.h);
    ctx.strokeRect(lx,ly,lw,lh); ctx.restore();
  }
}
function clamp(v,min,max){ return Math.max(min, Math.min(max,v)); }

/* =======================
   PPT Generation (PptxGenJS)
   ======================= */
function generatePPT(){
  if(Object.keys(appState.sceneMaterialMapping).length===0){ showStatus('장면별 자재를 선택해주세요.','error'); return; }
  if(!appState.minimapImage){ showStatus('미니맵 이미지를 업로드해주세요.','error'); return; }

  document.getElementById('progress').style.display='block';
  document.getElementById('generateBtn').disabled=true;

  try{
    var pptx = new PptxGenJS();
    pptx.defineLayout({name:'LAYOUT_16x9', width:10, height:5.625});
    pptx.layout = 'LAYOUT_16x9';

    var total=appState.sceneImages.length, progress=0;
    var preview=document.getElementById('slidePreview'); preview.innerHTML='';

    for(var i=0;i<appState.sceneImages.length;i++){
      var scene=appState.sceneImages[i];
      var selectedIds=appState.sceneMaterialMapping[i]||[];
      var mats=[]; selectedIds.forEach(function(mid){
        var m=appState.materials.find(function(x){ return x.id===mid; });
        if(m) mats.push(m);
      });

      var slide=pptx.addSlide();
      var title = koNormalize(scene.name.replace(/\.[^/.]+$/,''));
      slide.addText(title, { x:.5, y:.3, fontSize:20, bold:true, color:'363636', fontFace: koFont() });

      // Scene image (adaptive)
      var topY=0.8;
      var sceneH=(mats.length>6? 3.0 : 3.3);
      var sceneW=6.0;
      slide.addImage({ data:scene.data, x:.5, y:topY, w:sceneW, h:sceneH, rounding:6 });

      // Minimap (clamped)
      var miniImg=renderMinimapForExport(i, 2.9, 2.1);
      var mini={ x:6.7, y:topY+0.2, w:2.9, h:2.1 };
      if(mini.x+mini.w>9.6){ mini.w = 9.6 - mini.x; if(mini.w<2.2) mini.w=2.2; }
      slide.addText('미니맵', { x:mini.x, y:topY, fontSize:12, bold:true, color:'555555', fontFace: koFont() });
      slide.addImage({ data:miniImg, x:mini.x, y:mini.y, w:mini.w, h:mini.h, rounding:4 });

      // Table (never overflow)
      var space=0.25, tableY=topY+sceneH+space;
      var availableH = 5.2 - tableY;                       // bottom margin ~0.4"
      var rowCount = mats.length + 1;
      var rowH = Math.max(0.16, Math.min(0.28, availableH/Math.max(rowCount,1)));
      var tableX = 0.5, tableW = 8.6;                      // 좁혀서 안정 배치
      var colW = [0.55, 1.35, 1.35, 1.2, 2.9, 1.5, 0.75];

      // 만약 합계가 tableW 초과면 스케일 다운
      var sum=colW.reduce(function(a,b){return a+b;},0);
      if(sum>tableW){ var f=tableW/sum; colW=colW.map(function(w){ return Math.max(0.5, w*f); }); }

      // rows: IMAGE 열은 오버레이로 넣을 것이므로 텍스트는 비움
      var rows=[['No.','탭명','MATERIAL','AREA','ITEM','REMARKS','IMAGE']];
      mats.forEach(function(m, idx){
        var rv = trim(m.remarks) || trim(m.brand);
        rows.push([
          String(idx+1),
          m.tabName||'',
          (m.material||m.category||''),
          m.area||'',
          m.item||'',
          rv||'',
          '' // 이미지 텍스트 비움
        ]);
      });

      slide.addTable(rows, {
        x:tableX, y:tableY, w:tableW,
        colW: colW,
        fontSize: 9,
        border:{type:'solid', color:'CCCCCC', pt:1},
        fill:'FFFFFF', valign:'middle', margin:2, color:'333333',
        rowH: rowH
      });

      // --- (PPT) IMAGE 열에 썸네일 오버레이 삽입 ---
      (function placeThumbs(){
        // 열 누적폭으로 IMAGE 셀의 x 좌표 계산
        var acc=[0]; for(var ci=0; ci<colW.length; ci++) acc[ci+1]=acc[ci]+colW[ci];
        var imgColLeft = tableX + acc[6]; // IMAGE 열의 왼쪽 x
        var cellPadX=0.05, cellPadY=0.05;
        var maxW=Math.max(0.45, colW[6] - (cellPadX*2));
        var maxH=Math.max(0.30, rowH    - (cellPadY*2));

        for(var r=0; r<mats.length; r++){
          var m=mats[r];
          var url = (m.image || m.imageUrl || '');
          if(!isLikelyImage(url)) continue;
          var x = imgColLeft + cellPadX;
          var y = tableY + rowH * (1 + r) + cellPadY; // +1: 헤더 다음
          slide.addImage({ data:url, x:x, y:y, w:maxW, h:maxH, rounding:2 });
        }
      })();

      // ---- HTML Preview (Step 3) with thumbnails ----
      var div=document.createElement('div');
      div.style.cssText='margin-bottom:30px;padding:20px;border:1px solid #ddd;border-radius:8px;background:#fafafa;';
      var trows='';
      mats.forEach(function(mm, idx){
        var rv = trim(mm.remarks) || trim(mm.brand) || '';
        var imgTd = isLikelyImage(mm.image||mm.imageUrl) ? thumbHtml(mm.image||mm.imageUrl) : '';
        trows += '<tr style="height:35px;">'+
          '<td style="border:1px solid #ddd;padding:4px;text-align:center;font-size:0.8em;">'+(idx+1)+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(mm.tabName||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;font-weight:bold;">'+(mm.material||mm.category||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(mm.area||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(mm.item||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+rv+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;text-align:center;">'+imgTd+'</td>'+
          '</tr>';
      });
      var miniHtml = '<img src="'+miniImg+'" style="max-width:100%;height:120px;object-fit:cover;border-radius:5px;">';
      div.innerHTML =
        '<h4 style="margin-bottom:15px;color:#2c3e50;">'+title+'</h4>'+
        '<div style="display:flex;gap:20px;align-items:flex-start;">'+
          '<div style="flex:2;"><img src="'+scene.data+'" style="max-width:100%;height:200px;object-fit:cover;border-radius:5px;"></div>'+
          '<div style="flex:1;"><h5>미니맵</h5>'+miniHtml+'</div>'+
        '</div>'+
        '<table style="width:100%;margin-top:15px;border-collapse:collapse;font-size:0.85em;">'+
          '<thead><tr style="background:#f8f9fa;height:30px;">'+
            '<th style="border:1px solid #ddd;padding:5px;width:40px;font-size:0.8em;">No.</th>'+
            '<th style="border:1px solid #ddd;padding:5px;width:120px;font-size:0.8em;">탭명</th>'+
            '<th style="border:1px solid #ddd;padding:5px;width:120px;font-size:0.8em;">MATERIAL</th>'+
            '<th style="border:1px solid #ddd;padding:5px;width:60px;font-size:0.8em;">AREA</th>'+
            '<th style="border:1px solid #ddd;padding:5px;font-size:0.8em;">ITEM</th>'+
            '<th style="border:1px solid #ddd;padding:5px;font-size:0.8em;">REMARKS</th>'+
            '<th style="border:1px solid #ddd;padding:5px;font-size:0.8em;">IMAGE</th>'+
          '</tr></thead><tbody>'+trows+'</tbody></table>';
      preview.appendChild(div);

      progress++; updateProgress(progress/total*100);
    }

    pptx.writeFile({fileName:'착공도서.pptx'}).then(function(){
      updateProgress(100);
      showStatus('PPT가 성공적으로 생성되었습니다! 파일이 다운로드됩니다.','success');
      document.getElementById('progress').style.display='none';
      document.getElementById('generateBtn').disabled=false;
      document.getElementById('previewArea').style.display='block';
    });
  }catch(err){
    showStatus('PPT 생성 중 오류: '+err.message,'error');
    document.getElementById('progress').style.display='none';
    document.getElementById('generateBtn').disabled=false;
  }
}

/* render minimap to dataURL (PNG) */
function renderMinimapForExport(sceneIndex, outW, outH){
  var off=document.createElement('canvas');
  var W=(outW||3)*96, H=(outH||2)*96; off.width=W; off.height=H;
  var c2=off.getContext('2d');
  try{
    var img=minimapImgObj;
    if(img){
      var iw=img.naturalWidth, ih=img.naturalHeight;
      var scale=Math.min(W/iw,H/ih), dw=iw*scale, dh=ih*scale;
      var dx=(W-dw)/2, dy=(H-dh)/2;
      c2.fillStyle='#fff'; c2.fillRect(0,0,W,H);
      c2.drawImage(img,dx,dy,dw,dh);
      var box=appState.minimapBoxes[sceneIndex];
      if(box){
        c2.save();
        c2.strokeStyle='red'; c2.lineWidth=3; c2.fillStyle='rgba(255,0,0,0.12)';
        var rx=box.x*W, ry=box.y*H, rw=box.w*W, rh=box.h*H;
        c2.fillRect(rx,ry,rw,rh); c2.strokeRect(rx,ry,rw,rh);
        c2.restore();
      }
    }
  }catch(e){}
  return off.toDataURL('image/png');
}

/* UI helpers */
function updateProgress(pct){ document.getElementById('progressFill').style.width = pct+'%'; }
function showStatus(msg,type){
  var el=document.getElementById('status'); el.textContent=msg; el.className='status '+type; el.style.display='block';
  if(type==='success') setTimeout(function(){ el.style.display='none'; }, 5000);
}
