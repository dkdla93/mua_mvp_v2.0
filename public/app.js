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

// 특정 시트에서 주어진 헤더명이 들어간 "열" 인덱스를 찾는다 (상단 20행 탐색)
function getColIndexByHeader(sheetData, headerNames){
  if (!sheetData || !sheetData.length) return -1;
  const names = headerNames.map(h=>String(h).toUpperCase());
  const scanRows = Math.min(20, sheetData.length);
  for (let r=0; r<scanRows; r++){
    const row = sheetData[r] || [];
    for (let c=0; c<row.length; c++){
      const v = trim(row[c]).toUpperCase();
      if (!v) continue;
      for (let i=0;i<names.length;i++){
        if (v.indexOf(names[i]) !== -1) return c;
      }
    }
  }
  return -1;
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
  appState.materials = [];

  var sheets = Object.keys(appState.allSheets);

  for (var s = 0; s < sheets.length; s++) {
    var sheetName = sheets[s];
    // 표지/서문(A.MAIN 등) 스킵 규칙 유지
    if (/^A\./.test(sheetName)) continue;

    var data = appState.allSheets[sheetName];
    if (!data || !data.length) continue;

    // === 열 인덱스 탐색(헤더 행이 꼭 1행이 아니어도 OK) ===
    var remarksCol = getColIndexByHeader(data, ['REMARKS','REMARK']);
    var imageCol   = getColIndexByHeader(data, ['IMAGE']);

    var current = null;
    var currentCategory   = '';  // MATERIAL / SWITCH / LIGHT 등 (왼쪽 큰 카테고리)
    var currentGroupLabel = '';  // A열의 구역 라벨 (예: WALL COVERING [도배], PAINT [도장] 등)

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row || row.length < 2) continue;

      // 왼쪽 큰 카테고리(MATERIAL / SWITCH / LIGHT 등) 추적
      var left0Upper = trim(row[0]).toUpperCase();
      if (left0Upper &&
          (left0Upper.indexOf('MATERIAL') !== -1 ||
           left0Upper.indexOf('SWITCH')   !== -1 ||
           left0Upper.indexOf('LIGHT')    !== -1)) {
        // 원문 그대로 보존
        currentCategory = trim(row[0]);
      }

      // A열의 그룹 라벨(예: WALL COVERING [도배], PAINT [도장] 등) 갱신
      // 헤더성 단어(MATERIAL/DESCRIPTION)는 라벨로 취급하지 않음
      var a0 = trim(row[0]);
      if (a0 && !/^(MATERIAL|DESCRIPTION)$/i.test(a0)) {
        currentGroupLabel = a0;
      }

      // 행 내 키 감지 (DESCRIPTION은 키로 취급하지 않음)
      var hit = findKeyInRow(row, ['AREA','MATERIAL','ITEM','REMARKS','REMARK','IMAGE']);
      if (!hit) continue;

      if (hit.key === 'AREA') {
        // 새 항목 시작
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

        // AREA 행에서는 REMARKS/IMAGE를 강제 채우지 않음 (ITEM 행에서 확정)

      } else if (hit.key === 'MATERIAL' && current) {
        // 헤더 오염 방지: 오른쪽 값이 헤더성 단어면 무시
        var mv = trim(valueRightOf(row, hit.idx) || '');
        if (mv && !/^(DESCRIPTION|IMAGE|EA|REMARKS)$/i.test(mv)) {
          current.material = mv;
        }

        // 보조로 같은 줄의 REMARKS/IMAGE도 반영(있다면) — 단, 이후 ITEM에서 다시 덮어씀
        if (remarksCol > -1) {
          var rem2 = trim(row[remarksCol] || '');
          if (rem2) current.remarks = rem2;
        }
        if (imageCol > -1) {
          var imgCell2 = row[imageCol] != null ? String(row[imageCol]) : '';
          var imgUrl2  = extractImageUrl(imgCell2);
          if (isLikelyImage(imgUrl2)) { current.imageUrl = imgUrl2; current.image = imgUrl2; }
        }

      } else if (hit.key === 'ITEM' && current) {
        // ITEM 값
        current.item = trim(valueRightOf(row, hit.idx) || '');

        // ✅ REMARKS의 정답: ITEM 행의 같은 줄 remarksCol(E열)
        if (remarksCol > -1) {
          var rem3 = trim(row[remarksCol] || '');
          if (rem3) current.remarks = rem3;
        }

        // 같은 줄 IMAGE도 함께
        if (imageCol > -1) {
          var imgCell3 = row[imageCol] != null ? String(row[imageCol]) : '';
          var imgUrl3  = extractImageUrl(imgCell3);
          if (isLikelyImage(imgUrl3)) { current.imageUrl = imgUrl3; current.image = imgUrl3; }
        }

      } else if ((hit.key === 'REMARKS' || hit.key === 'REMARK') && current) {
        // (안전망) 어떤 시트에서 REMARKS가 행 키로 오는 경우도 커버
        var rv = trim(valueRightOf(row, hit.idx) || '');
        if (rv) current.remarks = rv;

      } else if (hit.key === 'IMAGE' && current) {
        // (안전망) 어떤 시트에서 IMAGE가 행 키로 오는 경우도 커버
        var val = trim(valueRightOf(row, hit.idx) || '');
        var url = extractImageUrl(val);
        if (isLikelyImage(url)) { current.imageUrl = url; current.image = url; }
      }
    } // rows

    // 마지막 항목 푸시
    if (current) {
      if (!trim(current.material)) {
        current.material = currentGroupLabel || currentCategory || sheetName || '';
      }
      appState.materials.push(current);
      current = null;
    }
  } // sheets

  // 파싱 후 UI 진행
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
