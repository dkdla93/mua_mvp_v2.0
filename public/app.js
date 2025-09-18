/* ===== Global State ===== */
var appState = {
  excelData: null,
  allSheets: {},
  currentSheet: null,
  minimapImage: null,
  sceneImages: [],
  materials: [],                 // { id, tabName, material, area, item, remarks, brand, imageUrl, image, category }
  sceneMaterialMapping: {},      // sceneIdx -> [materialId]
  currentSelectedScene: 0,
  minimapBoxes: {}               // sceneIdx -> {x,y,w,h} (0..1)
};

var canvas, ctx, isDrawing=false, startX=0, startY=0, currentRect=null, minimapImgObj=null;

/* ===== Env helpers ===== */
function isMac(){ return /Macintosh|Mac OS X/.test(navigator.userAgent); }
function koFont(){ return isMac() ? "Apple SD Gothic Neo" : "Malgun Gothic"; }

/* ===== Boot ===== */
document.addEventListener("DOMContentLoaded", function(){ initializeEventListeners(); });

function initializeEventListeners(){
  document.getElementById('excelFile').addEventListener('change', handleExcelUpload);
  document.getElementById('minimapFile').addEventListener('change', handleMinimapUpload);
  document.getElementById('sceneFiles').addEventListener('change', handleSceneUpload);
  document.getElementById('generateBtn').addEventListener('click', generatePPT);
}

/* =======================================================================
   Excel Handling (robust row parser: key can be in any of B/C/D; value is
   in the next non-empty cell to the right)
   ======================================================================= */
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

/* --- helpers: robust key/value finder in a row --- */
function findKeyInRow(row, keywords){
  // return {idx, key} for first match among keywords
  for(var c=0;c<row.length;c++){
    var v = (row[c]||'').toString().trim().toUpperCase();
    for(var k=0;k<keywords.length;k++){
      if(v.indexOf(keywords[k])!==-1) return {idx:c, key:keywords[k]};
    }
  }
  return null;
}
function valueRightOf(row, keyIdx){
  // prefer the next non-empty cell to the right (keyIdx+1 or +2)
  for(var c=keyIdx+1; c<Math.min(row.length, keyIdx+3); c++){
    var vv = row[c];
    if(typeof vv!=='undefined' && vv!==null && (vv.toString().trim()!=="")) return vv;
  }
  // fallback: a known column (C=2 or D=3) if exists
  if(row[2] && row[2].toString().trim()!=="") return row[2];
  if(row[3] && row[3].toString().trim()!=="") return row[3];
  return '';
}

function parseExcelData(){
  appState.materials = [];
  var current = null;
  var currentCategory = '';
  var sheets = Object.keys(appState.allSheets);

  for(var s=0; s<sheets.length; s++){
    var sheetName = sheets[s];
    if(/^A\./.test(sheetName)) continue;
    var data = appState.allSheets[sheetName];

    currentCategory = '';

    for(var r=1; r<data.length; r++){
      var row = data[r]; if(!row || row.length<2) continue;

      // Update category block header (left-most large title in col A)
      var left = (row[0]||'').toString().trim().toUpperCase();
      if(left && (left.indexOf('MATERIAL')!==-1 || left.indexOf('SWITCH')!==-1 || left.indexOf('LIGHT')!==-1))
        currentCategory = (row[0]||'').toString().trim();

      var hit = findKeyInRow(row, ['AREA','MATERIAL','ITEM','REMARKS','REMARK','IMAGE']);
      if(!hit) continue;

      if(hit.key==='AREA'){
        // start a new record
        if(current) appState.materials.push(current);
        current = {
          id: appState.materials.length+1,
          tabName: sheetName,
          displayId: '#'+sheetName,
          category: currentCategory || 'MATERIAL',
          material: '',
          area: (valueRightOf(row, hit.idx) || ''),
          item: '',
          remarks: '',
          brand: '',
          imageUrl: '',
          image: null
        };
      } else if(hit.key==='MATERIAL' && current){
        current.material = valueRightOf(row, hit.idx) || '';
      } else if(hit.key==='ITEM' && current){
        current.item = valueRightOf(row, hit.idx) || '';
      } else if((hit.key==='REMARKS' || hit.key==='REMARK') && current){
        current.remarks = valueRightOf(row, hit.idx) || '';
      } else if(hit.key==='IMAGE' && current){
        var imgVal = valueRightOf(row, hit.idx) || '';
        current.imageUrl = imgVal;
        current.image    = imgVal || null;
      }
    }
    if(current){ appState.materials.push(current); current=null; }
  }
  setTimeout(checkAllFilesUploaded,100);
}

/* =======================================================================
   Minimap handling
   ======================================================================= */
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

/* =======================================================================
   Scene images
   ======================================================================= */
function handleSceneUpload(e){
  var files = Array.from(e.target.files); if(files.length===0) return;
  appState.sceneImages = [];
  var loaded=0;
  for(var i=0;i<files.length;i++){
    (function(idx,f){
      var r=new FileReader();
      r.onload=function(ev){
        appState.sceneImages.push({name:f.name,data:ev.target.result,index:idx});
        loaded++; if(loaded===files.length){ displaySceneInfo(); checkAllFilesUploaded(); }
      };
      r.readAsDataURL(f);
    })(i,files[i]);
  }
}

function displaySceneInfo(){
  var html = '<strong>업로드 완료:</strong> '+appState.sceneImages.length+'개 장면 이미지<br>';
  html += '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:10px;">';
  for(var i=0;i<appState.sceneImages.length;i++){
    var s=appState.sceneImages[i];
    html += '<div style="text-align:center;"><img src="'+s.data+'" style="width:80px;height:60px;object-fit:cover;border-radius:3px;"><div style="font-size:0.8em;margin-top:5px;">'+s.name+'</div></div>';
  }
  html+='</div>';
  var el = document.getElementById('sceneInfo'); el.innerHTML = html; el.style.display='block';
  setTimeout(checkAllFilesUploaded,100);
}

/* =======================================================================
   UI build
   ======================================================================= */
function checkAllFilesUploaded(){
  var hasExcel = appState.currentSheet!==null && appState.materials.length>0;
  var hasMinimap = !!appState.minimapImage;
  var hasScenes = appState.sceneImages.length>0;

  if(hasExcel && hasMinimap && hasScenes){
    try{
      createMaterialInterface();
      document.getElementById('matchingStep').style.display='block';
      document.getElementById('generateStep').style.display='block';
      document.getElementById('minimapDrawWrap').style.display='block';
      showStatus('모든 파일이 업로드되었습니다! 장면별 자재와 미니맵 위치를 지정해주세요.','success');
    }catch(e){
      showStatus('인터페이스 생성 중 오류: '+e.message,'error');
    }
  }else{
    var missing=[]; if(!hasExcel) missing.push('엑셀 파일'); if(!hasMinimap) missing.push('미니맵 이미지'); if(!hasScenes) missing.push('장면 이미지들');
    if(missing.length>0) showStatus('누락된 항목: '+missing.join(', '),'error');
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
  for(var i=0;i<appState.sceneImages.length;i++){
    var s=appState.sceneImages[i];
    var div=document.createElement('div');
    div.className='scene-item-selector';
    if(i===0) div.classList.add('active');
    div.innerHTML='<img src="'+s.data+'" alt="'+s.name+'" class="scene-thumb">'+
      '<div><div style="font-weight:bold;">'+s.name+'</div>'+
      '<div style="font-size:0.8em;color:#7f8c8d;">장면 '+(i+1)+'</div>'+
      '<div style="font-size:0.8em;color:#27ae60;" id="scene-count-'+i+'">자재 0개 선택됨</div></div>';
    (function(idx){ div.onclick=function(){ selectScene(idx); }; })(i);
    c.appendChild(div);
  }
}

function createMaterialTabFilter(){
  var c=document.getElementById('materialTabButtons'); c.innerHTML='';
  var tabs={}; for(var i=0;i<appState.materials.length;i++){ var m=appState.materials[i]; if(m.tabName) tabs[m.tabName]=true; }
  var names=Object.keys(tabs);

  var allBtn=document.createElement('button'); allBtn.className='material-tab-filter-btn active';
  allBtn.textContent='전체 보기'; allBtn.onclick=function(){ filterMaterialsByTab(null); updateActiveFilterButton(allBtn); };
  c.appendChild(allBtn);

  names.forEach(function(tabName){
    var btn=document.createElement('button'); btn.className='material-tab-filter-btn'; btn.textContent=tabName;
    btn.onclick=function(){ filterMaterialsByTab(tabName); updateActiveFilterButton(btn); };
    c.appendChild(btn);
  });
}

function updateActiveFilterButton(activeBtn){
  var buttons=document.querySelectorAll('.material-tab-filter-btn');
  for(var i=0;i<buttons.length;i++) buttons[i].classList.remove('active');
  activeBtn.classList.add('active');
}

function createMaterialTable(){
  var c=document.getElementById('materialTableList'); c.innerHTML='';
  if(appState.materials.length===0){
    c.innerHTML='<tr><td colspan="9" style="text-align:center;color:#999;padding:20px;">자재가 없습니다.</td></tr>'; return;
  }
  for(var i=0;i<appState.materials.length;i++){
    var m=appState.materials[i];
    var tr=document.createElement('tr'); tr.className='material-row'; tr.setAttribute('data-material-id', m.id); tr.style.height='35px';

    var imageHtml=(m.image||m.imageUrl)
      ? '<img src="'+(m.image||m.imageUrl)+'" style="width:40px;height:30px;object-fit:cover;border-radius:3px;" alt="자재">'
      : '<div style="width:40px;height:30px;background:#f0f0f0;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:8px;color:#999;">없음</div>';

    var remarksValue='';
    if(m.remarks && m.remarks.toString().trim()!=='') remarksValue=m.remarks.toString().trim();
    else if(m.brand && m.brand.toString().trim()!=='') remarksValue=m.brand.toString().trim();

    tr.innerHTML =
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;font-size:0.8em;">'+(i+1)+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(m.tabName||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;font-weight:bold;">'+(m.material||m.category||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(m.area||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(m.item||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(remarksValue||'')+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;">'+imageHtml+'</td>'+
      '<td style="border:1px solid #ddd;padding:4px;text-align:center;"><input type="checkbox" id="checkbox_'+m.id+'" style="width:16px;height:16px;"></td>';

    (function(mid){
      var cb = tr.querySelector('input[type="checkbox"]');
      cb.addEventListener('change', function(e){ toggleMaterial(mid, e.target.checked); });
      tr.addEventListener('click', function(e){
        if(e.target.type!=='checkbox'){
          var cb2=document.getElementById('checkbox_'+mid);
          if(cb2){ cb2.checked=!cb2.checked; toggleMaterial(mid, cb2.checked); }
        }
      });
    })(m.id);

    c.appendChild(tr);
  }
}

function selectScene(sceneIndex){
  var items=document.querySelectorAll('.scene-item-selector');
  for(var i=0;i<items.length;i++) items[i].classList.remove('active');
  if(items[sceneIndex]) items[sceneIndex].classList.add('active');
  appState.currentSelectedScene=sceneIndex;
  var name=appState.sceneImages[sceneIndex].name;
  document.getElementById('currentSceneTitle').textContent = name.replace(/\.[^/.]+$/, '');
  updateCheckboxStates(); drawMinimapForScene();
}

function toggleMaterial(materialId, checked){
  if(!appState.sceneMaterialMapping[appState.currentSelectedScene]) appState.sceneMaterialMapping[appState.currentSelectedScene]=[];
  var arr=appState.sceneMaterialMapping[appState.currentSelectedScene];
  var idx=arr.indexOf(materialId);
  if(checked && idx===-1) arr.push(materialId);
  else if(!checked && idx>-1) arr.splice(idx,1);
  updateCheckboxStates();
}

function filterMaterialsByTab(tab){
  var rows=document.querySelectorAll('.material-row');
  for(var i=0;i<rows.length;i++){
    var row=rows[i];
    var mid=parseInt(row.getAttribute('data-material-id'),10);
    var m=null; for(var j=0;j<appState.materials.length;j++){ if(appState.materials[j].id===mid){ m=appState.materials[j]; break; } }
    if(m) row.style.display = (tab===null || m.tabName===tab) ? 'table-row' : 'none';
  }
}

function updateCheckboxStates(){
  var selectedIds=appState.sceneMaterialMapping[appState.currentSelectedScene]||[];
  for(var i=0;i<appState.materials.length;i++){
    var m=appState.materials[i], cb=document.getElementById('checkbox_'+m.id);
    if(cb){ var row=cb.closest('tr');
      if(selectedIds.indexOf(m.id)>-1){ cb.checked=true; if(row) row.style.backgroundColor='#e8f5e8'; }
      else{ cb.checked=false; if(row) row.style.backgroundColor=''; }
    }
  }
  var el=document.getElementById('scene-count-'+appState.currentSelectedScene);
  if(el) el.textContent='자재 '+selectedIds.length+'개 선택됨';
}

/* =======================================================================
   Minimap drawing
   ======================================================================= */
function setupMinimapCanvas(){
  canvas=document.getElementById('minimapCanvas'); ctx=canvas.getContext('2d');
  canvas.onmousedown=function(e){ isDrawing=true; var r=canvas.getBoundingClientRect(); startX=e.clientX-r.left; startY=e.clientY-r.top; currentRect={x:startX,y:startY,w:0,h:0}; };
  canvas.onmousemove=function(e){ if(!isDrawing)return; var r=canvas.getBoundingClientRect(); var x=e.clientX-r.left; var y=e.clientY-r.top; currentRect.w=x-startX; currentRect.h=y-startY; drawMinimapForScene(currentRect); };
  canvas.onmouseup=function(){ if(!isDrawing)return; isDrawing=false; if(currentRect){ var nx=clamp(Math.min(currentRect.x,currentRect.x+currentRect.w)/canvas.width,0,1); var ny=clamp(Math.min(currentRect.y,currentRect.y+currentRect.h)/canvas.height,0,1); var nw=clamp(Math.abs(currentRect.w)/canvas.width,0,1); var nh=clamp(Math.abs(currentRect.h)/canvas.height,0,1); appState.minimapBoxes[appState.currentSelectedScene]={x:nx,y:ny,w:nw,h:nh}; drawMinimapForScene(); } };
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

/* =======================================================================
   PPT Generation (PptxGenJS)  — overflow guards added
   ======================================================================= */
function generatePPT(){
  if(Object.keys(appState.sceneMaterialMapping).length===0){ showStatus('장면별 자재를 선택해주세요.','error'); return; }
  if(!appState.minimapImage){ showStatus('미니맵 이미지를 업로드해주세요.','error'); return; }

  document.getElementById('progress').style.display='block';
  document.getElementById('generateBtn').disabled=true;

  try{
    var pptx = new PptxGenJS();
    pptx.defineLayout({name:'LAYOUT_16x9', width:10, height:5.625});
    pptx.layout = 'LAYOUT_16x9';

    var total = appState.sceneImages.length, progress=0;
    var preview = document.getElementById('slidePreview'); preview.innerHTML='';

    for(var i=0;i<appState.sceneImages.length;i++){
      var scene = appState.sceneImages[i];
      var selectedIds = appState.sceneMaterialMapping[i] || [];
      var mats=[]; for(var j=0;j<selectedIds.length;j++){ var mid=selectedIds[j]; for(var k=0;k<appState.materials.length;k++){ if(appState.materials[k].id===mid){ mats.push(appState.materials[k]); break; } } }

      var slide = pptx.addSlide();
      var title = scene.name.replace(/\.[^/.]+$/,'');
      slide.addText(title, { x:.5, y:.3, fontSize:20, bold:true, color:'363636', fontFace: koFont() });

      // Scene image (adaptive height)
      var topY=0.8;
      var sceneH = (mats.length>6 ? 3.0 : 3.4);
      var sceneW = 6.1;
      slide.addImage({ data:scene.data, x:.5, y:topY, w:sceneW, h:sceneH, rounding:6 });

      // Minimap (larger but clamped to slide bounds)
      var miniImg = renderMinimapForExport(i, 3.0, 2.2);
      var mini = { x:6.7, y:topY+0.2, w:3.0, h:2.2 };
      // clamp to slide width (10")
      if(mini.x + mini.w > 9.7){ mini.w = 9.7 - mini.x; if(mini.w<2.2) mini.w=2.2; }
      slide.addText('미니맵', { x:mini.x, y:topY, fontSize:12, bold:true, color:'555555', fontFace: koFont() });
      slide.addImage({ data:miniImg, x:mini.x, y:mini.y, w:mini.w, h:mini.h, rounding:4 });

      // Table placement (avoid overflow)
      var space=0.3, tableY = topY + sceneH + space;
      var available = 5.2 - tableY;                       // keep ~0.4" bottom margin
      var rowCount = mats.length + 1;
      var rowH = Math.max(0.18, Math.min(0.30, available / Math.max(rowCount,1)));
      var tableX = 0.5, tableW = 9.0;                     // narrower to avoid bleed
      var rows = [['No.','탭명','MATERIAL','AREA','ITEM','REMARKS','IMAGE']];

      for(var r=0;r<mats.length;r++){
        var m=mats[r];
        var remarks = (m.remarks && m.remarks.toString().trim()!==''
                      && m.remarks.toString().trim().toUpperCase()!=='REMARKS')
                      ? m.remarks.toString().trim()
                      : (m.brand && m.brand.toString().trim()!=='' ? m.brand.toString().trim() : '');
        rows.push([ String(r+1), m.tabName||'', (m.material||m.category||''), m.area||'', m.item||'', (m.remarks||remarks||''), (m.imageUrl?'있음':'') ]);
      }

      slide.addTable(rows, {
        x:tableX, y:tableY, w:tableW,
        colW:[0.6,1.4,1.4,1.3,3.0,1.5,0.8],
        fontSize:10, border:{type:'solid',color:'CCCCCC',pt:1}, fill:'FFFFFF',
        valign:'middle', margin:2, color:'333333', rowH:rowH
      });

      // HTML Preview (for step 3)
      var div=document.createElement('div');
      div.style.cssText='margin-bottom:30px;padding:20px;border:1px solid #ddd;border-radius:8px;background:#fafafa;';
      var trows='';
      mats.forEach(function(mm,idx){
        var rv=(mm.remarks&&mm.remarks.toString().trim()!==''&&mm.remarks.toString().trim().toUpperCase()!=='REMARKS')
                ? mm.remarks.toString().trim()
                : (mm.brand&&mm.brand.toString().trim()!=='' ? mm.brand.toString().trim() : '');
        trows += '<tr style="height:35px;">'+
          '<td style="border:1px solid #ddd;padding:4px;text-align:center;font-size:0.8em;">'+(idx+1)+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(mm.tabName||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;font-weight:bold;">'+(mm.material||mm.category||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(mm.area||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(mm.item||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(rv||mm.remarks||'')+'</td>'+
          '<td style="border:1px solid #ddd;padding:4px;font-size:0.8em;">'+(mm.imageUrl?'있음':'')+'</td></tr>';
      });

      div.innerHTML =
        '<h4 style="margin-bottom:15px;color:#2c3e50;">'+title+'</h4>'+
        '<div style="display:flex;gap:20px;align-items:flex-start;">'+
        '<div style="flex:2;"><img src="'+scene.data+'" style="max-width:100%;height:200px;object-fit:cover;border-radius:5px;"></div>'+
        '<div style="flex:1;"><h5>미니맵</h5><img src="'+miniImg+'" style="max-width:100%;height:120px;object-fit:cover;border-radius:5px;"></div>'+
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
function updateProgress(pct){ document.getElementById('progressFill').style.width = (pct)+'%'; }
function showStatus(msg,type){
  var el=document.getElementById('status'); el.textContent=msg; el.className='status '+type; el.style.display='block';
  if(type==='success') setTimeout(function(){ el.style.display='none'; }, 5000);
}
