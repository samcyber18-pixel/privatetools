'use strict';

const modal = document.querySelector('#tool-modal');
const content = document.querySelector('#tool-content');
const openButtons = document.querySelectorAll('[data-open-tool]');
let lastFocused = null;

const MIME_EXT = { 'image/jpeg':'jpg', 'image/png':'png', 'image/webp':'webp' };

function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) return '—';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(bytes < 10240 ? 1 : 0)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(bytes < 10485760 ? 1 : 0)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
}
function safeName(name) { return name.replace(/[\\/:*?"<>|\x00-\x1F]/g, '-').trim() || 'file'; }
function result(el, message, isError = false) { el.innerHTML = `<div class="result${isError ? ' error' : ''}">${message}</div>`; }
function setBusy(button, busy, idleLabel, busyLabel) { button.disabled = busy; button.textContent = busy ? busyLabel : idleLabel; }
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob); const a = document.createElement('a');
  a.href = url; a.download = safeName(name); document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
function isImage(file) { return file && ['image/jpeg','image/png','image/webp'].includes(file.type); }
function isPdf(file) { return file && (file.type === 'application/pdf' || /\.pdf$/i.test(file.name)); }

function baseTool(title, subtitle, inner) {
  return `<p class="eyebrow">PRIVATE TOOLS</p><h2 id="modal-title" class="tool-title">${title}</h2><p id="modal-subtitle" class="tool-subtitle">${subtitle}</p>${inner}`;
}
function openTool(name) {
  const renderers = { 'image-compress':renderImageCompress, 'image-convert':renderImageConvert, 'pdf-merge':renderPdfMerge, 'pdf-split':renderPdfSplit };
  if (!renderers[name]) return;
  lastFocused = document.activeElement;
  content.innerHTML = renderers[name]();
  modal.classList.remove('hidden'); modal.setAttribute('aria-hidden','false'); document.body.style.overflow = 'hidden';
  bindTool(name);
  content.querySelector('input,select,button')?.focus();
}
function closeTool() {
  modal.classList.add('hidden'); modal.setAttribute('aria-hidden','true'); content.innerHTML = ''; document.body.style.overflow = '';
  lastFocused?.focus?.(); lastFocused = null;
}
openButtons.forEach(btn => btn.addEventListener('click', () => openTool(btn.dataset.openTool)));
document.querySelectorAll('[data-close-modal]').forEach(el => el.addEventListener('click', closeTool));
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !modal.classList.contains('hidden')) closeTool(); });

function bindDropZone(zone, input, onFiles, multiple = true) {
  const accept = fileList => {
    const files = Array.from(fileList || []); onFiles(multiple ? files : files.slice(0,1));
  };
  input.addEventListener('change', () => accept(input.files));
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('dragover'); accept(e.dataTransfer.files); });
}
function renderFileList(el, files, removeAt = null) {
  if (!files.length) { el.innerHTML = '<p class="muted small-note">No files selected yet.</p>'; return; }
  el.innerHTML = files.map((f,i) => `<div class="file-row"><span>${i+1}</span><span class="name" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</span><span class="size muted">${formatBytes(f.size)}</span>${removeAt ? `<button type="button" data-remove-index="${i}" aria-label="Remove ${escapeHtml(f.name)}">×</button>` : ''}</div>`).join('');
  removeAt && el.querySelectorAll('[data-remove-index]').forEach(btn => btn.addEventListener('click', () => removeAt(Number(btn.dataset.removeIndex))));
}

async function loadImage(file) {
  if (typeof createImageBitmap === 'function') {
    const bitmap = await createImageBitmap(file);
    return { source:bitmap, width:bitmap.width, height:bitmap.height, close:()=>bitmap.close() };
  }
  const url = URL.createObjectURL(file);
  const img = await new Promise((resolve,reject) => { const image = new Image(); image.onload=()=>resolve(image); image.onerror=()=>reject(new Error('The image could not be decoded.')); image.src=url; });
  URL.revokeObjectURL(url);
  return { source:img, width:img.naturalWidth, height:img.naturalHeight, close:()=>{} };
}
async function imageBlob(file, mime, quality, width, height) {
  const image = await loadImage(file);
  try {
    const canvas = document.createElement('canvas'); canvas.width = Math.max(1, Math.round(width || image.width)); canvas.height = Math.max(1, Math.round(height || image.height));
    const ctx = canvas.getContext('2d'); if (!ctx) throw new Error('Canvas is not available in this browser.');
    ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
    if (mime === 'image/jpeg') { ctx.fillStyle = '#fff'; ctx.fillRect(0,0,canvas.width,canvas.height); }
    ctx.drawImage(image.source,0,0,canvas.width,canvas.height);
    const blob = await new Promise(resolve => canvas.toBlob(resolve,mime,quality));
    if (!blob) throw new Error('The browser could not create the output image.');
    return { blob, width:canvas.width, height:canvas.height };
  } finally { image.close(); }
}
function resizeDimensions(ow, oh, widthValue, heightValue, keepRatio) {
  const w = Number(widthValue) > 0 ? Math.round(Number(widthValue)) : null;
  const h = Number(heightValue) > 0 ? Math.round(Number(heightValue)) : null;
  if (!w && !h) return [ow,oh]; if (!keepRatio) return [w || ow, h || oh];
  if (w && !h) return [w, Math.max(1, Math.round(oh*w/ow))];
  if (!w && h) return [Math.max(1, Math.round(ow*h/oh)),h];
  return [w,h];
}

function renderImageCompress() { return baseTool('Compress images','Choose one or more images. Everything is processed locally in your browser.',`
  <div class="drop-zone"><label class="file-label" for="compress-files">Choose images</label><input id="compress-files" type="file" accept="image/jpeg,image/png,image/webp" multiple><p class="muted">JPG, PNG or WebP · drag and drop supported</p></div>
  <div id="compress-list" class="file-list"></div>
  <div class="form-grid"><div class="field"><label for="compress-format">Output format</label><select id="compress-format"><option value="same">Keep original format</option><option value="image/webp">WebP</option><option value="image/jpeg">JPG</option></select></div><div class="field"><label for="compress-quality">Quality <span id="quality-value" class="range-value">80%</span></label><input id="compress-quality" type="range" min="0.1" max="1" step="0.05" value="0.8"></div></div>
  <p class="small-note">PNG uses lossless browser encoding, so the quality slider has no lossy effect on PNG. Choose WebP or JPG for stronger size reduction.</p>
  <div class="action-row"><button class="primary-button" id="compress-run">Compress and download</button><button class="secondary-button" id="compress-clear">Clear</button></div>
  <div id="compress-progress" class="progress-wrap hidden"><div class="progress"><div></div></div><p class="small-note" id="compress-progress-text"></p></div><div id="compress-result" aria-live="polite"></div>`); }
function renderImageConvert() { return baseTool('Convert & resize images','Convert between JPG, PNG and WebP and optionally resize the output.',`
  <div class="drop-zone"><label class="file-label" for="convert-file">Choose an image</label><input id="convert-file" type="file" accept="image/jpeg,image/png,image/webp"><p class="muted">JPG, PNG or WebP · drag and drop supported</p></div>
  <div id="convert-info" class="muted small-note">No image selected.</div><div class="form-grid"><div class="field"><label for="convert-format">Output format</label><select id="convert-format"><option value="image/jpeg">JPG</option><option value="image/png">PNG</option><option value="image/webp">WebP</option></select></div><div class="field"><label for="convert-quality">Quality <span id="convert-quality-value" class="range-value">90%</span></label><input id="convert-quality" type="range" min="0.1" max="1" step="0.05" value="0.9"></div><div class="field"><label for="convert-width">Width (optional)</label><input id="convert-width" type="number" min="1" placeholder="Original"></div><div class="field"><label for="convert-height">Height (optional)</label><input id="convert-height" type="number" min="1" placeholder="Original"></div></div>
  <label class="checkbox-field"><input id="convert-keep-ratio" type="checkbox" checked> Keep aspect ratio when one dimension is supplied</label><div class="action-row"><button class="primary-button" id="convert-run">Convert and download</button><button class="secondary-button" id="convert-clear">Clear</button></div><div id="convert-result" aria-live="polite"></div>`); }
function renderPdfMerge() { return baseTool('Merge PDFs','Combine PDFs locally. The files remain in your browser during processing.',`
  <div class="drop-zone"><label class="file-label" for="merge-files">Choose PDFs</label><input id="merge-files" type="file" accept="application/pdf,.pdf" multiple><p class="muted">Select at least two PDFs · drag and drop supported</p></div><div id="merge-list" class="file-list"></div><p class="small-note">Files are merged in the order shown. Remove files with × before merging.</p><div class="action-row"><button class="primary-button" id="merge-run">Merge and download</button><button class="secondary-button" id="merge-clear">Clear</button></div><div id="merge-progress" class="progress-wrap hidden"><div class="progress"><div></div></div><p class="small-note" id="merge-progress-text"></p></div><div id="merge-result" aria-live="polite"></div>`); }
function renderPdfSplit() { return baseTool('Split a PDF','Pick pages and create a new PDF locally in your browser.',`
  <div class="drop-zone"><label class="file-label" for="split-file">Choose a PDF</label><input id="split-file" type="file" accept="application/pdf,.pdf"><p class="muted">One PDF · drag and drop supported</p></div><div id="split-info" class="muted small-note">No PDF selected.</div><div class="field"><label for="split-pages">Pages</label><input id="split-pages" type="text" inputmode="numeric" placeholder="Example: 1,3,5-7"><p class="small-note">Enter page numbers/ranges separated by commas. Example: 1,3,5-7.</p></div><div class="action-row"><button class="primary-button" id="split-run">Create PDF and download</button><button class="secondary-button" id="split-clear">Clear</button></div><div id="split-result" aria-live="polite"></div>`); }

function bindTool(name) { if(name==='image-compress') bindImageCompress(); if(name==='image-convert') bindImageConvert(); if(name==='pdf-merge') bindPdfMerge(); if(name==='pdf-split') bindPdfSplit(); }

function bindImageCompress() {
  const input=document.querySelector('#compress-files'),zone=input.closest('.drop-zone'),list=document.querySelector('#compress-list'),quality=document.querySelector('#compress-quality'),qualityValue=document.querySelector('#quality-value'),run=document.querySelector('#compress-run'),clear=document.querySelector('#compress-clear'),outFormat=document.querySelector('#compress-format'),resultEl=document.querySelector('#compress-result'),progress=document.querySelector('#compress-progress'),bar=progress.querySelector('.progress>div'),progressText=document.querySelector('#compress-progress-text');
  let files=[]; const update=()=>renderFileList(list,files,i=>{files.splice(i,1);update();});
  bindDropZone(zone,input,incoming=>{files=incoming.filter(isImage);update();resultEl.innerHTML='';});
  quality.addEventListener('input',()=>qualityValue.textContent=`${Math.round(Number(quality.value)*100)}%`);
  clear.addEventListener('click',()=>{files=[];input.value='';resultEl.innerHTML='';progress.classList.add('hidden');update();});
  run.addEventListener('click',async()=>{if(!files.length)return result(resultEl,'Choose at least one JPG, PNG or WebP image.',true);setBusy(run,true,'Compress and download','Processing…');progress.classList.remove('hidden');bar.style.width='0%';let before=0,after=0;try{for(let i=0;i<files.length;i++){const file=files[i];before+=file.size;const selected=outFormat.value;const mime=selected==='same'?file.type:selected;const output=await imageBlob(file,mime,Number(quality.value));after+=output.blob.size;downloadBlob(output.blob,`${file.name.replace(/\.[^.]+$/,'')}-compressed.${MIME_EXT[mime]}`);bar.style.width=`${Math.round((i+1)/files.length*100)}%`;progressText.textContent=`Processed ${i+1} of ${files.length}`;}const diff=((after/before)-1)*100;result(resultEl,`Done. ${formatBytes(before)} → ${formatBytes(after)} · ${diff<=0?`${Math.abs(diff).toFixed(0)}% smaller`:`${diff.toFixed(0)}% larger`}.`);}catch(err){result(resultEl,`Could not process the images: ${escapeHtml(err.message||'Unknown error')}`,true);}finally{setBusy(run,false,'Compress and download','Processing…');}});
  update();
}

function bindImageConvert() {
  const input=document.querySelector('#convert-file'),zone=input.closest('.drop-zone'),info=document.querySelector('#convert-info'),run=document.querySelector('#convert-run'),clear=document.querySelector('#convert-clear'),resultEl=document.querySelector('#convert-result'),quality=document.querySelector('#convert-quality'),qualityValue=document.querySelector('#convert-quality-value'); let file=null;
  const showInfo=async next=>{file=next&&isImage(next)?next:null;if(!file){info.textContent='No image selected.';return;}try{const image=await loadImage(file);info.textContent=`${file.name} · ${formatBytes(file.size)} · ${image.width}×${image.height}px`;image.close();}catch{info.textContent=`${file.name} · ${formatBytes(file.size)} · image dimensions unavailable`;}};
  bindDropZone(zone,input,incoming=>showInfo(incoming[0]),false); quality.addEventListener('input',()=>qualityValue.textContent=`${Math.round(Number(quality.value)*100)}%`);
  clear.addEventListener('click',()=>{file=null;input.value='';info.textContent='No image selected.';resultEl.innerHTML='';});
  run.addEventListener('click',async()=>{if(!file)return result(resultEl,'Choose a JPG, PNG or WebP image first.',true);setBusy(run,true,'Convert and download','Processing…');try{const image=await loadImage(file);const [width,height]=resizeDimensions(image.width,image.height,document.querySelector('#convert-width').value,document.querySelector('#convert-height').value,document.querySelector('#convert-keep-ratio').checked);image.close();const mime=document.querySelector('#convert-format').value;const output=await imageBlob(file,mime,Number(quality.value),width,height);downloadBlob(output.blob,`${file.name.replace(/\.[^.]+$/,'')}-converted.${MIME_EXT[mime]}`);result(resultEl,`Done. ${width}×${height}px · ${formatBytes(output.blob.size)}.`);}catch(err){result(resultEl,`Could not convert the image: ${escapeHtml(err.message||'Unknown error')}`,true);}finally{setBusy(run,false,'Convert and download','Processing…');}});
}

function bindPdfMerge() {
  const input=document.querySelector('#merge-files'),zone=input.closest('.drop-zone'),list=document.querySelector('#merge-list'),run=document.querySelector('#merge-run'),clear=document.querySelector('#merge-clear'),resultEl=document.querySelector('#merge-result'),progress=document.querySelector('#merge-progress'),bar=progress.querySelector('.progress>div'),progressText=document.querySelector('#merge-progress-text'); let files=[];
  const update=()=>renderFileList(list,files,i=>{files.splice(i,1);update();});
  bindDropZone(zone,input,incoming=>{files=incoming.filter(isPdf);update();resultEl.innerHTML='';});
  clear.addEventListener('click',()=>{files=[];input.value='';resultEl.innerHTML='';progress.classList.add('hidden');update();});
  run.addEventListener('click',async()=>{if(files.length<2)return result(resultEl,'Choose at least two PDF files.',true);if(!window.PDFLib)return result(resultEl,'The PDF engine did not load. Check your connection and reload the page.',true);setBusy(run,true,'Merge and download','Merging…');progress.classList.remove('hidden');bar.style.width='0%';try{const output=await PDFLib.PDFDocument.create();for(let i=0;i<files.length;i++){const source=await PDFLib.PDFDocument.load(await files[i].arrayBuffer());const pages=await output.copyPages(source,source.getPageIndices());pages.forEach(p=>output.addPage(p));bar.style.width=`${Math.round((i+1)/files.length*100)}%`;progressText.textContent=`Read ${i+1} of ${files.length} PDFs`;}const bytes=await output.save();downloadBlob(new Blob([bytes],{type:'application/pdf'}),'merged.pdf');result(resultEl,`Done. Merged ${files.length} PDFs into a ${formatBytes(bytes.length)} file.`);}catch(err){result(resultEl,`Could not merge the PDFs: ${escapeHtml(err.message||'Unknown PDF error')}`,true);}finally{setBusy(run,false,'Merge and download','Merging…');}});update();
}

function bindPdfSplit() {
  const input=document.querySelector('#split-file'),zone=input.closest('.drop-zone'),info=document.querySelector('#split-info'),run=document.querySelector('#split-run'),clear=document.querySelector('#split-clear'),resultEl=document.querySelector('#split-result');let file=null;
  const showInfo=async next=>{file=next&&isPdf(next)?next:null;if(!file){info.textContent='No PDF selected.';return;}if(!window.PDFLib){info.textContent=`${file.name} · ${formatBytes(file.size)} · PDF engine unavailable`;return;}try{const doc=await PDFLib.PDFDocument.load(await file.arrayBuffer());const pages=doc.getPageCount();info.textContent=`${file.name} · ${formatBytes(file.size)} · ${pages} page${pages===1?'':'s'}`;}catch{info.textContent='This PDF could not be read.';}};
  bindDropZone(zone,input,incoming=>showInfo(incoming[0]),false); clear.addEventListener('click',()=>{file=null;input.value='';info.textContent='No PDF selected.';document.querySelector('#split-pages').value='';resultEl.innerHTML='';});
  run.addEventListener('click',async()=>{if(!file)return result(resultEl,'Choose a PDF first.',true);if(!window.PDFLib)return result(resultEl,'The PDF engine did not load. Check your connection and reload the page.',true);const raw=document.querySelector('#split-pages').value.trim();if(!raw)return result(resultEl,'Enter the pages you want, for example 1,3,5-7.',true);setBusy(run,true,'Create PDF and download','Creating…');try{const source=await PDFLib.PDFDocument.load(await file.arrayBuffer());const indices=parsePages(raw,source.getPageCount());if(!indices.length)throw new Error(`No valid pages were found. The document has ${source.getPageCount()} pages.`);const output=await PDFLib.PDFDocument.create();const pages=await output.copyPages(source,indices);pages.forEach(p=>output.addPage(p));const bytes=await output.save();downloadBlob(new Blob([bytes],{type:'application/pdf'}),`${file.name.replace(/\.[^.]+$/,'')}-split.pdf`);result(resultEl,`Done. Created a ${indices.length}-page PDF from pages ${indices.map(i=>i+1).join(', ')}.`);}catch(err){result(resultEl,`Could not split the PDF: ${escapeHtml(err.message||'Unknown PDF error')}`,true);}finally{setBusy(run,false,'Create PDF and download','Creating…');}});
}
function parsePages(raw,total){const set=new Set();for(const part of raw.split(',')){const v=part.trim();if(/^\d+$/.test(v)){const n=Number(v);if(n>=1&&n<=total)set.add(n-1);continue;}const m=v.match(/^(\d+)\s*-\s*(\d+)$/);if(!m)continue;let a=Number(m[1]),b=Number(m[2]);if(a>b)[a,b]=[b,a];for(let n=Math.max(1,a);n<=Math.min(total,b);n++)set.add(n-1);}return [...set].sort((a,b)=>a-b);}
