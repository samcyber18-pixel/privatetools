'use strict';
const $ = (s) => document.querySelector(s);
const esc = (v) => String(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = (b) => { if (!Number.isFinite(b)) return '—'; if (b < 1024) return `${b} B`; if (b < 1048576) return `${(b/1024).toFixed(1)} KB`; return `${(b/1048576).toFixed(1)} MB`; };
function safeName(name){return name.replace(/[\/:*?"<>|\x00-\x1F]/g,'-').trim()||'file';}
function downloadBlob(blob,name){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=safeName(name);document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function showResult(msg,error=false){const el=$('#office-result');el.innerHTML=`<div class="result${error?' error':''}">${msg}</div>`;}
function setBusy(busy,idle='Convert and download',active='Converting…'){const b=$('#office-run');b.disabled=busy;b.textContent=busy?active:idle;}
function setProgress(value,text){const wrap=$('#office-progress');const bar=wrap.querySelector('.progress>div');wrap.classList.remove('hidden');bar.style.width=`${Math.max(0,Math.min(100,value))}%`;$('#office-progress-text').textContent=text||'';}
function resetProgress(){const w=$('#office-progress');w.classList.add('hidden');w.querySelector('.progress>div').style.width='0%';$('#office-progress-text').textContent='';}
function validExt(file, exts){
  if(!file) return false;
  const n=(file.name||'').toLowerCase();
  const type=(file.type||'').toLowerCase();
  return exts.some(e=>n.endsWith(e)) || (exts.includes('.pdf') && type==='application/pdf') || (exts.includes('.docx') && type==='application/vnd.openxmlformats-officedocument.wordprocessingml.document') || (exts.includes('.pptx') && type==='application/vnd.openxmlformats-officedocument.presentationml.presentation');
}
function wireInput(input,zone,onFile){
  if(!input || !zone) return;
  const pick=(files)=>{
    const f=files && files.length ? files[0] : null;
    onFile(f||null);
  };
  const syncFromInput=()=>pick(input.files);
  input.addEventListener('change',syncFromInput);
  input.addEventListener('input',syncFromInput);
  zone.addEventListener('click',e=>{
    if(e.target.closest('label') || e.target.closest('input')) return;
    input.click();
  });
  zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('dragover')});
  zone.addEventListener('dragleave',e=>{if(!zone.contains(e.relatedTarget)) zone.classList.remove('dragover')});
  zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('dragover');pick(e.dataTransfer?.files)});
}
const configs={
  'docx-to-pdf':{accept:'.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document',sourceExts:['.docx'],sourceLabel:'DOCX',label:'Choose a Word document',hint:'DOCX · processed in your browser',idle:'Convert to PDF'},
  'pdf-to-docx':{accept:'.pdf,application/pdf',sourceExts:['.pdf'],sourceLabel:'PDF',label:'Choose a PDF',hint:'PDF · text-focused conversion',idle:'Convert to Word'},
  'pptx-to-pdf':{accept:'.pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation',sourceExts:['.pptx'],sourceLabel:'PPTX',label:'Choose a PowerPoint file',hint:'PPTX · rendered locally to PDF',idle:'Convert to PDF'},
  'pdf-to-pptx':{accept:'.pdf,application/pdf',sourceExts:['.pdf'],sourceLabel:'PDF',label:'Choose a PDF',hint:'PDF · each page becomes a PowerPoint slide',idle:'Convert to PowerPoint'}
};
async function pdfJs(){
  if(window.__pdfjs) return window.__pdfjs;
  const pdfjs=await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/6.3.289/pdf.min.mjs');
  pdfjs.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/6.3.289/pdf.worker.min.mjs';
  window.__pdfjs=pdfjs; return pdfjs;
}
async function docxLib(){
  if(window.__docxLib) return window.__docxLib;
  window.__docxLib=await import('https://esm.sh/docx@9.8.1');
  return window.__docxLib;
}
async function jspdfLib(){
  if(window.__jspdfLib) return window.__jspdfLib;
  window.__jspdfLib=await import('https://esm.sh/jspdf@4.2.1');
  return window.__jspdfLib;
}
function init(){
  const kind=document.body.dataset.officeTool;
  const cfg=configs[kind]; if(!cfg) return;
  const input=$('#office-file'),zone=$('#office-drop'),info=$('#office-info'),run=$('#office-run'),clear=$('#office-clear');
  if(!input || !zone || !info || !run || !clear) return;
  const allowedExts = cfg.sourceExts;
  let file=null;
  const setSelectedFile=(candidate)=>{
    const accepted = candidate && validExt(candidate, allowedExts);
    file = accepted ? candidate : null;
    if(file){
      info.textContent=`Selected: ${file.name} · ${fmt(file.size)}`;
      run.disabled=false;
      run.removeAttribute('aria-disabled');
    }else{
      info.textContent=candidate ? `Unsupported file type. Please choose ${allowedExts.join(' or ')}.` : 'No file selected.';
      run.disabled=false;
      run.removeAttribute('aria-disabled');
    }
    const result=$('#office-result');
    if(result) result.innerHTML='';
    resetProgress();
  };
  wireInput(input,zone,setSelectedFile);
  // Some browsers populate FileList only after the picker closes; reading it
  // again on focus makes the UI resilient to that behavior.
  input.addEventListener('focus',()=>setTimeout(()=>{ if(input.files?.length) setSelectedFile(input.files[0]); },0));
  clear.addEventListener('click',()=>{file=null;input.value='';info.textContent='No file selected.';const result=$('#office-result');if(result) result.innerHTML='';resetProgress();});
  run.addEventListener('click',async()=>{
    if(!file)return showResult(`Choose a ${cfg.sourceLabel} file first.`,true);
    setBusy(true,cfg.idle,'Converting…'); resetProgress(); setProgress(2,'Loading conversion engine…');
    try{
      if(kind==='docx-to-pdf') await docxToPdf(file);
      else if(kind==='pdf-to-docx') await pdfToDocx(file);
      else if(kind==='pptx-to-pdf') await pptxToPdf(file);
      else if(kind==='pdf-to-pptx') await pdfToPptx(file);
    }catch(err){showResult(`Conversion failed: ${esc(err?.message||'Unknown error')}`,true);}
    finally{setBusy(false,cfg.idle,'Converting…');}
  });
  run.textContent=cfg.idle;
  $('#office-label').textContent=cfg.label;
  $('#office-hint').textContent=cfg.hint;
}
async function nextFrame(){await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));}
async function waitForFonts(){if(document.fonts?.ready) await document.fonts.ready;}

async function docxToPdf(file){
  if(!window.docx || typeof window.docx.parseAsync !== 'function' || typeof window.docx.renderDocument !== 'function'){
    throw new Error('The Word rendering library did not load correctly. Reload the page and try again.');
  }
  if(!window.Paged || typeof window.Paged.Previewer !== 'function'){
    throw new Error('The pagination engine did not load. Reload the page and try again.');
  }
  if(!window.html2canvas) throw new Error('The PDF rendering library did not load. Reload the page and try again.');

  const {jsPDF}=await jspdfLib();
  setProgress(6,'Reading Word document…');

  let stage=document.querySelector('#conversion-stage');
  if(!stage){
    stage=document.createElement('div');
    stage.id='conversion-stage';
    document.body.appendChild(stage);
  }
  stage.replaceChildren();
  stage.className='docx-conversion-stage';

  const styleHost=document.createElement('div');
  const renderHost=document.createElement('div');
  styleHost.className='docx-style-host';
  renderHost.className='docx-body-host';
  stage.append(styleHost,renderHost);

  const options={
    inWrapper:true,
    breakPages:true,
    ignoreWidth:false,
    ignoreHeight:false,
    ignoreFonts:false,
    ignoreLastRenderedPageBreak:false,
    renderHeaders:true,
    renderFooters:true,
    renderFootnotes:true,
    renderEndnotes:true,
    useBase64URL:true,
    renderChanges:false,
    renderComments:false,
    renderAltChunks:true,
    experimental:false,
    debug:false
  };

  const doc=await window.docx.parseAsync(await file.arrayBuffer(), options);
  const nodes=await window.docx.renderDocument(doc, options);
  if(!Array.isArray(nodes) || nodes.length===0) throw new Error('The Word document could not be rendered.');

  const styleNodes=nodes.filter(n=>n && n.nodeName==='STYLE');
  const bodyNodes=nodes.filter(n=>n && n.nodeName!=='STYLE');
  for(const node of styleNodes) styleHost.appendChild(node);
  for(const node of bodyNodes) renderHost.appendChild(node);

  if(document.fonts?.ready) await document.fonts.ready;
  await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));

  const sourcePages=[...renderHost.querySelectorAll('.docx-wrapper>section.docx')];
  if(!sourcePages.length) throw new Error('The Word document did not produce a usable page layout.');

  // Use the first source page to recover the document's page geometry. The
  // pagination engine will then perform normal overflow fragmentation inside
  // that page-sized frame instead of relying on source-declared breaks alone.
  const source=sourcePages[0];
  const sourceRect=source.getBoundingClientRect();
  const pageWidth=Math.round(sourceRect.width || source.offsetWidth);
  const pageHeight=Math.round(sourceRect.height || source.offsetHeight);
  if(!pageWidth || !pageHeight) throw new Error('Word page dimensions could not be measured.');

  const article=source.querySelector(':scope > article') || source;
  const content=document.createElement('div');
  content.className='docx-pagination-content';
  content.innerHTML=article.innerHTML;

  // Preserve the source Word page's useful visual geometry while deliberately
  // removing its fixed height. Paged.js then determines where flowing blocks
  // must break when a page is full.
  const articleStyle=getComputedStyle(article);
  const sectionStyle=getComputedStyle(source);
  content.style.boxSizing='border-box';
  content.style.width=`${pageWidth}px`;
  content.style.margin='0';
  content.style.padding=articleStyle.padding;
  content.style.fontFamily=articleStyle.fontFamily;
  content.style.fontSize=articleStyle.fontSize;
  content.style.lineHeight=articleStyle.lineHeight;
  content.style.color=articleStyle.color;
  content.style.background=sectionStyle.background;

  const pagedRoot=document.createElement('div');
  pagedRoot.id='paged-root';
  pagedRoot.className='paged-output';
  stage.appendChild(pagedRoot);
  const pageMmW=pageWidth*25.4/96;
  const pageMmH=pageHeight*25.4/96;

  const paginationCss=`
    @page { size: ${pageMmW}mm ${pageMmH}mm; margin: 0; }
    .paged-output { width: ${pageWidth}px; }
    .docx-pagination-content { width: ${pageWidth}px; box-sizing: border-box; overflow: visible; }
    .docx-pagination-content, .docx-pagination-content * { max-width: none; }
    .pagedjs_pages { margin: 0 !important; padding: 0 !important; }
    .pagedjs_page { margin: 0 !important; }
    .pagedjs_sheet { margin: 0 !important; overflow: hidden !important; }
    .pagedjs_area { overflow: visible !important; }
  `;

  // Paged.js accepts the rendered DOCX CSS as inline stylesheet objects.
  // Keeping it in the same preview document preserves fonts, paragraph styles,
  // tables, images, headers and footers as far as the DOCX renderer exposes them.
  const docxStyles=styleNodes
    .map(el=>({[window.location.href]:el.textContent || ''}))
    .filter(x=>Object.values(x)[0]);
  docxStyles.push({[window.location.href]:paginationCss});

  stage.style.display='block';
  stage.style.position='absolute';
  stage.style.left='-20000px';
  stage.style.top='0';
  stage.style.width=`${pageWidth}px`;
  stage.style.background='#fff';
  stage.style.zIndex='-1';

  setProgress(18,'Paginating document…');
  const previewer=new window.Paged.Previewer();
  const flow=await previewer.preview(content, docxStyles, pagedRoot);
  if(!flow || !flow.total) throw new Error('The pagination engine did not produce any pages.');

  if(document.fonts?.ready) await document.fonts.ready;
  await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));

  const pages=[...pagedRoot.querySelectorAll('.pagedjs_page')];
  if(!pages.length) throw new Error('No paginated Word pages were produced.');

  setProgress(28,`Preparing ${pages.length} PDF page${pages.length===1?'':'s'}…`);
  const pdf=new jsPDF({
    orientation:pageWidth>=pageHeight?'landscape':'portrait',
    unit:'mm',
    format:[pageMmW,pageMmH],
    compress:true
  });

  for(let i=0;i<pages.length;i++){
    const page=pages[i];
    const rect=page.getBoundingClientRect();
    const captureW=Math.max(1,Math.ceil(rect.width));
    const captureH=Math.max(1,Math.ceil(rect.height));
    const canvas=await html2canvas(page,{
      backgroundColor:'#ffffff',
      scale:1.5,
      useCORS:true,
      logging:false,
      width:captureW,
      height:captureH,
      windowWidth:captureW,
      windowHeight:captureH
    });
    if(i) pdf.addPage([pageMmW,pageMmH],pageWidth>=pageHeight?'landscape':'portrait');
    const img=canvas.toDataURL('image/jpeg',0.94);
    pdf.addImage(img,'JPEG',0,0,pageMmW,pageMmH,'FAST');
    setProgress(28+Math.round((i+1)/pages.length*67),`Rendered page ${i+1} of ${pages.length}`);
  }

  const blob=pdf.output('blob');
  downloadBlob(blob,`${file.name.replace(/\.docx$/i,'')}.pdf`);
  showResult(`Done. Created a ${pages.length}-page PDF (${fmt(blob.size)}). <strong>Note:</strong> browser rendering can still differ from Microsoft Word for advanced features, but normal flowing document content is paginated by the browser pagination engine.`);
  try{previewer.dispose?.()}catch{}
}

async function pdfToDocx(file){
  const pdfjs=await pdfJs();const {Document,Packer,Paragraph,TextRun,HeadingLevel}=await docxLib();
  setProgress(12,'Opening PDF…');const data=new Uint8Array(await file.arrayBuffer());const pdf=await pdfjs.getDocument({data}).promise;const paras=[];
  for(let p=1;p<=pdf.numPages;p++){const page=await pdf.getPage(p);const tc=await page.getTextContent();const lines=groupPdfText(tc.items);paras.push(new Paragraph({text:`Page ${p}`,heading:HeadingLevel.HEADING_2}));for(const line of lines)paras.push(new Paragraph({children:[new TextRun(line)]}));if(p<pdf.numPages)paras.push(new Paragraph({pageBreakBefore:true}));setProgress(15+Math.round(p/pdf.numPages*75),`Extracted page ${p} of ${pdf.numPages}`);}
  if(paras.length===0)paras.push(new Paragraph('No extractable text was found in this PDF.'));
  const doc=new Document({sections:[{children:paras}]});const blob=await Packer.toBlob(doc);downloadBlob(blob,`${file.name.replace(/\.pdf$/i,'')}.docx`);showResult(`Done. Created a Word document from ${pdf.numPages} page${pdf.numPages===1?'':'s'} (${fmt(blob.size)}). <strong>Text-first conversion:</strong> complex page layout, forms, and some graphics may not be preserved.`);
}
async function pptxToPdf(file){
  if(!window.PptxViewJS||!window.html2canvas)throw new Error('The PowerPoint rendering libraries did not load. Reload the page and try again.');
  const {jsPDF}=await jspdfLib();setProgress(10,'Loading PowerPoint…');const canvas=$('#pptx-render-canvas');const viewer=new window.PptxViewJS.PPTXViewer({canvas});await viewer.loadFile(file);const count=viewer.getSlideCount();if(!count)throw new Error('No slides were found in the PowerPoint file.');canvas.width=1280;canvas.height=720;const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:[254,142.875],compress:true});for(let i=0;i<count;i++){await viewer.goToSlide(i);await viewer.render();await nextFrame();if(i)pdf.addPage([254,142.875],'landscape');const data=canvas.toDataURL('image/jpeg',0.95);pdf.addImage(data,'JPEG',0,0,254,142.875,'FAST');setProgress(15+Math.round((i+1)/count*80),`Rendered slide ${i+1} of ${count}`);}const blob=pdf.output('blob');downloadBlob(blob,`${file.name.replace(/\.pptx$/i,'')}.pdf`);showResult(`Done. Created a ${count}-slide PDF (${fmt(blob.size)}). <strong>Note:</strong> the PDF is rendered from slides, so the output is visually oriented rather than editable text.`);}
async function pdfToPptx(file){
  const pdfjs=await pdfJs();if(!window.PptxGenJS)throw new Error('The PowerPoint generation library did not load. Reload the page and try again.');setProgress(10,'Opening PDF…');const pdf=await pdfjs.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;const pptx=new PptxGenJS();pptx.layout='LAYOUT_WIDE';pptx.author='PrivateTools';pptx.subject='PDF to PowerPoint conversion';pptx.title=file.name;const canvas=$('#pdf-render-canvas');const ctx=canvas.getContext('2d',{alpha:false});for(let p=1;p<=pdf.numPages;p++){const page=await pdf.getPage(p);const base=page.getViewport({scale:1});const scale=Math.min(2.2,1600/base.width);const viewport=page.getViewport({scale});canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height);await page.render({canvasContext:ctx,viewport}).promise;const slide=pptx.addSlide();slide.background={color:'FFFFFF'};slide.addImage({data:canvas.toDataURL('image/jpeg',0.95),x:0,y:0,w:13.333,h:7.5});setProgress(15+Math.round(p/pdf.numPages*80),`Rendered page ${p} of ${pdf.numPages}`);}const data=await pptx.write({outputType:'blob'});const blob=data instanceof Blob?data:new Blob([data],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});downloadBlob(blob,`${file.name.replace(/\.pdf$/i,'')}.pptx`);showResult(`Done. Created a ${pdf.numPages}-slide PowerPoint (${fmt(blob.size)}). <strong>Note:</strong> each PDF page is placed on a slide as an image, so the slide content is not independently editable.`);}
window.addEventListener('DOMContentLoaded',init);
