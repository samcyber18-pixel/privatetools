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
function validExt(file, exts){const n=file?.name?.toLowerCase()||'';return !!file && exts.some(e=>n.endsWith(e));}
function wireInput(input,zone,onFile){const pick=(files)=>{const f=files?.[0];onFile(f||null)};input.addEventListener('change',()=>pick(input.files));zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('dragover')});zone.addEventListener('dragleave',()=>zone.classList.remove('dragover'));zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('dragover');pick(e.dataTransfer.files)})}
const configs={
  'docx-to-pdf':{accept:'.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document',label:'Choose a Word document',hint:'DOCX · processed in your browser',idle:'Convert to PDF'},
  'pdf-to-docx':{accept:'.pdf,application/pdf',label:'Choose a PDF',hint:'PDF · text-focused conversion',idle:'Convert to Word'},
  'pptx-to-pdf':{accept:'.pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation',label:'Choose a PowerPoint file',hint:'PPTX · rendered locally to PDF',idle:'Convert to PDF'},
  'pdf-to-pptx':{accept:'.pdf,application/pdf',label:'Choose a PDF',hint:'PDF · each page becomes a PowerPoint slide',idle:'Convert to PowerPoint'}
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
  let file=null;
  wireInput(input,zone,f=>{
    file=f && validExt(f, kind.includes('docx')?['.docx']:kind.includes('pptx')?['.pptx']:['.pdf']) ? f:null;
    info.textContent=file?`${file.name} · ${fmt(file.size)}`:'No file selected.';
    $('#office-result').innerHTML='';resetProgress();
  });
  clear.addEventListener('click',()=>{file=null;input.value='';info.textContent='No file selected.';$('#office-result').innerHTML='';resetProgress();});
  run.addEventListener('click',async()=>{
    if(!file)return showResult(`Choose a ${kind.includes('pptx')?'PPTX':kind.includes('docx')?'DOCX':'PDF'} file first.`,true);
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
  const {jsPDF}=await jspdfLib();
  setProgress(8,'Reading Word document…');

  let stage=document.querySelector('#conversion-stage');
  if(!stage){
    stage=document.createElement('div');
    stage.id='conversion-stage';
    stage.className='docx-conversion-stage';
    stage.setAttribute('aria-hidden','true');
    document.body.appendChild(stage);
  }
  stage.replaceChildren();

  const styleHost=document.createElement('div');
  const bodyHost=document.createElement('div');
  styleHost.className='docx-style-host';
  bodyHost.className='docx-body-host';
  stage.append(styleHost, bodyHost);

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
  for(const node of nodes){
    (node.nodeName === 'STYLE' ? styleHost : bodyHost).appendChild(node);
  }

  await waitForFonts();
  await nextFrame();

  const sourcePages=[...bodyHost.querySelectorAll('.docx-wrapper>section.docx')];
  if(!sourcePages.length) throw new Error('No Word page container was produced.');

  // docx-preview honors explicit/page-break markers, but it does not perform
  // live re-pagination for ordinary flowing text. A normal multi-page Word
  // file can therefore render as one fixed-height section with the overflow
  // clipped. We paginate the rendered block-level content ourselves using the
  // actual Word page height/width produced by the library.
  const outputWrapper=document.createElement('div');
  outputWrapper.className='docx-wrapper docx-wrapper-output';
  bodyHost.appendChild(outputWrapper);

  const generatedPages=[];
  for(let si=0; si<sourcePages.length; si++){
    const source=sourcePages[si];
    const rect=source.getBoundingClientRect();
    const pageWidth=rect.width || source.offsetWidth;
    const pageHeight=rect.height || parseFloat(getComputedStyle(source).minHeight) || source.offsetHeight;
    if(!pageWidth || !pageHeight) throw new Error(`Word page ${si+1} has no measurable page geometry.`);

    const article=source.querySelector(':scope > article');
    const header=source.querySelector(':scope > header');
    const footer=source.querySelector(':scope > footer');

    if(!article){
      const clone=source.cloneNode(true);
      clone.style.width=`${pageWidth}px`;
      clone.style.height=`${pageHeight}px`;
      clone.style.minHeight=`${pageHeight}px`;
      outputWrapper.appendChild(clone);
      generatedPages.push(clone);
      continue;
    }

    const articleRect=article.getBoundingClientRect();
    const footerTop=footer ? footer.getBoundingClientRect().top : (rect.top+pageHeight);
    const usableHeight=Math.max(1, footerTop-articleRect.top);
    const children=[...article.children];

    // If this section already contains separate source page content, keep it;
    // otherwise split its flowing blocks across as many physical pages as are
    // necessary. We use block boundaries so text is not sliced mid-line.
    let groups=[[]];
    let groupStartTop=null;
    for(const child of children){
      const cr=child.getBoundingClientRect();
      const top=cr.top-articleRect.top;
      const bottom=cr.bottom-articleRect.top;
      if(groupStartTop===null) groupStartTop=top;
      const wouldOverflow = groups[groups.length-1].length>0 && (bottom-groupStartTop)>usableHeight+0.5;
      if(wouldOverflow){
        groups.push([]);
        groupStartTop=top;
      }
      groups[groups.length-1].push(child);
    }
    if(!groups.length || !groups[0].length) groups=[[]];

    for(let gi=0; gi<groups.length; gi++){
      const page=source.cloneNode(false);
      page.style.width=`${pageWidth}px`;
      page.style.height=`${pageHeight}px`;
      page.style.minHeight=`${pageHeight}px`;
      page.style.overflow='hidden';

      if(header) page.appendChild(header.cloneNode(true));
      const pageArticle=article.cloneNode(false);
      pageArticle.innerHTML='';
      pageArticle.style.minHeight='0';
      pageArticle.style.height='auto';
      for(const child of groups[gi]) pageArticle.appendChild(child.cloneNode(true));
      page.appendChild(pageArticle);
      if(footer) page.appendChild(footer.cloneNode(true));

      outputWrapper.appendChild(page);
      generatedPages.push(page);
    }
  }

  if(!generatedPages.length) throw new Error('No Word pages could be generated.');

  // Hide the original renderer output now that all measurements/clones are complete.
  // It remains in the DOM so its computed styles stay valid while generated pages render.
  for(const source of sourcePages) source.style.display='none';

  await nextFrame();
  await waitForFonts();
  await nextFrame();

  const measured=generatedPages.map((page,index)=>{
    const r=page.getBoundingClientRect();
    const width=r.width || page.offsetWidth;
    const height=r.height || page.offsetHeight;
    if(!width || !height) throw new Error(`Word page ${index+1} has no measurable layout.`);
    return {page,width,height};
  });

  setProgress(22,`Paginated ${measured.length} Word page${measured.length===1?'':'s'}; preparing PDF…`);

  const pxToMm=25.4/96;
  const first=measured[0];
  const pdf=new jsPDF({
    orientation:first.width>=first.height?'landscape':'portrait',
    unit:'mm',
    format:[first.width*pxToMm,first.height*pxToMm],
    compress:true
  });

  for(let i=0;i<measured.length;i++){
    const {page,width,height}=measured[i];
    const previous={position:page.style.position,left:page.style.left,top:page.style.top,margin:page.style.margin};
    page.style.position='relative';
    page.style.left='0';
    page.style.top='0';
    page.style.margin='0';
    await nextFrame();
    const canvas=await html2canvas(page,{
      backgroundColor:'#ffffff',
      scale:2,
      useCORS:true,
      allowTaint:false,
      logging:false,
      imageTimeout:20000,
      width:Math.ceil(width),
      height:Math.ceil(height),
      scrollX:0,
      scrollY:0,
      windowWidth:Math.max(window.innerWidth,Math.ceil(width)),
      windowHeight:Math.max(window.innerHeight,Math.ceil(height))
    });
    page.style.position=previous.position;
    page.style.left=previous.left;
    page.style.top=previous.top;
    page.style.margin=previous.margin;

    if(i){
      pdf.addPage([width*pxToMm,height*pxToMm],width>=height?'landscape':'portrait');
    }
    const jpeg=canvas.toDataURL('image/jpeg',0.95);
    pdf.addImage(jpeg,'JPEG',0,0,width*pxToMm,height*pxToMm,'FAST');
    setProgress(22+Math.round((i+1)/measured.length*72),`Added page ${i+1} of ${measured.length}`);
    canvas.width=1;
    canvas.height=1;
  }

  const blob=pdf.output('blob');
  downloadBlob(blob,`${file.name.replace(/\.docx$/i,'')}.pdf`);
  showResult(`Done. Created a ${measured.length}-page PDF (${fmt(blob.size)}). <strong>Pagination:</strong> content is split at rendered block boundaries to prevent ordinary multi-page documents from being clipped into a single page. Complex Word-specific layout can still differ from Microsoft Word.`);
  stage.replaceChildren();
}

function groupPdfText(items){
  const lines=[];
  for(const item of items){const text=(item.str||'').trim();if(!text)continue;const y=Math.round(item.transform?.[5]||0);const x=item.transform?.[4]||0;let line=lines.find(l=>Math.abs(l.y-y)<=3);if(!line){line={y,parts:[]};lines.push(line)}line.parts.push({x,text});}
  lines.sort((a,b)=>b.y-a.y);return lines.map(l=>l.parts.sort((a,b)=>a.x-b.x).map(p=>p.text).join(' ')).filter(Boolean);
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
