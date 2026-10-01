'use strict';

const $ = (s) => document.querySelector(s);
const esc = (v) => String(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = (b) => { if (!Number.isFinite(b)) return '—'; if (b < 1024) return `${b} B`; if (b < 1048576) return `${(b/1024).toFixed(1)} KB`; return `${(b/1048576).toFixed(1)} MB`; };
function safeName(name){return name.replace(/[\\/:*?"<>|\x00-\x1F]/g,'-').trim()||'file';}
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
  clear.addEventListener('click',()=>{file=null;input.value='';info.textContent='No file selected.';$('#office-result').innerHTML='';resetProgress()});
  run.addEventListener('click',async()=>{
    if(!file)return showResult(`Choose a ${kind.includes('pptx')?'PPTX':kind.includes('docx')?'DOCX':'PDF'} file first.`,true);
    setBusy(true,cfg.idle,'Converting…'); resetProgress(); setProgress(2,'Loading conversion engine…');
    try{
      if(kind==='docx-to-pdf') await docxToPdf(file);
      else if(kind==='pdf-to-docx') await pdfToDocx(file);
      else if(kind==='pptx-to-pdf') await pptxToPdf(file);
      else if(kind==='pdf-to-pptx') await pdfToPptx(file);
    }catch(err){showResult(`Conversion failed: ${esc(err?.message||'Unknown error')}`,true)}
    finally{setBusy(false,cfg.idle,'Converting…')}
  });
  run.textContent=cfg.idle;
  $('#office-label').textContent=cfg.label;
  $('#office-hint').textContent=cfg.hint;
}

async function docxToPdf(file){
  if(!window.mammoth || !window.html2canvas) throw new Error('The document conversion libraries did not load. Reload the page and try again.');
  const {jsPDF}=await jspdfLib(); setProgress(12,'Reading Word document…');
  const arrayBuffer=await file.arrayBuffer();
  const converted=await mammoth.convertToHtml({arrayBuffer});
  const stage=$('#conversion-stage'); stage.innerHTML=converted.value;
  stage.classList.add('docx-render-stage');
  await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
  setProgress(30,'Rendering document pages…');
  const canvas=await html2canvas(stage,{backgroundColor:'#ffffff',scale:1.5,useCORS:true,logging:false});
  const pdf=new jsPDF({orientation:'portrait',unit:'mm',format:'a4',compress:true});
  const pageW=210,pageH=297, pxPerPage=canvas.width*(pageH/pageW);
  const total=Math.max(1,Math.ceil(canvas.height/pxPerPage));
  for(let i=0;i<total;i++){
    if(i) pdf.addPage();
    const slice=document.createElement('canvas');slice.width=canvas.width;slice.height=Math.min(pxPerPage,canvas.height-i*pxPerPage);
    const ctx=slice.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,slice.width,slice.height);ctx.drawImage(canvas,0,i*pxPerPage,canvas.width,slice.height,0,0,slice.width,slice.height);
    const img=slice.toDataURL('image/jpeg',0.92);const h=(slice.height/slice.width)*pageW;pdf.addImage(img,'JPEG',0,0,pageW,h,'FAST');setProgress(30+Math.round((i+1)/total*65),`Rendered page ${i+1} of ${total}`);
  }
  const blob=pdf.output('blob');downloadBlob(blob,`${file.name.replace(/\.docx$/i,'')}.pdf`);showResult(`Done. Created a ${total}-page PDF (${fmt(blob.size)}).`);
}

function groupPdfText(items){
  const lines=[];
  for(const item of items){
    const text=(item.str||'').trim(); if(!text) continue;
    const y=Math.round(item.transform?.[5]||0); const x=item.transform?.[4]||0;
    let line=lines.find(l=>Math.abs(l.y-y)<=3);
    if(!line){line={y,parts:[]};lines.push(line)}
    line.parts.push({x,text});
  }
  lines.sort((a,b)=>b.y-a.y);
  return lines.map(l=>l.parts.sort((a,b)=>a.x-b.x).map(p=>p.text).join(' ')).filter(Boolean);
}
async function pdfToDocx(file){
  const pdfjs=await pdfJs();const {Document,Packer,Paragraph,TextRun,HeadingLevel}=await docxLib();
  setProgress(12,'Opening PDF…');const data=new Uint8Array(await file.arrayBuffer());const pdf=await pdfjs.getDocument({data}).promise;
  const paras=[];
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);const tc=await page.getTextContent();const lines=groupPdfText(tc.items);
    paras.push(new Paragraph({text:`Page ${p}`,heading:HeadingLevel.HEADING_2}));
    for(const line of lines) paras.push(new Paragraph({children:[new TextRun(line)]}));
    if(p<pdf.numPages) paras.push(new Paragraph({pageBreakBefore:true}));
    setProgress(15+Math.round(p/pdf.numPages*75),`Extracted page ${p} of ${pdf.numPages}`);
  }
  if(paras.length===0) paras.push(new Paragraph('No extractable text was found in this PDF.'));
  const doc=new Document({sections:[{children:paras}]});const blob=await Packer.toBlob(doc);downloadBlob(blob,`${file.name.replace(/\.pdf$/i,'')}.docx`);showResult(`Done. Created a Word document from ${pdf.numPages} PDF page${pdf.numPages===1?'':'s'} (${fmt(blob.size)}).`);
  showResult(`Done. Created a Word document from ${pdf.numPages} page${pdf.numPages===1?'':'s'} (${fmt(blob.size)}). <strong>Text-first conversion:</strong> complex page layout, forms, and some graphics may not be preserved.`);
}

async function pptxToPdf(file){
  if(!window.PptxViewJS || !window.html2canvas) throw new Error('The PowerPoint rendering libraries did not load. Reload the page and try again.');
  const {jsPDF}=await jspdfLib();
  setProgress(10,'Loading PowerPoint…');
  const canvas=$('#pptx-render-canvas');const viewer=new window.PptxViewJS.PPTXViewer({canvas});
  await viewer.loadFile(file);const count=viewer.getSlideCount();if(!count)throw new Error('No slides were found in the PowerPoint file.');
  canvas.width=1280;canvas.height=720;
  const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:[254,142.875],compress:true});
  for(let i=0;i<count;i++){
    await viewer.goToSlide(i);await viewer.render();await new Promise(r=>requestAnimationFrame(r));
    if(i) pdf.addPage([254,142.875],'landscape');
    const data=canvas.toDataURL('image/jpeg',0.92);pdf.addImage(data,'JPEG',0,0,254,142.875,'FAST');setProgress(15+Math.round((i+1)/count*80),`Rendered slide ${i+1} of ${count}`);
  }
  const blob=pdf.output('blob');downloadBlob(blob,`${file.name.replace(/\.pptx$/i,'')}.pdf`);showResult(`Done. Created a ${count}-slide PDF (${fmt(blob.size)}). <strong>Note:</strong> the PDF is rendered from slides, so the output is visually oriented rather than editable text.`);
}

async function pdfToPptx(file){
  const pdfjs=await pdfJs();
  if(!window.PptxGenJS) throw new Error('The PowerPoint generation library did not load. Reload the page and try again.');
  setProgress(10,'Opening PDF…');const pdf=await pdfjs.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;const pptx=new PptxGenJS();pptx.layout='LAYOUT_WIDE';pptx.author='PrivateTools';pptx.subject='PDF to PowerPoint conversion';pptx.title=file.name;
  const canvas=$('#pdf-render-canvas');const ctx=canvas.getContext('2d',{alpha:false});
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);const base=page.getViewport({scale:1});const scale=Math.min(2.2,1600/base.width);const viewport=page.getViewport({scale});canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height);await page.render({canvasContext:ctx,viewport}).promise;
    const slide=pptx.addSlide();slide.background={color:'FFFFFF'};slide.addImage({data:canvas.toDataURL('image/jpeg',0.92),x:0,y:0,w:13.333,h:7.5});setProgress(15+Math.round(p/pdf.numPages*80),`Rendered page ${p} of ${pdf.numPages}`);
  }
  const data=await pptx.write({outputType:'blob'});const blob=data instanceof Blob?data:new Blob([data],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});downloadBlob(blob,`${file.name.replace(/\.pdf$/i,'')}.pptx`);showResult(`Done. Created a ${pdf.numPages}-slide PowerPoint (${fmt(blob.size)}). <strong>Note:</strong> each PDF page is placed on a slide as an image, so the slide content is not independently editable.`);
}

window.addEventListener('DOMContentLoaded',init);
