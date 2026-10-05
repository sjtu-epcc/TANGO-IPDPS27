import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile, FileBlob } from '@oai/artifact-tool';

const workspaceDir = '/Users/yixiao/Desktop/TANGO-IPDPS27';
const buildDir = path.join(workspaceDir, '.codex-slide-build/image-to-ppt');
const SKILL_DIR = '/Users/yixiao/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations';
const FINAL_PPTX = path.join(workspaceDir, 'output/ppt/调度示意图_可编辑版.pptx');
const RUNTIME_PYTHON = '/Users/yixiao/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3';
const { finalizePresentation, resolvePresentationFont } = await import(pathToFileURL(path.join(SKILL_DIR,'container_tools/artifact_tool_utils.mjs')).href);
const FONT = resolvePresentationFont({fontFamily:'Times New Roman', availableFonts:['Times New Roman']});
const W = 2048, H = 1144;
const p = Presentation.create({slideSize:{width:W,height:H}});
const s = p.slides.add();
s.background.fill = '#FFFFFF';
const C = { ink:'#101010',gray:'#F5F6F8',row:'#F7F9F9',blue:'#EAF3FF',peach:'#FFF5EE',D:'#D9F3EE',P:'#FDE8D8',R0:'#E2E1F8',R1:'#DDEFFD',green:'#149C81',red:'#F21D24' };
let index=0;
function shape(geometry,x,y,w,h,fill='none',stroke='none',sw=0,r=0,name='') {
  return s.shapes.add({geometry,name:name||`${geometry}-${++index}`,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:sw,style:'solid'},...(r?{borderRadius:r}:{})});
}
function box(x,y,w,h,fill='#FFFFFF',sw=2.6,r=7,stroke=C.ink,name='') {return shape('rect',x,y,w,h,fill,stroke,sw,r,name);}
function text(str,x,y,w,h,size=28,bold=false,align='center',color=C.ink,italic=false,name='') {
  const q=shape('textbox',x,y,w,h,'none','none',0,0,name||str.replace(/\s+/g,' ').slice(0,60));
  q.text=str;
  q.text.style={typeface:FONT,fontSize:size,bold,italic,color,alignment:align,verticalAlignment:'middle',autoFit:'none',wrap:'none',insets:{left:0,right:0,top:0,bottom:0}};
  return q;
}
function line(x1,y1,x2,y2,color=C.ink,width=2.5,dashed=false,name='') {
  const left=Math.min(x1,x2),top=Math.min(y1,y2),w=Math.max(1,Math.abs(x2-x1)),h=Math.max(1,Math.abs(y2-y1));
  return s.shapes.add({geometry:'custom',name:name||`line-${++index}`,position:{left,top,width:w,height:h},fill:'none',line:{fill:color,width,style:dashed?'dashed':'solid'},customPaths:[{width:w,height:h,commands:[{moveTo:{x:x1-left,y:y1-top}},{lineTo:{x:x2-left,y:y2-top}}]}]});
}
function polygon(pts,fill=C.ink,name='') {
  const xs=pts.map(v=>v[0]),ys=pts.map(v=>v[1]);const left=Math.min(...xs),top=Math.min(...ys),width=Math.max(...xs)-left,height=Math.max(...ys)-top;
  return s.shapes.add({geometry:'custom',name:name||`polygon-${++index}`,position:{left,top,width,height},fill,line:{fill:'none',width:0},customPaths:[{width,height,commands:pts.map((v,i)=>i?{lineTo:{x:v[0]-left,y:v[1]-top}}:{moveTo:{x:v[0]-left,y:v[1]-top}}).concat({close:{}})}]});
}
function arrow(x1,y1,x2,y2,width=3.2,head=17,color=C.ink,name='') {
  const a=Math.atan2(y2-y1,x2-x1),len=Math.hypot(x2-x1,y2-y1),tip=Math.min(head,len*.45),bx=x2-tip*Math.cos(a),by=y2-tip*Math.sin(a),half=tip*.56;
  line(x1,y1,bx,by,color,width,false,`${name} stem`);
  polygon([[x2,y2],[bx-half*Math.sin(a),by+half*Math.cos(a)],[bx+half*Math.sin(a),by-half*Math.cos(a)]],color,`${name} arrowhead`);
}
function badge(n,x,y,d=56) {shape('ellipse',x,y,d,d,'#000000','none',0,0,`Step ${n}`);text(String(n),x,y-1,d,d,44,true,'center','#FFFFFF');}
function token(label,x,y,w,h,fill=C.P,sz=29,r=6,stroke=C.ink) {box(x,y,w,h,fill,2.2,r,stroke,label+' editable block');text(label,x+3,y+1,w-6,h-2,sz,true);}
function panel(x,y,w,h,headH,fill,name) {box(x,y,w,h,'#FFFFFF',3,8,C.ink,name);box(x,y,w,headH,fill,2.4,8);}
function cross(x,y,size=26) {line(x,y,x+size,y+size,C.red,6);line(x,y+size,x+size,y,C.red,6);}
function check(x,y,size=28) {line(x,y+size*.55,x+size*.28,y+size,C.green,5);line(x+size*.28,y+size,x+size,y,C.green,5);}

// Repeat loop and flow links.
line(1992,432,2026,432,C.ink,3.2);
line(2026,432,2026,61,C.ink,3.2);
line(2026,61,663,61,C.ink,3.2);
arrow(663,61,663,102,3.2,20,C.ink,'Next step');
text('Next step',1262,9,235,46,35,true);
arrow(423,282,468,282,3.2,18,C.ink,'Requests to temporal scheduling');
arrow(1008,282,1092,282,3.2,20,C.ink,'Selected requests');
text('Selected',1011,297,79,29,23);
text('requests',1011,326,79,29,23);
arrow(1629,282,1699,282,3.2,19,C.ink,'Mixed batch');
text('Mixed',1632,298,66,29,25);
text('batch',1632,328,66,29,25);
arrow(1847,353,1847,396,3.2,19,C.ink,'Execute to update');

// Request queue.
panel(38,165,385,249,66,C.gray,'Requests');
text('Requests',72,177,313,44,39,true);
text('Running',53,247,183,36,30,true,'left');
text('Ongoing decode',53,279,226,33,27,false,'left');
token('D',293,250,80,54,C.D,30,6,'#417D74');
line(46,322,416,322,'#7C7C7C',1.8,true);
text('Waiting',53,341,164,33,30,true,'left');
text('New / resumed',53,372,210,32,27,false,'left');
token('R1',226,337,82,57,C.P);
token('R2',326,337,82,57,C.P);

// Temporal scheduling.
panel(468,102,540,414,68,C.gray,'Temporal scheduling');
badge(1,490,109,56);
text('Temporal scheduling',573,109,420,56,43,true);
text('Select requests for this step',508,174,461,39,31);
box(492,220,494,55,C.row,2,7,'#777777','Running first row');
text('Running first',506,229,178,37,29,true,'left');
token('D',687,227,101,42,C.D,29,5,'#417D74');
box(492,291,494,59,C.peach,2,7,'#847871','Waiting FCFS row');
text('Waiting: FCFS',505,303,179,34,26,true,'left');
token('R1',689,298,105,45,C.P,29);
arrow(808,322,856,322,2.8,16);
token('R2',868,298,107,45,C.P,29);
arrow(735,352,735,404,3.1,18);
text('Continue / admit',757,355,225,39,30,false,'left');
box(492,405,494,71,C.row,2.4,7,C.ink,'Selected this step row');
text('Selected this step',501,422,183,39,24.5,true,'left');
token('D',689,415,88,51,C.D,29);
token('R1',788,415,89,51,C.P,29);
token('R2',889,415,87,51,C.P,29);
text('Token / KV limits; preemption',515,478,447,34,28);

// Spatial scheduling.
panel(1092,102,537,414,68,C.peach,'Spatial scheduling');
badge(2,1113,109,56);
text('Spatial scheduling',1214,111,395,54,43,true);
box(1114,192,148,92,C.D,2.5,8);
box(1285,192,151,92,C.peach,2.5,8);
box(1456,192,151,92,C.peach,2.5,8);
text('D',1132,204,112,38,31,true);
text('decode token',1118,243,140,31,27);
text('R1',1301,204,119,38,31,true);
text('prefill chunk',1289,243,143,31,27);
text('R2',1472,204,119,38,31,true);
text('prefill chunk',1460,243,143,31,27);
arrow(1353,294,1353,345,3,18);
text('Allocate & pack',1382,300,226,40,31,false,'left');
text('Budget',1292,344,94,42,31);
text('B',1392,344,27,42,31,false,'center',C.ink,true);
line(1131,395,1590,395,C.ink,2.5);
line(1131,384,1131,406,C.ink,2.5);
line(1590,384,1590,406,C.ink,2.5);
box(1131,413,94,38,C.D,2.1,0);
box(1225,413,196,38,C.P,2.1,0);
box(1421,413,169,38,C.P,2.1,0);
text('D',1141,457,72,37,29,true,'center',C.ink,true);
text('R1',1268,457,113,37,29,true,'center',C.ink,true);
text('R2',1450,457,113,37,29,true,'center',C.ink,true);

// GPU batch execution and state update.
panel(1699,148,293,205,66,C.gray,'Execute batch');
badge(3,1713,156,52);
text('Execute batch',1773,158,212,49,34,true);
box(1807,246,80,76,'#FFFFFF',3,7,C.ink,'GPU chip');
text('GPU',1808,259,78,47,34);
for(let i=0;i<5;i++){const a=1820+i*13;line(a,226,a,246,C.ink,3);line(a,322,a,341,C.ink,3);}
for(let i=0;i<5;i++){const a=257+i*13;line(1788,a,1807,a,C.ink,3);line(1887,a,1905,a,C.ink,3);}
box(1699,396,293,72,C.gray,3,8);
badge(4,1713,405,52);
text('Update states',1781,406,199,51,36,true);

// Dotted callout links.
line(468,513,34,613,'#333333',2.2,true);
line(1000,516,1034,614,'#333333',2.2,true);
line(1095,516,1049,614,'#333333',2.2,true);
line(1629,516,2023,614,'#333333',2.2,true);

// Problem 1: urgency-aware ordering.
panel(31,614,1003,482,62,C.blue,'Problem 1: FCFS ignores urgency');
text('Problem 1: FCFS ignores urgency',104,624,858,44,38,true);
text('R0: earlier long request; R1: loose TTFT; R2: tight TTFT',146,681,785,34,28);
text('All waiting at t0; D continues decoding',226,716,625,34,28);
text('R2 TTFT ddl',365,755,188,38,27,true);
text('R1 TTFT ddl',601,755,189,38,27,true);
text('FCFS',55,820,110,47,32,true,'left');
text('better',55,904,130,38,31,true,'left');
text('schedule',55,940,130,38,31,true,'left');
const segs=[196,196,212,195];
let tx=186;
['R0','R0','R1','R2'].forEach((v,i)=>{box(tx,818,segs[i],53,[C.R0,C.R0,C.R1,C.P][i],2.2,0);text(v,tx+18,822,segs[i]-36,46,29);tx+=segs[i];});
tx=186;
['R2','R1','R0','R0'].forEach((v,i)=>{box(tx,904,segs[i],51,[C.P,C.R1,C.R0,C.R0][i],2.2,0);text(v,tx+14,908,segs[i]-40,44,29);tx+=segs[i];});
line(459,789,459,976,C.red,2.8,true,'R2 deadline');
line(696,789,696,976,C.red,2.8,true,'R1 deadline');
cross(758,834,24);cross(973,834,24);
check(341,914,28);check(540,914,28);
arrow(186,984,795,984,2.8,17);
line(186,973,186,995,C.ink,2.2);
text('t',168,990,27,45,35,true,'center',C.ink,true);
text('0',188,1009,19,27,24,true);
text('Execution steps',811,967,190,39,28,false,'left');
text('Better ordering meets both TTFT deadlines',119,1035,841,44,31,true);

// Problem 2: token and P/D composition versus step time.
panel(1049,614,980,482,62,C.peach,'Problem 2: Fixed budget ignores step time');
text('Problem 2: Fixed budget ignores step time',1093,624,893,44,38,true);
text('Token cap: 2048 (illustrative)',1269,683,537,38,30);
text('Batch composition',1246,726,350,42,31,true);
text('Step time',1695,713,231,37,28,true);
text('TPOT SLO',1783,743,187,33,28,true);
text('1: Fill cap',1066,788,177,39,29,true,'left');
text('2: Fewer tokens',1066,856,177,39,25,true,'left');
text('3: Change P/D',1066,922,177,39,28,true,'left');
const bx=1245,bw=351;
box(bx,788,144,40,C.D,2.2,0);box(bx+144,788,bw-144,40,C.P,2.2,0);
text('D',bx+8,789,127,38,29,true);text('P',bx+156,789,184,38,29,true);
box(bx,855,124,41,C.D,2.2,0);box(bx+124,855,110,41,C.P,2.2,0);box(bx+234,855,bw-234,41,'#FFFFFF',2.2,0);
text('D',bx+8,856,108,39,29,true);text('P',bx+128,856,102,39,29,true);text('unused',bx+241,856,102,39,26);
box(bx,921,214,41,C.D,2.2,0);box(bx+214,921,bw-214,41,C.P,2.2,0);
text('D',bx+8,922,198,39,29,true);text('P',bx+220,922,125,39,29,true);
line(1628,778,1628,972,C.ink,2);
box(1628,788,322,40,C.R0,2.2,0);
box(1628,855,161,41,C.R0,2.2,0);
box(1628,921,185,41,C.R0,2.2,0);
line(1875,777,1875,980,C.red,2.8,true,'TPOT SLO deadline');
cross(1962,795,24);check(1805,858,27);check(1827,926,27);
box(1303,993,38,33,C.D,2.2,0);text('Decode',1352,990,110,42,26,false,'left');
box(1471,993,38,33,C.P,2.2,0);text('Prefill',1519,990,110,42,26,false,'left');
box(1661,993,38,33,C.R0,2.2,0);text('Illustrative; workload-dependent',1717,990,293,42,22,false,'left');
text('Token count and P/D mix jointly affect TPOT',1120,1035,849,44,31,true);

s.speakerNotes.textFrame.setText('Source: user-supplied scheduling diagram image. Original English labels and illustrative values reproduced. Diagram shapes, text, timelines, and bars are native editable PowerPoint objects.');
const candidatePath=path.join(buildDir,'candidate.pptx');
await (await PresentationFile.exportPptx(p)).save(candidatePath);
const png=await p.export({slide:s,format:'png',scale:1});
await fs.writeFile(path.join(buildDir,'draft.png'),new Uint8Array(await png.arrayBuffer()));
const layout=await s.export({format:'layout'});await fs.writeFile(path.join(buildDir,'layout.json'),await layout.text());
const result=await finalizePresentation({workspaceDir,candidatePath,finalPath:FINAL_PPTX,pythonExecutable:RUNTIME_PYTHON,integrityValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-bullet-geometry','--validate-heading-fit'],explicitTotalSlideCount:1,requiredNativeTableOwnerSlides:[],requiredNativeChartOwnerSlides:[],fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(buildDir,'validation-final.json')});
console.log(JSON.stringify({output:FINAL_PPTX,validation:result}));
const finalDeck=await PresentationFile.importPptx(await FileBlob.load(FINAL_PPTX));
const finalSlide=finalDeck.slides.items[0];
const finalPng=await finalDeck.export({slide:finalSlide,format:'png',scale:1});
await fs.writeFile(path.join(buildDir,'final.png'),new Uint8Array(await finalPng.arrayBuffer()));
