import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {Presentation,PresentationFile,FileBlob} from '@oai/artifact-tool';
import sharp from 'sharp';
import {chapters,units} from './course_content.mjs';

const root=path.resolve('.');
const privateDir=path.join(root,'.codex-build');
const skillDir='C:\\Users\\Administrator\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.905.11957\\skills\\presentations';
const runtime='C:\\Users\\Administrator\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies';
process.env.RUNTIME_NODE_MODULES=path.join(runtime,'node','node_modules');
const fontPolicy={basis:'design',families:['Microsoft YaHei','SimSun']};
const finalPath=path.join(root,'output','中国艺术设计史_十章分期三论扩展版_2026-09-07.pptx');
const candidatePath=path.join(root,'.codex-finalizer','expanded-candidate.pptx');
const renderDir=path.join(privateDir,'expanded-renders');
await fs.mkdir(renderDir,{recursive:true});
const C={paper:'#F7F2E9',ink:'#292D2C',red:'#A83D2C',teal:'#244955',muted:'#60635E',white:'#FFFFFF'};
const W=12191695/9525,H=720;
const p=Presentation.create({slideSize:{width:W,height:H}});
const sourceArray=JSON.parse(await fs.readFile(path.join(privateDir,'sources.json'),'utf8'));
const sources={};
for(const s of sourceArray){
 sources[s.id]={...JSON.parse(await fs.readFile(path.join(privateDir,'research-assets',s.id+'.json'),'utf8')),...s};
}
sources.R06.imagePath=path.join(privateDir,'research-assets','R06.jpg');
sources.R06.imageUrl='https://www.shanghaimuseum.net/mu/upload/202402/04ab5a7a-6afa-4efe-a72f-0a0341fca4d9.jpg';
sources.R27.imagePath=path.join(privateDir,'research-assets','R27-correct.jpg');
sources.R27.imageUrl='https://www.nxkg.org.cn/uploadfile/2019/0131/20190131093411325.jpg';
const academic=[
 {id:'P01',title:'中国国家博物馆藏错金银云纹铜犀尊制作工艺研究',org:'刘薇、吴娜、张鹏宇、王克青、杨琴、王力之、朱晓云',detail:'《中国国家博物馆馆刊》2024年第3期。本课采用官网摘要，未引用未核验页码。',url:'https://www.chnmuseum.cn/yj/xscg/xslw/202404/t20240415_266361.shtml'},
 {id:'P02',title:'新版学科目录“设计”类学科迎来重大结构性调整的几点解读',org:'夏燕靖（按文内作者简介）',detail:'中国社会科学网转载，页面标注来源为《工业工程设计》。仅采用设计史教学方法讨论，不据此说明当前学科政策。网页署名与简介用字不一致。',url:'https://www.cssn.cn/ysx/ysx_xksy/202208/t20220802_5444283.shtml'},
 {id:'P03',title:'全球设计史观与中国设计主体性的建构',org:'祝帅',detail:'副题：兼谈“世界设计史”研究的学科意义及学术启示。《艺术设计研究》，2021年。以北京大学机构库的题录和摘要为依据，不作全文引文。',url:'https://ir.pku.edu.cn/handle/20.500.11897/612802'},
 {id:'P04',title:'《20世纪中国工艺美术史》——百年工艺之旅',org:'中国社会科学网',detail:'2024-05-09图书介绍。用于了解邱春林著作的研究范围，不替代原书内容。',url:'https://www.cssn.cn/ztzl/ztzl_skwzg/ggfy/jyyj/ctsgy/202405/t20240509_5750086.shtml'}
];
for(const a of academic)sources[a.id]=a;
const wiki=[
 {id:'W01',title:'仰韶文化',url:'https://zh.wikipedia.org/wiki/仰韶文化'},
 {id:'W02',title:'良渚文化',url:'https://zh.wikipedia.org/wiki/良渚文化'},
 {id:'W03',title:'中国陶瓷史',url:'https://zh.wikipedia.org/wiki/中國陶瓷史'},
 {id:'W04',title:'大龙邮票',url:'https://zh.wikipedia.org/wiki/大龙邮票'},
 {id:'W05',title:'月份牌',url:'https://zh.wikipedia.org/wiki/月份牌'}
];
const manifest=[];
function tx(s,str,x,y,w,h,size=25,color=C.ink,bold=false,font='Microsoft YaHei'){
 const t=s.shapes.add({geometry:'textbox',name:'text-'+s.shapes.items?.length,position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 t.text=str;t.text.style={typeface:font,fontSize:size,bold,color,autoFit:'none'};
 return t;
}
function slide(label,title,refs=[],notes=''){
 const s=p.slides.add();s.background.fill=C.paper;
 tx(s,label,58,28,1120,34,19,C.red,true);
 tx(s,title,58,77,1154,75,42,C.ink,true,'SimSun');
 const n=p.slides.count;
 tx(s,'中国艺术设计史  /  十章分期研读',58,673,600,24,15,C.muted);
 tx(s,String(n).padStart(3,'0'),1150,673,70,24,15,C.muted);
 const rs=refs.map(id=>sources[id]).filter(Boolean);
 tx(s,rs.map(a=>`[${a.id}] ${a.org}《${a.title}》`).join('  '),58,634,1154,35,15,C.muted);
 s.speakerNotes.textFrame.setText(`${notes}\n\n资料与图片核验日期：2026-09-07。\n${rs.map(a=>`[${a.id}] ${a.org}：《${a.title}》。${a.detail??''}\n正文来源：${a.url}\n${a.imageUrl?'图片原始地址：'+a.imageUrl+'\n图片来源：'+a.org+'公开页面。保留原图水印与完整比例。':''}`).join('\n\n')}\n\n三论阐释、问题和课堂练习为教学分析，不是古人原话。馆藏断代与研究推定在正文区分。图片版权归原机构或权利人，仅供本课学习研究参考，公开再发行或商业使用须另核授权。`);
 manifest.push({page:n,label,title,refs,notes});
 return s;
}
function block(s,label,text,y,{x=58,w=1154,size=25}={}){
 tx(s,label,x,y,w,37,26,C.teal,true);
 tx(s,text,x,y+44,w,100,size);
}
function three(s,parts,start=175,gap=142){
 const labels=['一、本体论','二、认识论','三、实践论'];
 for(let i=0;i<3;i++){
   const y=start+i*gap;
   tx(s,labels[i],58,y,223,41,29,i===0?C.red:C.teal,true);
   tx(s,['对象与价值','观念与证据','制作与使用'][i],58,y+47,214,34,20,C.muted);
   tx(s,parts[i],294,y,910,119,25.5);
 }
}
const cover=slide('课程扩展版','',[]);
tx(cover,'中国艺术设计史',76,164,1140,100,70,C.ink,true,'SimSun');
tx(cover,'一至十章分期研读',81,285,1100,62,36,C.red,true);
tx(cover,'本体论  ·  认识论  ·  实践论',81,397,1100,55,32,C.teal);
tx(cover,'时代讲解、实物研读与研究讨论',81,480,1100,42,25,C.muted);
tx(cover,'依据现有课件十章主题扩展  /  2026年9月',81,555,1100,35,21,C.muted);
manifest[0].title='中国艺术设计史：一至十章分期研读';
{
 const s=slide('课程方法  /  阅读层级','三种角度与四级标题',['P02']);
 three(s,[
 '本体论追问设计对象是什么，以及它与人、社会和环境构成何种关系。器物既有材料结构，也可能承担礼仪、记忆、身份和信息传递的作用。',
 '认识论兼顾历史中的造物知识与今天的研究方法。将观察、记录、研究解释与课堂推论分开，检验我们凭什么认定年代、用途或工艺。',
 '实践论讨论制作组织、使用动作与社会后果。课堂练习以复制品、纸模或文本分析进行，不操作文物。一级为章，二级为时代专题，三级为三论，四级为具体观察或论点。']);
}
for(const half of [chapters.slice(0,5),chapters.slice(5)]){
 const s=slide('课程导航','十章与时代单元');
 half.forEach((c,i)=>{const y=166+i*88;tx(s,String(c.n).padStart(2,'0'),58,y,74,50,34,C.red,true);tx(s,c.title,154,y,490,43,29,C.ink,true);tx(s,c.scope,662,y+3,550,55,23,C.teal);});
}
for(const c of chapters){
 const selected=units.filter(u=>u.chapter===c.n);
 const s=slide(`第${c.n}章  /  章节导论`,c.title,selected.map(u=>u.ref).slice(0,1));
 tx(s,c.scope,58,160,1154,43,27,C.red,true);
 block(s,'一、时代与核心问题',c.intro,222);
 block(s,'二、跨学科阅读',c.compare,399);
 tx(s,c.question,58,576,1154,40,26,C.teal,true);
 for(let j=0;j<selected.length;j++){
  const u=selected[j],src=sources[u.ref];
  const refs=[u.ref,...(u.ref==='R12'?['P01']:[])];
  const label=`第${c.n}章  /  第${j+1}节  ${u.era}`;
  const theory=slide(label,u.topic,refs,`时代范围：${u.dates}\n${u.ontology}\n${u.epistemology}\n${u.praxis}\n推荐课堂顺序：先讲对象与历史环境，再检验证据，最后讨论制作与使用。`);
  tx(theory,u.dates,58,143,1148,30,20,C.muted);
  three(theory,[u.ontology,u.epistemology,u.praxis],187,140);
  const artifactTitle=u.ref==='R27'?'西夏文佛经《本续》':u.ref==='R32'?'杭穉英画室广告':u.ref==='R33'?'红旗牌轿车':u.ref==='R34'?'《汉字信息处理》报版样张':src.title;
  const caseSlide=slide(label+'  /  实物研读',artifactTitle,refs,`${u.record}\n\n1.观察：${u.observe}\n2.解释：${u.interpret}\n3.讨论：${u.question}\n\n图片为所引原网页上的对应实物或刊文图版。图像未用于生成、补绘或替代文物实照。`);
  tx(caseSlide,u.record,58,153,1154,73,22,C.muted);
  const bytes=await fs.readFile(src.imagePath);
  caseSlide.images.add({blob:new Uint8Array(bytes),contentType:(await sharp(bytes).metadata()).format==='png'?'image/png':'image/jpeg',alt:`${artifactTitle}。来源：${src.org}。${src.url}`,fit:'contain',position:{left:58,top:238,width:557,height:368}});
  tx(caseSlide,'1  观察',655,238,551,34,24,C.teal,true);
  tx(caseSlide,u.observe,655,278,551,101,23);
  tx(caseSlide,'2  解释与边界',655,383,551,34,24,C.teal,true);
  tx(caseSlide,u.interpret,655,423,551,102,23);
  tx(caseSlide,'3  研讨问题',655,529,551,32,24,C.red,true);
  tx(caseSlide,u.question,655,568,551,57,22);
 }
 const end=slide(`第${c.n}章  /  比较与应用`,'章末研讨：'+c.question,selected.slice(0,2).map(u=>u.ref));
 block(end,'一、比较对象',selected.map(u=>u.era+'：'+u.topic).join('。')+'。',169,{size:24});
 block(end,'二、论证要求',c.compare,337,{size:25});
 block(end,'三、学习任务',c.task,496,{size:25});
}
{
 const s=slide('研究深化  /  技术论文','从外观判断到检测证据',['P01']);
 block(s,'论文问题', '犀尊表面的金银纹饰如何制作？肉眼看到线条，无法直接确定嵌槽怎样形成。该研究把材料、微观形貌与器表分布联系起来。',168);
 block(s,'摘要中的可用结论','作者采用X射线相关分析及显微观察，认为嵌槽主要通过铸造形成，并进一步支持金银丝嵌错的判断。课件只采用公开摘要中明确表述的结果。',327);
 block(s,'课堂研究任务','分别提出“铸槽”与“后刻槽”会留下的痕迹。阅读原文时检查检测位置、比较样本与替代解释，而非只摘录结论。',491);
}
{
 const s=slide('研究深化  /  中国社会科学网','设计史中的人、技术与制度',['P02','P04']);
 block(s,'文本依据','夏燕靖的文章讨论设计史教学中的器物罗列问题，并提出联系人的活动、技术与市场。《20世纪中国工艺美术史》的介绍则提供20世纪行业与社会变迁的阅读线索。',168);
 block(s,'可深化的地方性研究','以本地窑场、木工作坊、印刷企业或老品牌为对象，追踪材料来源、订单、工种分工与用户。把地方资料与通史叙述对读，观察相同年代内的差异。',329);
 block(s,'史料边界','中国社会科学网转载页与图书介绍不自动等同于一篇已阅读全文的期刊论文。本课标出材料类型。正式论文须补核原刊、页码、版本及原始史料。',490);
}
{
 const s=slide('研究深化  /  方法比较','中国设计与全球设计史',['P03']);
 block(s,'摘要提示的概念问题','祝帅的论文讨论“设计史／设计的历史”“世界设计史／全球设计史”“中国设计／设计在中国”等问题。这些区分能帮助界定研究对象，而不是给作品贴国别标签。',168);
 block(s,'跨章节应用','以唐代载乐俑、辽代皮囊壶和近代广告为例，分别追踪图像、材料、技法与生产者的流动。形式相似只能提示联系，仍要寻找可验证的传播路径。',330);
 block(s,'小论文题目','任选一件案例，说明它的地方生产环境与跨地域联系。全文区分事实、解释及仍待解决的问题，引用至少一份馆藏记录和一篇可核验的研究文献。',492);
}
{
 const s=slide('参考阅读  /  教材与通史','三类教材的交叉阅读');
 block(s,'工艺美术史','田自秉《中国工艺美术史（修订本）》，东方出版中心，2010年。按材料、门类、制作方法与时代特征建立器物知识。',169);
 block(s,'中国美术史','薄松年主编《中国美术史教程》。将造像、绘画、装饰的形式与历史环境放在一起分析。具体版本以任课教师采用本为准。',328);
 block(s,'中国现代设计史研究','近现代部分可延伸阅读文字、媒介、工业生产及行业组织研究，并结合邱春林《20世纪中国工艺美术史》。本课未取得上述教材全文，不提供未经核实的页码或逐页引文。',487);
 s.speakerNotes.textFrame.setText('课程以用户现有课件的十章主题为基础重组，不能据此声称已与夏燕靖第3版逐页比对。\n田自秉书目信息：https://caixuan.aijiaocai.com/textbook/details?textbook_id=328097\n薄松年书目入口（南京大学参考书目）：https://art.nju.edu.cn/87/71/c55330a624497/page.htm\n邱春林著作介绍：https://www.cssn.cn/ztzl/ztzl_skwzg/ggfy/jyyj/ctsgy/202405/t20240509_5750086.shtml\n跨教材比较为本课的阅读建议，不是对未读取教材章节的直接转述。');
}
{
 const s=slide('参考阅读  /  维基百科','术语入口与交叉核验');
 tx(s,'维基百科用于查找术语、年代线索与进一步参考文献。器物断代、工艺和图片，以对应馆藏或研究记录为主要依据。',58,167,1154,87,25);
 wiki.forEach((a,i)=>{tx(s,`[${a.id}] ${a.title}`,58,277+i*61,300,38,25,C.teal,true);tx(s,a.url,368,280+i*61,850,39,21);});
 s.speakerNotes.textFrame.setText('维基百科为协作百科，页面持续修订。本课检索日期2026-09-07。\n'+wiki.map(a=>`[${a.id}] ${a.title}\n${a.url}`).join('\n\n')+'\n引用时应保留具体版本链接或访问日期，并核对条目列出的原始文献。本课不以百科推测替代馆藏信息。');
}
for(let start=0;start<sourceArray.length;start+=7){
 const slice=sourceArray.slice(start,start+7);
 const s=slide('引文目录  /  实物与图版',`资料索引 ${Math.floor(start/7)+1}：${slice[0].id}—${slice.at(-1).id}`);
 slice.forEach((a,i)=>{const y=163+i*65;tx(s,`[${a.id}] ${a.org}《${a.title}》`,58,y,1154,34,22,C.teal,true);tx(s,a.url,58,y+33,1154,30,15.5,C.ink);});
 s.speakerNotes.textFrame.setText(slice.map(a=>{const d=sources[a.id];return `[${a.id}] ${a.org}：《${a.title}》\n来源：${a.url}\n图片：${d.imageUrl}\n课件相关页：${manifest.filter(m=>m.refs.includes(a.id)).map(m=>m.page).join('、')}\n访问日期：2026-09-07。图片版权归原机构或权利人，公开再发行须核授权。`;}).join('\n\n'));
}
for(const slice of [academic.slice(0,2),academic.slice(2)]){
 const s=slide('引文目录  /  研究文献','论文与研究材料');
 slice.forEach((a,i)=>{const y=170+i*230;tx(s,`[${a.id}] ${a.title}`,58,y,1154,65,29,C.teal,true);tx(s,a.org+'。'+a.detail,58,y+70,1154,99,23);tx(s,a.url,58,y+178,1154,40,16);});
 s.speakerNotes.textFrame.setText(slice.map(a=>`[${a.id}] ${a.org}：《${a.title}》。${a.detail}\n${a.url}`).join('\n\n'));
}
{
 const s=slide('结课  /  证据与应用','一件器物的研究报告');
 three(s,[
 '明确对象、年代与出土地或收藏来源。分别描述材料、结构与形式，解释它在特定人群和环境中的用途，避免从一件精品概括整个时代。',
 '写清每个判断的证据。至少比较一条馆藏记录、一份研究文献与一件相关实物。遇到断代或用途分歧时，保留不同解释及其根据。',
 '复原可能的制作与使用过程，提出一项可检验的问题。最终报告包含实物图片、图注、完整引用和你自己的论证，现代设计启示必须说明适用条件。']);
}

await fs.writeFile(path.join(privateDir,'expanded-manifest.json'),JSON.stringify(manifest,null,2));
await fs.writeFile(path.join(privateDir,'expanded-sources.json'),JSON.stringify(sources,null,2));
await (await PresentationFile.exportPptx(p)).save(candidatePath);
console.log('DRAFT slides='+p.slides.count+' path='+candidatePath);
// Render draft examples before finalizing, then final import renders cover every page.
for(const i of [0,1,4,5,6,9,10,42,76,p.slides.count-2]){
 const png=await p.slides.getItem(i).export({format:'png',scale:1});
 await fs.writeFile(path.join(renderDir,`draft-${String(i+1).padStart(3,'0')}.png`),new Uint8Array(await png.arrayBuffer()));
}
if(process.argv.includes('--draft-only'))process.exit(0);
const {finalizePresentation}=await import(pathToFileURL(path.join(skillDir,'container_tools','artifact_tool_utils.mjs')).href);
await finalizePresentation({workspaceDir:root,candidatePath,finalPath,pythonExecutable:path.join(runtime,'python','python.exe'),integrityValidatorPath:path.join(skillDir,'container_tools','inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skillDir,'container_tools','inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','12191695,6858000','--validate-bullet-geometry','--validate-heading-fit'],fontPolicy,explicitTotalSlideCount:p.slides.count,requiredNativeTableOwnerSlides:[],verifyArtifactToolImport:true,receiptPath:path.join(root,'.codex-finalizer','expanded.validation.json')});
console.log('FINAL '+finalPath);
