import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "D:\\《中国艺术设计史》（夏燕靖 第3版）";
const sourcePath = path.join(workspaceDir, "中国艺术设计史_形式追随功能_课件.pptx");
const skillDir = "C:\\Users\\Administrator\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.905.11957\\skills\\presentations";
const runtimePython = "C:\\Users\\Administrator\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
const stagingDir = path.join(workspaceDir, ".codex-finalizer");
const finalPath = path.join(workspaceDir, "output", "中国艺术设计史_优化课件_2026-09-07.pptx");
const candidatePath = path.join(stagingDir, "中国艺术设计史_优化课件_candidate.pptx");
const receiptPath = path.join(stagingDir, "中国艺术设计史_优化课件.validation.json");
await fs.mkdir(stagingDir, { recursive: true });
await fs.mkdir(path.dirname(finalPath), { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));

const replaceText = (id, text) => {
  presentation.resolve(id).text = text;
};
const setNotes = (slideIndex, text) => {
  presentation.slides.getItem(slideIndex - 1).speakerNotes.textFrame.setText(text);
};

replaceText("sh/obq90bml", "实用功能\n汲水、炊煮、坐卧、农战、生产。\n器形服务于具体用途，如尖底瓶与汲水、铁器与农战。");
replaceText("sh/n6ls3alk", "礼制与象征\n等级、祭祀、礼制、天人关系。\n鼎、尊、簋与纹样受到用器制度和仪式语境的规约。");
replaceText("sh/943mhgre", "审美与生活\n文人趣味、市民生活、商品流通。\n宋瓷、明式家具与版画呈现审美取向和使用方式对形态的影响。");
replaceText("sh/v6l4jq94", "技术与材料\n《考工记》提出“天有时，地有气，材有美，工有巧”。功能提出使用要求，材料和工艺限定造物可以实现的形态。");

replaceText("sh/z2tcnm5s", "功能重心\n生存与仪式\n形式如何追随\n石器围绕砍砸、切割等劳动形成；彩陶纹样在日常器用之外，也进入仪式与观念表达。\n关键判断\n实用与象征在早期造物中已相互交织。");
replaceText("sh/87ipkzal", "功能重心\n定居生活与礼仪表达\n形式如何追随\n陶器围绕汲水、炊煮、储藏等需求形成器类；良渚玉器显示玉材在礼仪、身份与权力表达中的作用。\n关键判断\n玉器逐渐成为礼仪和身份表达的媒介，不能简单等同于实用工具。");
replaceText("sh/ydkbm5sv", "功能重心\n礼制象征\n形式如何追随\n鼎、尊、簋的形制、纹样与用器等级相互关联，体现礼制对造物的规约。\n关键判断\n礼制不仅影响器物形态，也影响材料、工艺与资源的配置。");
replaceText("sh/ri9g7uhw", "功能重心\n实用与技术理性\n形式如何追随\n铁器适应农耕和战争的需求而扩展；《考工记》把时令、环境、材料、工艺并列为造物条件。\n关键判断\n可借《考工记》理解功能、材料与工艺之间的关联。");
replaceText("sh/nex4jq5k", "功能重心\n大一统与天人秩序\n形式如何追随\n陵墓空间、兵马俑、服饰章纹与丧葬观念、等级秩序相互关联。\n关键判断\n器物与空间在此承担制度表达和身份识别的作用。");
replaceText("sh/zelojyl8", "功能重心\n宗教、审美与生活方式\n形式如何追随\n佛教推动石窟造像；高坐家具随垂足而坐等生活方式变化而发展。\n关键判断\n宗教传播和起居方式的改变，会带来器物与空间形态的调整。");
replaceText("sh/n69grmpw", "功能重心\n开放交流与帝国秩序\n形式如何追随\n金银器、都城与服饰呈现中外交流、制度秩序和生活风尚的共同影响。\n关键判断\n隋唐造物可从交流网络与社会秩序两个层面理解。");
replaceText("sh/ra943il8", "功能重心\n文人审美与市民生活\n形式如何追随\n宋瓷呈现克制、纯净的审美倾向；雕版印刷服务于文本复制与文化传播。\n关键判断\n审美取向与传播需求共同塑造宋代器物和视觉文化。");
replaceText("sh/o3i5w3ad", "功能重心\n商品流通与宫廷需求\n形式如何追随\n明式家具重视结构、比例与使用感受；青花瓷和版画的生产与流通更为活跃。\n关键判断\n明式家具可作为讨论结构、材料、人体尺度与审美关系的案例。");
replaceText("sh/jepwf2pc", "功能重心\n工业化与现代生活\n形式如何追随\n机器生产、大众消费和新媒介改变了设计的生产方式与使用场景。\n关键判断\n近现代设计把传统造物带入工业体系与现代设计教育的语境。");

replaceText("sh/snedsvml", "三重视角  中国古代造物可从实用、礼仪象征与审美生活三个维度观察。\n礼仪象征是理解不少器物形制、纹样和使用场景的重要线索，但不应取代对材料与使用方式的分析。");
replaceText("sh/pc7qhwn6", "技术与材料  功能提出使用要求，材料性能与工艺能力共同限定形态的实现。\n以《考工记》为例，可把时令、环境、材料和工艺放入同一造物条件中考察。");
replaceText("sh/honqdwnu", "历史变化  功能重心会随生产方式、礼仪制度、宗教传播和消费生活而改变。\n从史前工具到近现代设计，器物形态始终与具体的社会需求和技术条件相互关联。");

replaceText("sh/hsn2l4bu", "研究与资料来源\n教材提供章节脉络；期刊论文用于方法与史学问题；维基百科条目用于课堂核对与延伸检索。");
replaceText("sh/itw3upcf", "教材与学术期刊\n\n夏燕靖：《中国艺术设计史》（第3版）\n课件章节与基本叙事依据\n\n应宜文：《“设计”考辨——对中国设计史的一种新认识》\n《美术研究》2018年第1期\n\n祝帅：《全球设计史观与中国设计主体性的建构》\n《艺术设计研究》2021年\n\n杭间：《中国的设计史问题——关于上海的中国设计史年会》\n《美术观察》2007年第12期，第25页");
replaceText("sh/je5knut0", "维基百科核对条目\n\n仰韶文化、良渚文化、后母戊鼎、考工记\n秦始皇兵马俑、云冈石窟、唐三彩\n哥窑、汝窑、明式家具、路易斯·沙利文\n\n使用提示\n维基条目适合核对年代、地点、术语和关联条目。课堂论证应回到教材、考古报告、博物馆资料或学术论文。\n\n链接入口\nhttps://zh.wikipedia.org/wiki/仰韶文化\nhttps://zh.wikipedia.org/wiki/良渚文化\nhttps://zh.wikipedia.org/wiki/后母戊鼎\nhttps://zh.wikipedia.org/wiki/路易斯·沙利文");

setNotes(3, "课程主张：以功能、材料、技术、制度和审美的关系分析造物。\n参考：夏燕靖《\u4e2d\u56fd\u827a\u672f\u8bbe\u8ba1\u53f2》（第3版）；应宜文《“设计”考辨——对中国设计史的一种新认识》，《美术研究》2018年第1期。\n提示：‘形式追随功能’是用于课堂比较的分析框架，不等同于中国古代造物的原生术语。");
setNotes(5, "核对条目：维基百科《仰韶文化》 https://zh.wikipedia.org/wiki/仰韶文化 （访问：2026-09-07）。条目介绍其农业、石器、日用陶器和彩绘纹样。\n提示：纹样的具体象征含义需结合发掘报告和专业研究，不宜作单一解释。");
setNotes(6, "核对条目：维基百科《良渚文化》 https://zh.wikipedia.org/wiki/良渚文化 （访问：2026-09-07）。条目介绍良渚文化的年代、玉琮、玉璧、玉钺及礼仪与权力表达。\n课堂追问：礼仪功能如何改变材料选择与加工成本？");
setNotes(7, "核对条目：维基百科《后母戊鼎》 https://zh.wikipedia.org/wiki/后母戊鼎 （访问：2026-09-07）。条目列出其年代、尺寸、纹样与铸造信息，并提示‘后母戊/司母戊’释读存在讨论。\n提示：名称与铭文问题应保留学术争议。");
setNotes(8, "核对条目：维基百科《考工记》 https://zh.wikipedia.org/wiki/考工记 （访问：2026-09-07）。\n参考：应宜文《“设计”考辨——对中国设计史的一种新认识》，《美术研究》2018年第1期。\n课堂追问：文本中的‘材有美、工有巧’如何联系当代材料与工艺选择？");
setNotes(9, "核对条目：维基百科《秦始皇兵马俑》 https://zh.wikipedia.org/wiki/秦始皇兵马俑 （访问：2026-09-07）。\n课堂追问：丧葬制度如何影响规模、复制方式和视觉秩序？");
setNotes(10, "核对条目：维基百科《云冈石窟》 https://zh.wikipedia.org/wiki/云冈石窟 （访问：2026-09-07）。\n课堂追问：宗教传播、皇家支持和石材条件分别怎样影响石窟形态？");
setNotes(11, "核对条目：维基百科《唐三彩》 https://zh.wikipedia.org/wiki/唐三彩 （访问：2026-09-07）。\n课堂追问：随葬用途、釉彩工艺和跨地域交流如何共同影响器物风格？");
setNotes(12, "核对条目：维基百科《哥窑》 https://zh.wikipedia.org/wiki/哥窑 、《汝窑》 https://zh.wikipedia.org/wiki/汝窑 （访问：2026-09-07）。\n提示：窑口认定、传世品与考古材料的关系应查阅陶瓷史专著和博物馆资料。");
setNotes(13, "核对条目：维基百科《明式家具》 https://zh.wikipedia.org/wiki/明式家具 （访问：2026-09-07）。\n课堂追问：榫卯、硬木材料和人体尺度怎样共同形成‘结构美’？");
setNotes(14, "核对条目：维基百科《路易斯·沙利文》 https://zh.wikipedia.org/wiki/路易斯·沙利文 （访问：2026-09-07）。\n背景：1896年论文《The Tall Office Building Artistically Considered》常与‘form ever follows function’相关联。此口号用于比较中西设计思想时，应注意其历史语境。");
setNotes(17, "参考：祝帅《全球设计史观与中国设计主体性的建构——兼谈“世界设计史”研究的学科意义及学术启示》，《艺术设计研究》2021年；杭间《中国的设计史问题——关于上海的中国设计史年会》，《美术观察》2007年第12期。\n课堂提示：避免把中国设计史讲成孤立的器物史，也避免用单一西方功能主义概念覆盖全部中国经验。");
setNotes(18, "学术期刊检索与条目核对日期：2026-09-07。\n期刊：应宜文，《“设计”考辨——对中国设计史的一种新认识》，《美术研究》2018年第1期；祝帅，《全球设计史观与中国设计主体性的建构》，《艺术设计研究》2021年；杭间，《中国的设计史问题——关于上海的中国设计史年会》，《美术观察》2007年第12期，第25页。\n维基百科条目用于课堂检索入口，不能替代学术来源。\n链接： https://zh.wikipedia.org/wiki/仰韶文化 ; https://zh.wikipedia.org/wiki/良渚文化 ; https://zh.wikipedia.org/wiki/后母戊鼎 ; https://zh.wikipedia.org/wiki/路易斯·沙利文");

const updateSnapshot = await presentation.inspect({
  kind: "slide,textbox,notes",
  include: "id,slide,title,name,textPreview,textChars",
  maxChars: 30000,
});
await fs.writeFile(path.join(workspaceDir, ".codex-build", "updated_snapshot.ndjson"), updateSnapshot.ndjson, "utf8");
const montage = await presentation.export({ format: "png", montage: true, scale: 0.6 });
await fs.writeFile(path.join(workspaceDir, ".codex-build", "updated_montage.png"), new Uint8Array(await montage.arrayBuffer()));
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);

const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir, "container_tools", "artifact_tool_utils.mjs")).href);
await finalizePresentation({
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: runtimePython,
  integrityValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12191695,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
  explicitTotalSlideCount: 19,
  requiredNativeTableOwnerSlides: [],
  verifyArtifactToolImport: true,
  receiptPath,
});
console.log(finalPath);
