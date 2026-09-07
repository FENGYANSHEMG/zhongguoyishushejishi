import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "D:\\《中国艺术设计史》（夏燕靖 第3版）";
const sourcePath = path.join(workspaceDir, "output", "中国艺术设计史_优化课件_2026-09-07.pptx");
const skillDir = "C:\\Users\\Administrator\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.905.11957\\skills\\presentations";
const runtimePython = "C:\\Users\\Administrator\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
const stagingDir = path.join(workspaceDir, ".codex-finalizer");
const finalPath = path.join(workspaceDir, "output", "中国艺术设计史_深化课件_2026-09-07.pptx");
const candidatePath = path.join(stagingDir, "中国艺术设计史_深化课件_candidate.pptx");
const receiptPath = path.join(stagingDir, "中国艺术设计史_深化课件.validation.json");
await fs.mkdir(stagingDir, { recursive: true });
await fs.mkdir(path.dirname(finalPath), { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
const replaceText = (id, text) => { presentation.resolve(id).text = text; };
const setNotes = (slideIndex, text) => { presentation.slides.getItem(slideIndex - 1).speakerNotes.textFrame.setText(text); };

// Remove the cover's AI-generation label while retaining all cover geometry.
replaceText("sh/hofulsf2", "");

// The body text on chapter pages now follows a consistent four-level learning hierarchy:
// chapter title, period, learning focus, comparison/critical question.
replaceText("sh/z2tcnm5s", "学习主题\n生产工具与仪式表达如何共同塑造早期造物。\n跨教材比较\n《中国工艺美术史》重视石器和彩陶的工艺与门类。中国美术史重视造型和纹样。这里进一步追问器物在何种生活与仪式情境中使用。\n研讨问题\n彩陶纹样能否直接解释为宗教符号？需要哪些考古证据支持？");
replaceText("sh/87ipkzal", "学习主题\n定居生活如何推动陶器分工，玉器又如何进入礼仪与身份表达。\n跨教材比较\n工艺美术史可从器形、玉材和加工观察技术。美术史关注玉器的造型与纹样。设计史还需讨论生产组织、随葬制度和使用者。\n研讨问题\n玉琮的方外圆内为何值得从材料、工艺和礼仪三方面解释？");
replaceText("sh/ydkbm5sv", "学习主题\n青铜礼器的形制、纹样与用器制度如何形成整体。\n跨教材比较\n工艺美术史强调范铸、合金和装饰。美术史讨论兽面纹的视觉秩序。设计史可进一步连接礼制、资源动员和器物使用场景。\n研讨问题\n技术能力提升为什么没有只导向实用器，而能服务于礼制表达？");
replaceText("sh/ri9g7uhw", "学习主题\n铁器扩展与《考工记》如何显示造物中的技术理性。\n跨教材比较\n工艺美术史呈现铁器、漆器和织物门类。设计史应把时令、环境、材料和工艺视为同一条件链，观察技术如何回应生产与战争。\n研讨问题\n“材有美，工有巧”在当代材料选择中还能提供哪些判断？");
replaceText("sh/nex4jq5k", "学习主题\n秦汉的制度秩序如何进入陵墓、服饰和器物的形态。\n跨教材比较\n中国美术史常从雕塑、画像与建筑观看秦汉视觉文化。工艺美术史关注漆器、织绣和铜器。设计史还应讨论标准化、空间秩序与身份识别。\n研讨问题\n兵马俑的规模、排列与制造方式怎样共同表达帝国秩序？");
replaceText("sh/zelojyl8", "学习主题\n宗教传播与起居方式怎样改写空间、家具和造像。\n跨教材比较\n美术史关注石窟造像的样式演变。工艺美术史关注家具和器物。设计史将传播网络、石材条件和身体姿态纳入同一分析。\n研讨问题\n高坐家具的兴起，如何说明身体姿态也会影响设计形态？");
replaceText("sh/n69grmpw", "学习主题\n隋唐的开放交流怎样影响器物、都市与服饰的视觉语言。\n跨教材比较\n工艺美术史可比较金银器、三彩和丝织技术。美术史可讨论人物、壁画与装饰风格。设计史还要追踪贸易、制度和消费风尚。\n研讨问题\n跨地域交流在器物上留下的痕迹，怎样区分技术借用与审美选择？");
replaceText("sh/ra943il8", "学习主题\n宋代审美与文化传播如何共同塑造器物和视觉文化。\n跨教材比较\n美术史通常讨论文人画和书法的审美取向。工艺美术史重视瓷器与印刷。设计史把窑业组织、城市消费和文本复制放入解释框架。\n研讨问题\n“简约”是单纯的形式风格，还是与生产和使用方式有关？");
replaceText("sh/o3i5w3ad", "学习主题\n元明清的商品流通、宫廷需求与家具结构如何相互影响。\n跨教材比较\n工艺美术史以家具、陶瓷、染织和版画为主要对象。美术史强调绘画与装饰风格。设计史可考察工坊、市场、材料和人体尺度。\n研讨问题\n明式家具的结构美，怎样同时来自榫卯、硬木和使用者的身体经验？");
replaceText("sh/jepwf2pc", "学习主题\n工业化、设计教育与大众媒介如何重构近现代设计。\n跨教材比较\n中国现代设计史重视设计教育、图像传播和工业生产。20世纪工艺美术史还提示产业政策、技术改造和社会转型的重要性。\n研讨问题\n“传统工艺现代化”应如何同时评价技术效率、审美价值与社会影响？");

replaceText("sh/hsn2l4bu", "研究深化与资料来源\n教材用于建立史实框架。期刊论文用于方法讨论。中国社会科学网材料提供课程议题与研究线索。维基百科用于术语核对与延伸检索。");
replaceText("sh/itw3upcf", "教材与同行评议论文\n\n夏燕靖：《中国艺术设计史》（第3版）\n田自秉：《中国工艺美术史》\n\n应宜文：《“设计”考辨——对中国设计史的一种新认识》\n《美术研究》2018年第1期\n\n祝帅：《全球设计史观与中国设计主体性的建构》\n《艺术设计研究》2021年\n\n杭间：《中国的设计史问题》\n《美术观察》2007年第12期，第25页");
replaceText("sh/je5knut0", "中国社会科学网筛选线索\n\n《新版学科目录“设计”类学科迎来重大结构性调整的几点解读》\n将教学从器物罗列推进到人、技术、市场和制度。\n\n《20世纪中国工艺美术史》——百年工艺之旅\n近现代章节可加入产业政策、技术改造与社会转型。\n\n“无名的艺术史”研讨会报道\n可从物质文化与“无名性”深化设计史的材料观。\n\n使用边界\n以上为中国社会科学网的研究线索，不替代同行评议论文。维基条目仍仅用于课堂核对。");

const commonPrompt = "ComfyUI 配图提示词：16:9 横幅，中国艺术设计史教材插图，半写实数字绘画，暖米色纸张质感，主体置于画面右侧并留出左侧 45% 空白供文字排版，克制配色，博物馆教育图风格。负面提示词：文字、标志、水印、伪造馆藏编号、真实文物照片、夸张人物面部、畸形手部。使用边界：AI 图仅作历史情境示意，不替代文物照片、考古图版或原始史料。";
setNotes(5, `${commonPrompt}\n本页专用提示：史前河岸聚落中的磨制石器、陶器制作和抽象彩陶纹样，展示劳动与仪式并存，不复制任何具体出土文物。\n资料：维基百科《仰韶文化》 https://zh.wikipedia.org/wiki/仰韶文化 。`);
setNotes(6, `${commonPrompt}\n本页专用提示：新石器时代玉器作坊与方外圆内的抽象玉琮剪影，表现材料加工、礼仪和身份表达，不复原具体馆藏。\n资料：维基百科《良渚文化》 https://zh.wikipedia.org/wiki/良渚文化 。`);
setNotes(7, `${commonPrompt}\n本页专用提示：商周青铜铸造工坊的陶范、熔炉和礼器轮廓，表现范铸工艺与礼制秩序，不复制后母戊鼎实物。\n资料：维基百科《后母戊鼎》 https://zh.wikipedia.org/wiki/后母戊鼎 。`);
setNotes(8, `${commonPrompt}\n本页专用提示：战国工匠在材料架前选择木、金属和漆料，旁置抽象工尺与器物轮廓，表现“材有美，工有巧”的条件关系。\n资料：维基百科《考工记》 https://zh.wikipedia.org/wiki/考工记 。`);
setNotes(9, `${commonPrompt}\n本页专用提示：秦汉陵墓空间的轴线、队列和纹样化器物剪影，以示意图方式呈现秩序和标准化，不生成兵马俑实物仿制。\n资料：维基百科《秦始皇兵马俑》 https://zh.wikipedia.org/wiki/秦始皇兵马俑 。`);
setNotes(10, `${commonPrompt}\n本页专用提示：北魏石窟崖壁、脚手架和高坐家具轮廓组合，表现宗教传播、石材条件与起居方式的变化。\n资料：维基百科《云冈石窟》 https://zh.wikipedia.org/wiki/云冈石窟 。`);
setNotes(11, `${commonPrompt}\n本页专用提示：丝路商旅、金银器、低温釉陶器和都市街市的象征性拼贴，表现开放交流与消费风尚，不复制唐三彩器物。\n资料：维基百科《唐三彩》 https://zh.wikipedia.org/wiki/唐三彩 。`);
setNotes(12, `${commonPrompt}\n本页专用提示：宋代窑业、开片釉面纹理、雕版印刷木版和书卷的象征性组合，表现审美、生产和传播。\n资料：维基百科《哥窑》 https://zh.wikipedia.org/wiki/哥窑 ，《汝窑》 https://zh.wikipedia.org/wiki/汝窑 。`);
setNotes(13, `${commonPrompt}\n本页专用提示：明式椅子的榫卯节点、木材纹理、工坊工具和市集货架，强调结构、材料和身体尺度，不仿制馆藏家具。\n资料：维基百科《明式家具》 https://zh.wikipedia.org/wiki/明式家具 。`);
setNotes(14, `${commonPrompt}\n本页专用提示：近现代设计教室、印刷机、海报网格和产品草图的拼贴，表现工业化、设计教育与大众媒介。\n资料：何人可主编《中国工业设计史》；邱春林《20世纪中国工艺美术史》相关介绍。\n中国社会科学网线索： https://www.cssn.cn/ztzl/ztzl_skwzg/ggfy/jyyj/ctsgy/202405/t20240509_5750086.shtml 。`);
setNotes(17, "研究深化提示：把‘形式追随功能’改写为可讨论的问题，而不是历史定律。每个案例同时观察生产者、材料技术、使用者、制度与市场。\n中国社会科学网《新版学科目录“设计”类学科迎来重大结构性调整的几点解读》指出，中国古代设计史教学常停留在器物门类与时代背景罗列，可进一步讨论人与设计、市场与设计、技术与设计、设计思想演变。\n链接：https://www.cssn.cn/ysx/ysx_xksy/202208/t20220802_5444283.shtml 。");
setNotes(18, "资料筛选日期：2026-09-07。\n同行评议论文：应宜文，《“设计”考辨——对中国设计史的一种新认识》，《美术研究》2018年第1期；祝帅，《全球设计史观与中国设计主体性的建构》，《艺术设计研究》2021年；杭间，《中国的设计史问题》，《美术观察》2007年第12期，第25页。\n中国社会科学网研究线索：新版学科目录设计类学科解读（2022）；《20世纪中国工艺美术史》——百年工艺之旅（2024）；“无名的艺术史”物质文化与设计研究研讨会（2023）。这些材料为网刊报道或专题文章，需与同行评议论文区分使用。\n维基百科条目作为核对入口：仰韶文化、良渚文化、后母戊鼎、考工记、云冈石窟、唐三彩、哥窑、汝窑、明式家具、路易斯·沙利文。");

const snapshot = await presentation.inspect({ kind: "slide,textbox,notes", maxChars: 50000 });
await fs.writeFile(path.join(workspaceDir, ".codex-build", "deepened_snapshot.ndjson"), snapshot.ndjson, "utf8");
const montage = await presentation.export({ format: "png", montage: true, scale: 0.6 });
await fs.writeFile(path.join(workspaceDir, ".codex-build", "deepened_montage.png"), new Uint8Array(await montage.arrayBuffer()));
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
