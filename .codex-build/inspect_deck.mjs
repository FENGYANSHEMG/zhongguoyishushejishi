import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const sourcePath = "D:\\《中国艺术设计史》（夏燕靖 第3版）\\中国艺术设计史_形式追随功能_课件.pptx";
const outDir = "D:\\《中国艺术设计史》（夏燕靖 第3版）\\.codex-build";
const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
const snapshot = await presentation.inspect({
  kind: "deck,slide,textbox,shape,image,table,chart,notes,layout",
  include: "id,slide,name,title,textPreview,textChars,textLines,bbox,bboxUnit,rows,cols,chartType,isPlaceholder,placeholders",
  maxChars: 50000,
});
await fs.writeFile(`${outDir}\\deck_snapshot.ndjson`, snapshot.ndjson, "utf8");
const montage = await presentation.export({ format: "png", montage: true, scale: 0.6 });
await fs.writeFile(`${outDir}\\source_montage.png`, new Uint8Array(await montage.arrayBuffer()));
for (let index = 0; index < presentation.slides.count; index += 1) {
  const slide = presentation.slides.getItem(index);
  const png = await slide.export({ format: "png", scale: 1 });
  await fs.writeFile(`${outDir}\\source_slide_${String(index + 1).padStart(2, "0")}.png`, new Uint8Array(await png.arrayBuffer()));
}
console.log(`slides=${presentation.slides.count}`);
console.log(presentation.masters.items.map(item => `${item.name}:${item.id}`).join("\n"));
console.log(presentation.layouts.items.map(item => `${item.name}:${item.id}`).join("\n"));
