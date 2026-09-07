import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "D:\\《中国艺术设计史》（夏燕靖 第3版）";
const sourcePath = path.join(workspaceDir, "output", "中国艺术设计史_优化课件_2026-09-07.pptx");
const outDir = path.join(workspaceDir, ".codex-build", "final-renders");
await fs.mkdir(outDir, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
const snapshot = await presentation.inspect({ kind: "slide,textbox,notes", maxChars: 50000 });
await fs.writeFile(path.join(outDir, "final_snapshot.ndjson"), snapshot.ndjson, "utf8");
const montage = await presentation.export({ format: "png", montage: true, scale: 0.6 });
await fs.writeFile(path.join(outDir, "final_montage.png"), new Uint8Array(await montage.arrayBuffer()));
for (let index = 0; index < presentation.slides.count; index += 1) {
  const png = await presentation.slides.getItem(index).export({ format: "png", scale: 1.5 });
  await fs.writeFile(path.join(outDir, `slide-${String(index + 1).padStart(2, "0")}.png`), new Uint8Array(await png.arrayBuffer()));
}
console.log(`slides=${presentation.slides.count}`);
