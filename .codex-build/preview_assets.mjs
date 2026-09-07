import fs from 'node:fs/promises';
import path from 'node:path';
import sharp from 'sharp';
const dir=path.resolve('.codex-build/research-assets');
const files=(await fs.readdir(dir)).filter(f=>/^R\d+\.(jpg|png)$/.test(f)).sort();
const comps=[];
for(let i=0;i<files.length;i++){
 const input=path.join(dir,files[i]);
 const buf=await sharp(input).resize(210,180,{fit:'contain',background:'#f7f2e9'}).toBuffer();
 comps.push({input:buf,left:(i%6)*210,top:Math.floor(i/6)*210});
 const label=await sharp({text:{text:files[i].split('.')[0],font:'Arial 16',rgba:true}}).png().toBuffer();
 comps.push({input:label,left:(i%6)*210+8,top:Math.floor(i/6)*210+185});
}
await sharp({create:{width:1260,height:Math.ceil(files.length/6)*210,channels:3,background:'#f7f2e9'}}).composite(comps).png().toFile(path.join(dir,'contact.png'));
