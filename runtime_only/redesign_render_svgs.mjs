import fs from 'node:fs/promises';
import path from 'node:path';
import { requireRuntimeModule } from '/root/.codex/skills/builtins/presentations/container_tools/runtime_helpers.mjs';
const sharp=await requireRuntimeModule('sharp');
const root='/workspace/scratch/242dd116dead/gaza/redesign/assets';
for (const dir of [root,path.join(root,'icons')]){
 for (const name of await fs.readdir(dir)){
  if(name.endsWith('.svg')){
   const src=path.join(dir,name);const dst=src.replace(/\.svg$/,'.png');
   await sharp(src,{density:200}).png().toFile(dst);
  }
 }
}
console.log('SVGs rendered');
