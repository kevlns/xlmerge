import {readFileSync,existsSync,readdirSync} from 'node:fs';
import {resolve,join,dirname} from 'node:path';
const root=process.cwd();
const read=p=>readFileSync(join(root,p),'utf8');
const expected={
 'README.md':['功能','环境要求','安装','快速开始','使用说明','开发与验证','文档','许可与版权'],
 'AGENTS.md':['定位','调用流程','调用规范','结果验证','参考资料','许可与版权'],
};
for(const [file,headings] of Object.entries(expected)){
 const text=read(file),actual=[...text.matchAll(/^## (.+)$/gm)].map(m=>m[1]);
 if(JSON.stringify(actual)!==JSON.stringify(headings))throw Error(`${file} 文档结构不一致：${actual.join('、')}`);
 if(!text.includes('Copyright © 2026 Vant'))throw Error(`${file} 缺少统一版权`);
}
const paths=['README.md','AGENTS.md'];
function visit(path){if(!existsSync(join(root,path)))return;for(const entry of readdirSync(join(root,path),{withFileTypes:true})){const p=join(path,entry.name);if(entry.isDirectory())visit(p);else if(p.endsWith('.md'))paths.push(p);}}
visit('docs');visit('skills');
for(const path of paths){
 const text=read(path);
 if(/个人工具箱|个人开发工具|kevlns 工具家族|Part of the \*\*kevlns|First reviewable milestone/.test(text))throw Error(`${path} 包含非产品化表述`);
 let code=false;
 for(const line of text.split('\n')){
  if(line.startsWith('```')){code=!code;continue;}if(code)continue;
  for(const match of line.matchAll(/\[[^\]]*\]\(([^)]+)\)/g)){
   const target=match[1].split('#')[0];
   if(!target||/^(https?:|mailto:)/.test(target))continue;
   if(!existsSync(resolve(root,dirname(path),target)))throw Error(`${path} 链接不存在：${target}`);
  }
 }
}
if(!read('LICENSE').includes('Copyright © 2026 Vant'))throw Error('LICENSE 版权不一致');
console.log(`文档检查通过：${paths.length} 个文件，结构、品牌、版权和本地链接有效`);
