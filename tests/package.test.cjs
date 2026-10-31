'use strict';
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const toolRequire=require('node:module').createRequire(path.join(__dirname,'../tools/package.json'));
const CFB=toolRequire('cfb'),{unzipSync,strFromU8}=toolRequire('fflate'),iconv=toolRequire('iconv-lite'),ExcelJS=toolRequire('exceljs');
const {decompress}=require('../tools/ovba.cjs');const root=path.resolve(__dirname,'..');
let checks=0;function check(label,fn){fn();checks++;console.log('PASS',label);}
const bytes=fs.readFileSync(path.join(root,'ExcelOps.xlsm')),z=unzipSync(bytes),c=CFB.read(Buffer.from(z['xl/vbaProject.bin']));
const text=p=>strFromU8(z[p]);
check('ZIP contains macro-enabled workbook content type',()=>assert.match(text('[Content_Types].xml'),/application\/vnd.ms-excel.sheet.macroEnabled.main\+xml/));
check('VBA relationship and workbook codename',()=>{assert.match(text('xl/_rels/workbook.xml.rels'),/vbaProject.bin/);assert.match(text('xl/workbook.xml'),/codeName="ThisWorkbook"/);});
check('SHA256 matches release',()=>assert.equal(fs.readFileSync(path.join(root,'ExcelOps.xlsm.sha256'),'utf8').split(' ')[0],crypto.createHash('sha256').update(bytes).digest('hex')));
check('Project protection metadata is unprotected and visible',()=>{
 const project=CFB.find(c,'PROJECT').content.toString();
 const id=project.match(/ID="([^"]+)"/)[1],key=[...id].reduce((n,ch)=>(n+ch.charCodeAt(0))&255,0);
 for(const [name,expected] of [['CMG',Buffer.alloc(4)],['DPB',Buffer.alloc(1)],['GC',Buffer.from([255])]]){
  const hex=project.match(new RegExp(name+'="([A-F0-9]+)"'))[1],b=Buffer.from(hex,'hex');
  assert.equal(b[1]^b[0],2);assert.equal(b[2]^b[0],key);
  let prev=b[2],prev2=b[1],plainPrev=key,out=[];
  for(const e of b.subarray(3)){const p=e^((prev2+plainPrev)&255);out.push(p);prev2=prev;prev=e;plainPrev=p;}
  const value=Buffer.from(out).subarray((b[0]&6)>>1);assert.equal(value.readUInt32LE(),expected.length);assert.deepEqual(value.subarray(4),expected);
 }
});
const dir=decompress(CFB.find(c,'Root Entry/VBA/dir').content);let pos=0;const records=[];
while(pos<dir.length){const id=dir.readUInt16LE(pos),len=dir.readUInt32LE(pos+2);pos+=6;assert.ok(pos+len<=dir.length);const body=dir.subarray(pos,pos+len);pos+=len;if(id===9)pos+=2;records.push({id,body});}
check('Directory completely parsed; Turkish code page',()=>{assert.equal(pos,dir.length);assert.equal(records.find(r=>r.id===3).body.readUInt16LE(),1254);assert.equal(records.at(-1).id,0x10);});
const expected=[];
for(const directory of ['src','forms'])for(const file of fs.readdirSync(path.join(root,directory)).sort()){
 if(!/\.(bas|cls|frm)$/.test(file))continue;
 const name=path.parse(file).name;expected.push(name);
 const src=fs.readFileSync(path.join(root,directory,file),'utf8').replace(/\r?\n/g,'\r\n');const normalized=src.slice(src.indexOf('Attribute VB_Name'));
 check('Embedded source exactly matches '+name,()=>{const entry=CFB.find(c,'Root Entry/VBA/'+name);assert.ok(entry);assert.equal(iconv.decode(decompress(entry.content),'windows1254'),normalized);});
}
check('Directory module names, count and zero source offsets',()=>{
 assert.deepEqual(records.filter(r=>r.id===0x19).map(r=>r.body.toString()).sort(),expected.sort());assert.equal(records.find(r=>r.id===0xf).body.readUInt16LE(),expected.length);
 for(const r of records.filter(r=>r.id===0x31))assert.equal(r.body.readUInt32LE(),0);
});
check('Source-only project has no stale executable cache',()=>{assert.equal(CFB.find(c,'Root Entry/VBA/_VBA_PROJECT').content.toString('hex'),'cc61ffff000000');assert.ok(!c.FullPaths.some(p=>p.includes('__SRP_')));});
for(const name of ['frmMain','frmSettings'])check('Persisted empty UserForm '+name,()=>{assert.match(CFB.find(c,'PROJECT').content.toString(),new RegExp('BaseClass='+name));const f=CFB.find(c,'Root Entry/'+name+'/f').content;assert.equal(f.readUInt32LE(75),0);assert.equal(f.readUInt32LE(79),0);assert.ok(CFB.find(c,'Root Entry/'+name+'/\x01CompObj'));assert.match(CFB.find(c,'Root Entry/'+name+'/\x03VBFrame').content.toString(),new RegExp(name));});
check('Every worksheet codename has a VBA document',()=>{for(let i=1;i<=4;i++){const match=text('xl/worksheets/sheet'+i+'.xml').match(/codeName="([^"]+)"/);assert.ok(match);assert.ok(CFB.find(c,'Root Entry/VBA/'+match[1]));}});
(async()=>{
 const wb=new ExcelJS.Workbook();await wb.xlsx.load(bytes);
 check('Workbook pages and all settings are accessible',()=>{assert.deepEqual(wb.worksheets.map(s=>s.name),['Başlangıç','Ayarlar','Kılavuz','Log']);const cfg=wb.getWorksheet('Ayarlar');assert.equal(cfg.getCell('A2').value,'KaynakDosya');assert.ok(cfg.rowCount>=35);assert.ok(wb.getWorksheet('Kılavuz').rowCount>=20);});
 const cfg=wb.getWorksheet('Ayarlar'),keys=new Set();cfg.eachRow((r,i)=>{if(i>1)keys.add(r.getCell(1).value);});
 check('Static VBA setting references all exist',()=>{for(const d of ['src','forms'])for(const f of fs.readdirSync(path.join(root,d))){if(!/\.(bas|cls|frm)$/.test(f))continue;const src=fs.readFileSync(path.join(root,d,f),'utf8');for(const m of src.matchAll(/(?:Config\("|PutConfig ")([^"\r\n]+)"/g)){if(m[1]==='Filtre')continue;assert.ok(keys.has(m[1]),'Missing setting '+m[1]);}}});
 for(const f of ['PortfolioSplit.xlsx','ReconciliationB.xlsx']){const ex=new ExcelJS.Workbook();await ex.xlsx.readFile(path.join(root,'examples',f));check('Example '+f+' has input, hidden sheet, formula and dates',()=>{assert.equal(ex.getWorksheet('GizliNot').state,'hidden');assert.ok(ex.getWorksheet('Formul').getCell('A1').formula);assert.ok(ex.getWorksheet('Veri').getCell('O2').value instanceof Date);assert.equal(ex.getWorksheet('Veri').getCell('A2').value,'001');});}
 console.log(checks+' package checks passed. These do NOT execute VBA or Microsoft Excel.');
})().catch(e=>{console.error(e);process.exitCode=1;});
