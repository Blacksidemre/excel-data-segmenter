'use strict';
const assert = require('node:assert/strict');
const u16=n=>{const b=Buffer.alloc(2);b.writeUInt16LE(n);return b;};
const u32=n=>{const b=Buffer.alloc(4);b.writeUInt32LE(n);return b;};
const rec=(id,data=Buffer.alloc(0))=>Buffer.concat([u16(id),u32(data.length),data]);
// MS-OVBA compressed containers: literal-only compressed chunks avoid padding source text.
function compress(data){
 const chunks=[Buffer.from([1])];
 for(let p=0;p<data.length;p+=3600){const raw=data.subarray(p,p+3600),body=[];
  for(let j=0;j<raw.length;j+=8)body.push(Buffer.from([0]),raw.subarray(j,j+8));
  const bytes=Buffer.concat(body);chunks.push(u16(0xb000|(bytes.length-1)),bytes);
 }
 return Buffer.concat(chunks);
}
function decompress(data){
 assert.equal(data[0],1);const out=[];let p=1;
 while(p<data.length){const start=p,h=data.readUInt16LE(p);p+=2;const end=start+(h&4095)+3,base=out.length;
  assert.equal((h>>12)&7,3);assert.ok(end<=data.length);
  if(!(h&32768)){out.push(...data.subarray(p,end));p=end;continue;}
  while(p<end){const flags=data[p++];for(let bit=0;bit<8&&p<end;bit++){
   if(!(flags&(1<<bit)))out.push(data[p++]);else{const token=data.readUInt16LE(p);p+=2;
    const bits=Math.max(4,Math.ceil(Math.log2(out.length-base))),len=(token&((1<<(16-bits))-1))+3,off=(token>>(16-bits))+1;
    assert.ok(off<=out.length-base);for(let i=0;i<len;i++)out.push(out[out.length-off]);
   }
  }}assert.ok(out.length-base<=4096);
 }return Buffer.from(out);
}
function reference(name,guid,version,path,description){
 const lib=Buffer.from(`*\\G{${guid}}#${version}#0#${path}#${description}`,'ascii');
 return Buffer.concat([rec(0x16,Buffer.from(name)),rec(0x3e,Buffer.from(name,'utf16le')),rec(0x0d,Buffer.concat([u32(lib.length),lib,u32(0),u16(0)]))]);
}
function directory(modules){
 const header=[rec(1,u32(1)),rec(2,u32(0x409)),rec(0x14,u32(0x409)),rec(3,u16(1254)),rec(4,Buffer.from('ExcelOps')),
 rec(5),rec(0x40),rec(6),rec(0x3d),rec(7,u32(0)),rec(8,u32(0)),u16(9),u32(4),u32(1),u16(0),rec(0xc),rec(0x3c),
 reference('stdole','00020430-0000-0000-C000-000000000046','2.0','C:\\Windows\\System32\\stdole2.tlb','OLE Automation'),
 reference('Office','2DF8D04C-5BFA-101B-BDE5-00AA0044DE52','2.0','C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL','Microsoft Office 16.0 Object Library'),
 reference('MSForms','0D452EE1-E08F-101A-852E-02608C4D0BB4','2.0','C:\\Windows\\System32\\FM20.DLL','Microsoft Forms 2.0 Object Library'),
 rec(0xf,u16(modules.length)),rec(0x13,u16(0))];
 for(const m of modules)header.push(rec(0x19,Buffer.from(m.name)),rec(0x47,Buffer.from(m.name,'utf16le')),rec(0x1a,Buffer.from(m.name)),rec(0x32,Buffer.from(m.name,'utf16le')),
 rec(0x1c),rec(0x48),rec(0x31,u32(0)),rec(0x1e,u32(0)),rec(0x2c,u16(0)),rec(m.type==='module'?0x21:0x22),...(m.type==='form'?[rec(0x28)]:[]),rec(0x2b));
 header.push(rec(0x10));return Buffer.concat(header);
}
// MS-OVBA 2.4.3 PROJECT metadata encoding (not a security boundary).
function projectData(id,data,seed=0x55){
 const key=[...id].reduce((sum,c)=>(sum+c.charCodeAt(0))&255,0);
 let previous=seed^key,previous2=seed^2,plainPrevious=key;
 const out=[seed,previous2,previous];
 const raw=Buffer.concat([Buffer.alloc((seed&6)>>1,0x7a),u32(data.length),data]);
 for(const plain of raw){const encoded=plain^((previous2+plainPrevious)&255);out.push(encoded);previous2=previous;previous=encoded;plainPrevious=plain;}
 return Buffer.from(out).toString('hex').toUpperCase();
}
module.exports={compress,decompress,directory,rec,u16,u32,projectData};
