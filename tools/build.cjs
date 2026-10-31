'use strict';
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const CFB=require('cfb'),ExcelJS=require('exceljs'),{zipSync,unzipSync,strToU8,strFromU8}=require('fflate'),iconv=require('iconv-lite');
const {compress,directory,projectData}=require('./ovba.cjs');
const root=path.resolve(__dirname,'..');const read=p=>fs.readFileSync(path.join(root,p));
const modules=[];
for(const dir of ['src','forms'])for(const file of fs.readdirSync(path.join(root,dir)).sort()){
 if(!/\.(bas|cls|frm)$/.test(file))continue;
 let source=read(dir+'/'+file).toString('utf8').replace(/\r?\n/g,'\r\n');
 source=source.slice(source.indexOf('Attribute VB_Name'));
 modules.push({name:path.parse(file).name,type:file.endsWith('.bas')?'module':file.endsWith('.frm')?'form':(file==='ThisWorkbook.cls'||file.startsWith('Sheet'))?'document':'class',source});
}
const settings=[
 ['KaynakDosya','','Kaynak dosya','Kaynak seç düğmesini kullanın.'],
 ['Sayfa','Veri','Veri sayfası','Kaynak çalışma kitabındaki tam sayfa adı.'],
 ['Baslik','1','Başlık satırı','Bu satırın altındaki tüm kullanılan satırlar veri sayılır; dip toplam eklemeyin.'],
 ['GrupSutunu','M','Ayırma / renk / özet sütunu','Harf veya sayı; örnek M ya da 13.'],
 ['AdSutunu','M','Dosya adı sütunu','Her grupta tek dosya adı olmalı.'],
 ['Konum','','Kayıt ana klasörü','Kayıt konumu düğmesiyle mevcut bir klasör seçin.'],
 ['Klasor','30.10.2026 Masraflar','Çıktı klasörü adı','Ana klasör içinde oluşturulur; her işlem ayrı tarih-saat alt klasörü kullanır.'],
 ['AdFormati','{ad}','Dosya adı formatı','Alanlar: {ad}, {grup}, {sira}, {tarih}.'],
 ['OnEk','','Dosya ön eki','Örnek Masraf_'],['SonEk','','Dosya son eki','Örnek _Nihai'],
 ['Bicim','Kaynak','Çıktı biçimi','Kaynak: VBA dahil yerel biçimi korur. xlsx: çıktıdaki VBA kaldırılır.','Kaynak;xlsx'],
 ['Palet','Pastel','Renk paleti','Her grup benzersiz RGB; çok sayıda renkte görsel benzerlik kaçınılmazdır.','Pastel;Canli'],
 ['RenkDosya','Hayir','Renklendirme dosya oluştursun mu?','Hayir: salt okunur açık kaynakta önizleme. Evet: ayrı çıktı dosyası.','Hayir;Evet'],
 ['AnahtarSutunu','A','Mutabakat / mükerrer anahtarı','Tekil kayıt numarası. Kimlikleri baştaki sıfırlar için metin saklayın.'],
 ['TutarSutunu','N','Tutar sütunu','Özet ve mutabakatta zorunlu; kalite kontrolünde boş bırakılabilir.'],
 ['TarihSutunu','O','Tarih sütunu (isteğe bağlı)','Gerçek Excel tarih hücresi; metin tarihler kabul edilmez.'],
 ['IkinciDosya','','Mutabakat ikinci dosyası','İkinci dosya seç düğmesini kullanın.'],
 ['IkinciSayfa','Veri','İkinci dosya veri sayfası','Tam sayfa adı.'],
 ['IkinciBaslik','1','İkinci dosya başlık satırı','Pozitif tam sayı.'],
 ['IkinciAnahtar','A','İkinci dosya anahtar sütunu','Harf veya sayı.'],
 ['IkinciTutar','N','İkinci dosya tutar sütunu','Harf veya sayı.'],
 ['Tolerans','0,01','Mutabakat tutar toleransı','Excel bölgesel sayı biçiminize göre girin; örneğin Türkçe 0,01.'],
 ['ZorunluSutunlar','A;M','Boş alan kontrolü sütunları','Birden çok sütun için noktalı virgül kullanın.'],
 ['BeklenenSutun','','Beklenen değer kontrolü sütunu','Boş bırakılırsa bu kontrol yapılmaz.'],
 ['BeklenenDegerler','','Beklenen değerler','Örnek Garanti;Akbank;Denizbank'],
 ...[1,2,3].flatMap(i=>[
 [`Filtre${i}Sutun`,'',`Filtre ${i} sütunu`,'Boşsa kural kullanılmaz. Kurallar VE ile birleştirilir.'],
 [`Filtre${i}Tur`,'Dahil',`Filtre ${i} türü`,'Dahil/Haric listesi; Tutar/Tarih aralığı.','Dahil;Haric;Tutar;Tarih'],
 [`Filtre${i}Deger`,'',`Filtre ${i} değer / aralık`,'Liste: Garanti;Akbank | Tutar: 100;500 | Tarih: 2026-01-01;2026-12-31']])
];
const guide=[
 ['Tek dosyalık kullanım','Yalnızca ExcelOps.xlsm dosyasını indirin. Windows üzerinde masaüstü Excel 2024 LTSC gerekir.'],
 ['Makroları açma','Dosyayı kapatın. Windows dosya Özellikleri içinde varsa Engellemeyi kaldır seçeneğini uygulayın. Yalnızca güvendiğiniz bu dosyada makrolara izin verin. Kurum ilkeleri engelliyorsa BT yöneticinize başvurun.'],
 ['Menüyü açma','Makrolar etkinse açılışta menü görünür. Yeniden açmak: Alt+F8 > ExcelOpsAc > Çalıştır.'],
 ['İlk işlem','Kaynak seç → Kayıt konumu → Ayarlar. Veri sayfasını, başlık satırını, sütunları kontrol edin. Uygula ayarları bu dosyaya kaydeder.'],
 ['Dosya ayırma','Dosya Ayır: benzersiz grup başına tam kaynak kopyası. Başlık üstü satırlar ve diğer sayfalar korunur. Boş grup anahtarlı satırlar ayrılmış dosyalara alınmaz.'],
 ['Diğer sayfalar','Diğer sayfalarda bütün kaynak kayıtları kalır; çıktıyı paylaşmadan önce bunları ve gizli sayfaları inceleyin. Bu araç yalnız seçilen veri sayfasını ayırır.'],
 ['Dosya biçimi','Kaynak varsayılanı xlsm/xlsb/xls/xlsx biçimini korur. xlsx seçilirse geçici yerel kopya gerçek SaveAs dönüşümünden geçer; VBA içeriği çıktıdan kaldırılır.'],
 ['Renklendirme','Varsayılan olarak dosya oluşturmaz: kaynak salt okunur açılır ve renkler ekranda uygulanır. Diske kaydetmek için Farklı Kaydet kullanın. RenkDosya=Evet ayrı bir çıktı üretir.'],
 ['Renk görünürlüğü','Koşullu biçimlendirmeler normal dolguyu örtebilir. Korunurlar; otomatik kaldırılmazlar. Çok sayıda grubun renkleri benzersiz olsa da gözle ayırt etmek zor olabilir.'],
 ['Özet','Grup, adet, toplam, ortalama, en büyük tutar, ilk/son tarih. Hatalı veya boş anahtar/tutar kabul edilmez; önce Veri Kontrolü çalıştırın.'],
 ['Mutabakat','A kaynak dosya, B ikinci dosya. Anahtar başına adet/tutar karşılaştırılır. Mükerrerler toplanır ve işaretlenir. Olası eşleşmeler yalnızca öneridir; otomatik kabul edilmez.'],
 ['Veri kontrolü','Ayrı kopyada hatalı satırlar kırmızı olur. ExcelOps_Kalite sayfası kaynak satır numarasını ve hataları listeler. Tarihler gerçek Excel seri değeri olmalıdır.'],
 ['Akıllı filtre','En fazla 3 kural VE ile birleştirilir. Dahil/Haric listeleri noktalı virgülle ayrılır. Tutar alt/üst sınırı dahildir. Tarih yyyy-mm-dd;yyyy-mm-dd biçimindedir; son günün tamamı dahildir.'],
 ['Filtre çıktısı','Normal veri aralığında gizli ExcelOps_Filtre yardımcı sütunu oluşturulur ve bu sütun filtrelenir. Önceki filtre/elle gizlenen veri satırları bu işlemde sıfırlanır. Excel tabloları bu işlemde desteklenmez.'],
 ['Veri yapısı','Seçilen başlığın altındaki tüm kullanılan satırlar kayıt sayılır. Ara başlık, dip toplam veya birleşik veri hücreleri içermeyen düz bir liste kullanın. Kaynak dosyayı işlemden önce kapatın.'],
 ['Anahtar eşitliği','Baş/son boşluk ve NBSP temizlenir; büyük/küçük harf farkı yok sayılır. İç boşluklar, * ve ? normal metindir. Kimlikleri metin hücresi olarak saklayın.'],
 ['Çıktı ve hata','Her işlem ayrı tarih-saat klasöründedir; mevcut dosyalar üzerine yazılmaz. Durdur düğmesiyle iptal edebilirsiniz. Tamamlanan çıktılar tutulur; Log sayfası işlem/hata ayrıntılarını içerir.'],
 ['100.000+ satır','Dizi ve Dictionary kullanılır. 1.000 grup, 1.000 tam kitap kopyası demektir; disk/süre kaynağa bağlıdır. Büyük işler öncesi küçük örnekte doğrulayın.'],
 ['Doğrulama durumu','Bu sürümün paket ve gömülü kaynak tutarlılığı Linux üzerinde kontrol edilmiştir. Microsoft Excel 2024 üzerinde çalıştırma, UserForm ve büyük veri performans testleri henüz yapılmamıştır.'],
 ['Bağlantılar / formüller','Formüller ve harici bağlantılar kopyalanır; bağlantılar işlem sırasında yenilenmez. Silinen satırlara bağlı formüllerin sonuçlarını ve #REF! oluşumunu çıktıda kontrol edin. Kaynak VBA kullanıcı fonksiyonları makroları kapalı açılan kopyalarda hesaplanmaz.'],
 ['Desteklenmeyenler','Excel Web/Mac, parola korumalı içerik, imzalı makro imzasını koruma, birleştirilmiş veri hücreleri ve kaynak içindeki otomatik dış veri yenilemeleri bu sürümün doğrulama kapsamı dışındadır.']
];
function style(ws){ws.views=[{state:'frozen',ySplit:1}];ws.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};ws.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF183146'}};ws.getRow(1).height=28;}
async function workbook(){
 const wb=new ExcelJS.Workbook();wb.creator='ExcelOps';wb.created=new Date('2026-10-31T00:00:00Z');wb.modified=wb.created;
 const start=wb.addWorksheet('Başlangıç');start.columns=[{width:28},{width:100}];start.addRow(['EXCELOPS','Excel operasyon platformu · 1.0.0']);style(start);
 const info=[['BAŞLAYIN','Alt+F8 → ExcelOpsAc → Çalıştır'],['01 · Kaynak seç','İşlenecek Excel dosyasını seçin; önce diğer Excel pencerelerinde kapatın.'],['02 · Ayarlar','Veri sayfası, başlık, sütunlar ve çıktı konumunu belirleyin.'],['03 · İşlem','Dosya Ayır / Renklendir / Özet / Mutabakat / Veri Kontrolü / Akıllı Filtre'],['Tek dosya','Bilgisayarınıza yalnızca bu .xlsm dosyasını aktarmanız yeterlidir.'],['Gereksinim','Windows + masaüstü Excel 2024 LTSC ve bu dosya için makro izni.'],['Durum','Excel üzerinde kabul testi bekleyen sürüm. Paket doğrulaması tamamlandı.'],['Yardım','Kılavuz sayfasında adım adım kullanım ve sınırlar açıklanmıştır.'],['Kaynak güvenliği','Kaynak diskte değiştirilmez. Ayırmada diğer sayfalar bütün verileriyle korunur.']];info.forEach(x=>{const r=start.addRow(x);r.height=46;r.alignment={vertical:'middle',wrapText:true};});
 const cfg=wb.addWorksheet('Ayarlar');cfg.addRow(['Anahtar','Değer','Alan','Açıklama','Seçenekler']);for(const row of settings)cfg.addRow(row);cfg.columns=[{width:24},{width:46},{width:36},{width:100},{width:30}];cfg.getColumn(2).numFmt='@';cfg.getColumn(5).hidden=true;style(cfg);for(let i=2;i<=cfg.rowCount;i++){cfg.getRow(i).height=32;cfg.getRow(i).alignment={wrapText:true,vertical:'middle'};}
 const help=wb.addWorksheet('Kılavuz');help.columns=[{width:28},{width:115}];help.addRow(['Konu','Kullanım']);guide.forEach(x=>{const r=help.addRow(x);r.height=60;r.alignment={wrapText:true,vertical:'middle'};});style(help);
 const legal=help.addRow(['Lisans ve atıflar',read('LICENSE').toString('utf8')+'\n\n'+read('THIRD_PARTY_NOTICES.md').toString('utf8')]);legal.height=409;legal.alignment={wrapText:true,vertical:'top'};legal.font={size:8};
 const log=wb.addWorksheet('Log');log.addRow(['Zaman','İşlem','Hata kodu','Açıklama']);log.columns=[{width:24},{width:24},{width:16},{width:110}];log.getColumn(1).numFmt='dd.mm.yyyy hh:mm:ss';style(log);
 return wb;
}
function vba(){
 const projectId='{759788C6-0242-45C8-9B3F-C759F8141A3E}';
 const c=CFB.utils.cfb_new({root:'Root Entry'});const add=(name,data)=>CFB.utils.cfb_add(c,name,Buffer.from(data));
 let project='ID="{759788C6-0242-45C8-9B3F-C759F8141A3E}"\r\n';
 for(const m of modules){project+=(m.type==='document'?`Document=${m.name}/&H00000000`:m.type==='form'?`BaseClass=${m.name}`:m.type==='class'?`Class=${m.name}`:`Module=${m.name}`)+'\r\n';add('VBA/'+m.name,compress(iconv.encode(m.source,'windows1254')));}
 project+='Name="ExcelOps"\r\nHelpContextID="0"\r\nVersionCompatible32="393222000"\r\n\r\n[Host Extender Info]\r\n&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n\r\n[Workspace]\r\n';
 project=project.replace('\r\n\r\n[Host Extender Info]', '\r\nCMG="'+projectData(projectId,Buffer.alloc(4))+'"\r\nDPB="'+projectData(projectId,Buffer.alloc(1))+'"\r\nGC="'+projectData(projectId,Buffer.from([255]))+'"\r\n\r\n[Host Extender Info]');
 for(const m of modules)project+=m.name+'=0, 0, 0, 0, C\r\n';
 add('PROJECT',iconv.encode(project,'windows1254'));
 add('PROJECTwm',Buffer.concat([...modules.flatMap(m=>[Buffer.from(m.name+'\0'),Buffer.from(m.name+'\0','utf16le')]),Buffer.alloc(2)]));
 add('VBA/dir',compress(directory(modules)));add('VBA/_VBA_PROJECT',Buffer.from('cc61ffff000000','hex'));
 for(const m of modules.filter(m=>m.type==='form')){
  add(m.name+'/f',read('tools/assets/empty-form-f.bin'));add(m.name+'/o',Buffer.alloc(0));add(m.name+'/\x01CompObj',read('tools/assets/form-compobj.bin'));
  add(m.name+'/\x03VBFrame',Buffer.from(`VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ${m.name}\r\n   Caption = "ExcelOps"\r\n   ClientHeight = 9000\r\n   ClientWidth = 11000\r\n   StartUpPosition = 1\r\n   TypeInfoVer = 1\r\nEnd\r\n`));
 }
 CFB.utils.cfb_gc(c);return CFB.write(c,{type:'buffer'});
}
async function examples(){
 for(const second of [false,true]){const w=new ExcelJS.Workbook(),s=w.addWorksheet('Veri');s.addRow(['KayitNo','Aciklama','Banka','Avukat','Sube','Urun','Il','Durum','Alan9','Alan10','Alan11','Alan12','Portfoy','Tutar','Tarih']);
 const data=second?[["001","Garanti",100],["002","Akbank",225],["005","Denizbank",90]]:[["001","Garanti",100],["002","Akbank",200],["003","Garanti",150],["004","Denizbank",-20]];
 data.forEach(([id,group,amount],i)=>s.addRow([id,'Örnek kayıt',group,'Avukat A','Merkez','Kredi','İstanbul','Açık','','','','',group,amount,new Date(Date.UTC(2026,9,28+i))]));s.getColumn(1).numFmt='@';s.getColumn(15).numFmt='dd.mm.yyyy';s.columns.forEach(c=>c.width=17);s.autoFilter='A1:O'+s.rowCount;style(s);
 const hidden=w.addWorksheet('GizliNot',{state:'hidden'});hidden.getCell('A1').value='Bu sayfa ayırmada korunmalıdır.';
 const formula=w.addWorksheet('Formul');formula.getCell('A1').value={formula:'SUM(Veri!N2:N5)'};
 await w.xlsx.writeFile(path.join(root,'examples',second?'ReconciliationB.xlsx':'PortfolioSplit.xlsx'));
 }
}
(async()=>{const wb=await workbook();const zip=unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
 let types=strFromU8(zip['[Content_Types].xml']).replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml','application/vnd.ms-excel.sheet.macroEnabled.main+xml').replace('</Types>','<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>');zip['[Content_Types].xml']=strToU8(types);
 let rel=strFromU8(zip['xl/_rels/workbook.xml.rels']).replace('</Relationships>','<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');zip['xl/_rels/workbook.xml.rels']=strToU8(rel);
 let book=strFromU8(zip['xl/workbook.xml']).replace('<workbookPr','<workbookPr codeName="ThisWorkbook"');zip['xl/workbook.xml']=strToU8(book);zip['xl/vbaProject.bin']=vba();
 for(const [i,name] of ['SheetStart','SheetSettings','SheetGuide','SheetLog'].entries()){
  const key='xl/worksheets/sheet'+(i+1)+'.xml';let xml=strFromU8(zip[key]);
  if(xml.includes('<sheetPr'))xml=xml.replace('<sheetPr','<sheetPr codeName="'+name+'"');
  else xml=xml.replace(/(<worksheet[^>]*>)/,'$1<sheetPr codeName="'+name+'"/>');
  zip[key]=strToU8(xml);
 }

 const deterministic={};for(const key of Object.keys(zip).sort())deterministic[key]=[zip[key],{mtime:new Date('2026-10-31T00:00:00Z')}];
 const bytes=zipSync(deterministic,{level:9});fs.writeFileSync(path.join(root,'ExcelOps.xlsm'),bytes);
 await examples();fs.writeFileSync(path.join(root,'ExcelOps.xlsm.sha256'),crypto.createHash('sha256').update(bytes).digest('hex')+'  ExcelOps.xlsm\n');
 console.log('ExcelOps.xlsm:',bytes.length,'bytes;',modules.length,'VBA components; 2 forms.');
})().catch(e=>{console.error(e);process.exitCode=1;});
