const ExcelJS = require('exceljs');
const sharp = require('sharp');
const LIKERT = ['Sangat Tidak Setuju','Tidak Setuju','Netral','Setuju','Sangat Setuju'];
const short = ['STS','TS','N','S','SS'];
const groups = [
  {name:'Motivasi Kerja',start:1,end:10},
  {name:'Disiplin Kerja',start:11,end:20},
  {name:'Kinerja Pegawai',start:21,end:30}
];
const questionText = [
'Saya mendapatkan fasilitas kerja yang memadai dari perusahaan.','Saya merasa kondisi fisik tempat kerja nyaman untuk bekerja.','Saya mendapatkan waktu istirahat yang cukup saat bekerja.','Saya merasa aman saat bekerja di lingkungan perusahaan ini.','Saya merasa tenang karena perusahaan peduli pada keselamatan kerja.','Saya sering diajak berdiskusi atau kerja sama dalam tim.','Saya memiliki rekan kerja yang membantu ketika saya butuh bantuan.','Saya mendapat dukungan moral dari rekan kerja saat kesulitan.','Atasan saya memberikan apresiasi atas pencapaian saya.','Saya menerima bonus/insetif atas hasil kerja saya.',
'Saya selalu datang tepat waktu ke tempat kerja.','Saya jarang mengambil cuti tanpa alasan penting.','Saya selalu mematuhi aturan kerja yang berlaku.','Saya tidak menggunakan fasilitas kantor untuk keperluan pribadi.','Saya menyelesaikan pekerjaan sesuai dengan tenggat waktu.','Saya memastikan hasil pekerjaan sesuai standar perusahaan.','Saya bekerja dengan konsisten dari awal hingga akhir kerja.','Saya mempertahankan kualitas kerja setiap harinya.','Saya menjadikan teguran sebagai bahan perbaikan diri.','Saya tidak merasa tersinggung jika ditegur karena kesalahan.',
'Saya berusaha meningkatkan mutu kerja saya secara berkelanjutan.','Saya memastikan hasil kerja saya rapi dan sesuai standar.','Saya memenuhi target kerja yang diberikan atasan.','Saya tidak pernah menunda-nunda pekerjaan.','Saya merasa bertanggung jawab atas pekerjaan yang saya kerjakan.','Saya bersedia memperbaiki kesalahan dalam pekerjaan saya.','Saya aktif berkontribusi dalam kegiatan kelompok.','Saya bersedia membantu rekan kerja saat mereka kesulitan.','Saya mengambil tindakan lebih tanpa harus diperintah.','Saya sering mengusulkan ide-ide baru dalam pekerjaan.'
];
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));}
async function chartPng(title, counts){
  const W=1000,H=470,left=95,right=45,top=85,bottom=95,plotH=H-top-bottom,plotW=W-left-right,max=Math.max(1,...counts);
  const colors=['#8b5bb8','#a47acb','#c3a8e4','#7542ae','#4b1b85'];
  const barW=105,gap=(plotW-5*barW)/6;let bars='';
  for(let i=0;i<5;i++){const h=counts[i]/max*(plotH-35),x=left+gap+(barW+gap)*i,y=top+plotH-h;bars+=`<rect x="${x}" y="${y}" width="${barW}" height="${h}" rx="12" fill="${colors[i]}"/><text x="${x+barW/2}" y="${y-14}" font-size="22" font-weight="700" text-anchor="middle" fill="#3a2751">${counts[i]}</text><text x="${x+barW/2}" y="${top+plotH+38}" font-size="23" font-weight="700" text-anchor="middle" fill="#4b3c5e">${short[i]}</text>`;}
  let grid='';for(let j=0;j<=4;j++){const y=top+plotH-j*plotH/4;grid+=`<line x1="${left}" y1="${y}" x2="${W-right}" y2="${y}" stroke="#e9e1f0" stroke-width="2"/><text x="${left-18}" y="${y+7}" font-size="17" text-anchor="end" fill="#857991">${Math.round(max*j/4)}</text>`;}
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><rect width="100%" height="100%" fill="#ffffff"/><text x="${W/2}" y="45" text-anchor="middle" font-family="Arial" font-size="28" font-weight="700" fill="#35136c">${esc(title)}</text>${grid}${bars}<text x="${W/2}" y="${H-18}" text-anchor="middle" font-family="Arial" font-size="17" fill="#786c83">Distribusi jawaban (jumlah responden)</text></svg>`;
  return sharp(Buffer.from(svg)).png().toBuffer();
}
exports.handler = async (event) => {
  if(event.httpMethod!=='GET') return {statusCode:405,headers:{Allow:'GET'},body:'Method not allowed'};
  const password=event.headers['x-export-password'] || event.queryStringParameters?.key;
  if(!process.env.EXPORT_PASSWORD || password!==process.env.EXPORT_PASSWORD) return {statusCode:401,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:'Kata sandi ekspor tidak valid.'})};
  const token=process.env.NETLIFY_AUTH_TOKEN,siteId=process.env.NETLIFY_SITE_ID;
  if(!token||!siteId)return {statusCode:500,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:'Konfigurasi NETLIFY_AUTH_TOKEN atau NETLIFY_SITE_ID belum diisi di Netlify.'})};
  try{
    const headers={Authorization:`Bearer ${token}`,'Content-Type':'application/json'};
    const formsRes=await fetch(`https://api.netlify.com/api/v1/sites/${siteId}/forms`,{headers});
    if(!formsRes.ok)throw new Error(`Gagal membaca daftar form Netlify (${formsRes.status}).`);
    const forms=await formsRes.json();const form=forms.find(f=>f.name==='survey-responses');
    if(!form)throw new Error('Form survey-responses belum ditemukan. Pastikan website sudah di-deploy dan minimal satu pengiriman uji berhasil.');
    let submissions=[],page=1;
    while(true){const r=await fetch(`https://api.netlify.com/api/v1/forms/${form.id}/submissions?per_page=100&page=${page}`,{headers});if(!r.ok)throw new Error(`Gagal membaca data responden (${r.status}).`);const chunk=await r.json();if(!Array.isArray(chunk)||chunk.length===0)break;submissions.push(...chunk);if(chunk.length<100)break;page++;if(page>100)break;}
    submissions.sort((a,b)=>new Date(a.created_at||0)-new Date(b.created_at||0));
    const wb=new ExcelJS.Workbook();wb.creator='Kuesioner Penelitian';wb.created=new Date();wb.modified=new Date();wb.subject='Data responden penelitian';wb.company='Kuesioner Penelitian';
    const data=wb.addWorksheet('Data Responden');data.views=[{state:'frozen',ySplit:1,xSplit:4}];
    const headersData=['No','Waktu Pengisian','Gender','Usia','Pendidikan',...Array.from({length:30},(_,i)=>`P${i+1}`)];data.addRow(headersData);
    submissions.forEach((s,i)=>{const d=s.data||{};const row=[i+1,s.created_at||d.submitted_at||'',d.gender||'',d.age||'',d.education||'',...Array.from({length:30},(_,j)=>d[`P${j+1}`]||'')];data.addRow(row)});
    data.getRow(1).height=30;data.getRow(1).eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF55239B'}};c.alignment={vertical:'middle',horizontal:'center',wrapText:true}});
    data.columns=[{width:7},{width:23},{width:16},{width:22},{width:22},...Array(30).fill({width:24})];data.autoFilter={from:'A1',to:'AI1'};
    data.eachRow((row,r)=>{if(r>1){row.alignment={vertical:'middle',wrapText:true};if(r%2===0)row.eachCell(c=>c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF7F3FC'}})}});
    const recap=wb.addWorksheet('Rekapitulasi');recap.columns=[{width:23},{width:10},{width:75},{width:13},{width:13},{width:13},{width:13},{width:13},{width:14}];recap.addRow(['Variabel','Kode','Pernyataan','STS','TS','N','S','SS','Total']);
    recap.getRow(1).height=28;recap.getRow(1).eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF55239B'}};c.alignment={vertical:'middle',horizontal:'center',wrapText:true}});
    const allCounts={};
    for(const g of groups){const summary=[0,0,0,0,0];for(let q=g.start;q<=g.end;q++){const counts=LIKERT.map(v=>submissions.reduce((n,s)=>n+((s.data||{})[`P${q}`]===v?1:0),0));counts.forEach((v,i)=>summary[i]+=v);allCounts[q]=counts;recap.addRow([g.name,`P${q}`,questionText[q-1],...counts,counts.reduce((a,b)=>a+b,0)]);}recap.addRow([`TOTAL ${g.name}`,'','',...summary,summary.reduce((a,b)=>a+b,0)]);recap.addRow([]);}
    recap.eachRow((row,r)=>{if(r>1){row.alignment={vertical:'middle',wrapText:true};if(String(row.getCell(1).value||'').startsWith('TOTAL'))row.eachCell(c=>{c.font={bold:true,color:{argb:'FF35136C'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEDE4F8'}}})}});
    const diagram=wb.addWorksheet('Diagram');diagram.getCell('A1').value='DISTRIBUSI JAWABAN RESPONDEN';diagram.getCell('A1').font={bold:true,size:16,color:{argb:'FF35136C'}};diagram.getCell('A2').value=`Jumlah responden: ${submissions.length}`;diagram.getCell('A2').font={italic:true,color:{argb:'FF756B83'}};diagram.getColumn(1).width=22;diagram.getColumn(2).width=18;diagram.getColumn(3).width=18;diagram.getColumn(4).width=18;diagram.getColumn(5).width=18;diagram.getColumn(6).width=18;
    let row=4;for(const g of groups){const sums=[0,0,0,0,0];for(let q=g.start;q<=g.end;q++)allCounts[q].forEach((v,i)=>sums[i]+=v);diagram.getCell(`A${row}`).value=g.name;diagram.getCell(`A${row}`).font={bold:true,size:13,color:{argb:'FF35136C'}};diagram.addRow([]);diagram.addRow(['Kategori','STS','TS','N','S','SS']);const header=diagram.getRow(row+2);header.eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF7542AE'}}});diagram.addRow(['Jumlah',...sums]);const img=await chartPng(g.name,sums);const imageId=wb.addImage({buffer:img,extension:'png'});diagram.addImage(imageId,{tl:{col:0,row:row+3},ext:{width:660,height:310}});row+=23;}
    const buffer=await wb.xlsx.writeBuffer();const stamp=new Date().toISOString().slice(0,10);return {statusCode:200,isBase64Encoded:true,headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':`attachment; filename="hasil-kuesioner-${stamp}.xlsx"`,'Cache-Control':'no-store'},body:Buffer.from(buffer).toString('base64')};
  }catch(err){console.error(err);return {statusCode:500,headers:{'Content-Type':'application/json'},body:JSON.stringify({error:err.message||'Terjadi kesalahan saat membuat Excel.'})};}
};
