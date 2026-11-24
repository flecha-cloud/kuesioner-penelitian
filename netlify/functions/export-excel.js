
const ExcelJS = require('exceljs');
const sharp = require('sharp');

const LIKERT = [
  'Sangat Tidak Setuju',
  'Tidak Setuju',
  'Netral',
  'Setuju',
  'Sangat Setuju'
];

const short = ['STS', 'TS', 'N', 'S', 'SS'];

const groups = [
  {name:'Motivasi Kerja',start:1,end:10},
  {name:'Disiplin Kerja',start:11,end:20},
  {name:'Kinerja Pegawai',start:21,end:30}
];

const questionText = [
'Saya mendapatkan fasilitas kerja yang memadai dari perusahaan.',
'Saya merasa kondisi fisik tempat kerja nyaman untuk bekerja.',
'Saya mendapatkan waktu istirahat yang cukup saat bekerja.',
'Saya merasa aman saat bekerja di lingkungan perusahaan ini.',
'Saya merasa tenang karena perusahaan peduli pada keselamatan kerja.',
'Saya sering diajak berdiskusi atau kerja sama dalam tim.',
'Saya memiliki rekan kerja yang membantu ketika saya butuh bantuan.',
'Saya mendapat dukungan moral dari rekan kerja saat kesulitan.',
'Atasan saya memberikan apresiasi atas pencapaian saya.',
'Saya menerima bonus/insetif atas hasil kerja saya.',
'Saya selalu datang tepat waktu ke tempat kerja.',
'Saya jarang mengambil cuti tanpa alasan penting.',
'Saya selalu mematuhi aturan kerja yang berlaku.',
'Saya tidak menggunakan fasilitas kantor untuk keperluan pribadi.',
'Saya menyelesaikan pekerjaan sesuai dengan tenggat waktu.',
'Saya memastikan hasil pekerjaan sesuai standar perusahaan.',
'Saya bekerja dengan konsisten dari awal hingga akhir kerja.',
'Saya mempertahankan kualitas kerja setiap harinya.',
'Saya menjadikan teguran sebagai bahan perbaikan diri.',
'Saya tidak merasa tersinggung jika ditegur karena kesalahan.',
'Saya berusaha meningkatkan mutu kerja saya secara berkelanjutan.',
'Saya memastikan hasil kerja saya rapi dan sesuai standar.',
'Saya memenuhi target kerja yang diberikan atasan.',
'Saya tidak pernah menunda-nunda pekerjaan.',
'Saya merasa bertanggung jawab atas pekerjaan yang saya kerjakan.',
'Saya bersedia memperbaiki kesalahan dalam pekerjaan saya.',
'Saya aktif berkontribusi dalam kegiatan kelompok.',
'Saya bersedia membantu rekan kerja saat mereka kesulitan.',
'Saya mengambil tindakan lebih tanpa harus diperintah.',
'Saya sering mengusulkan ide-ide baru dalam pekerjaan.'
];

function esc(s) {
  return String(s).replace(/[&<>"']/g,c=>({
    '&':'&amp;',
    '<':'&lt;',
    '>':'&gt;',
    '"':'&quot;',
    "'":'&apos;'
  }[c]));
}

// MEMBUAT DIAGRAM LINGKARAN
async function chartPng(title, counts) {
  const W=1000,H=650,cx=330,cy=315,r=205;
  const colors=[
    '#8b5bb8',
    '#a47acb',
    '#c3a8e4',
    '#7542ae',
    '#4b1b85'
  ];

  const total=counts.reduce((a,b)=>a+b,0);
  let angle=-Math.PI/2;
  let slices='';

  if(total===0) {
    slices=`<circle cx="${cx}" cy="${cy}" r="${r}" fill="#eee8f5"/>`;
  } else {
    for(let i=0;i<counts.length;i++) {
      if(!counts[i]) continue;

      const sweep=counts[i]/total*Math.PI*2;
      const x1=cx+r*Math.cos(angle);
      const y1=cy+r*Math.sin(angle);
      const end=angle+sweep;
      const x2=cx+r*Math.cos(end);
      const y2=cy+r*Math.sin(end);

      const d=`M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${sweep>Math.PI?1:0} 1 ${x2} ${y2} Z`;

      slices+=`<path d="${d}" fill="${colors[i]}" stroke="#ffffff" stroke-width="4"/>`;
      angle=end;
    }
  }

  const legend=LIKERT.map((name,i)=>{
    const pct=total?(counts[i]/total*100).toFixed(1):'0.0';
    const y=145+i*88;

    return `
      <rect x="650" y="${y-22}" width="28" height="28" rx="6" fill="${colors[i]}"/>
      <text x="700" y="${y}" font-family="Arial" font-size="25" fill="#3a2751">${short[i]} - ${esc(name)}</text>
      <text x="700" y="${y+32}" font-family="Arial" font-size="22" fill="#786c83">${counts[i]} jawaban (${pct}%)</text>
    `;
  }).join('');

  const svg=`
  <svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <rect width="100%" height="100%" fill="#ffffff"/>
    <text x="${W/2}" y="55" text-anchor="middle" font-family="Arial" font-size="31" font-weight="700" fill="#35136c">${esc(title)}</text>
    ${slices}
    <text x="${cx}" y="${cy-8}" text-anchor="middle" font-family="Arial" font-size="34" font-weight="700" fill="#35136c">${total}</text>
    <text x="${cx}" y="${cy+30}" text-anchor="middle" font-family="Arial" font-size="21" fill="#786c83">jawaban</text>
    ${legend}
  </svg>`;

  return sharp(Buffer.from(svg)).png().toBuffer();
}

// PROSES EKSPOR EXCEL
exports.handler = async (event) => {

  if(event.httpMethod!=='GET') {
    return {
      statusCode:405,
      headers:{Allow:'GET'},
      body:'Method not allowed'
    };
  }

  const password=event.headers['x-export-password'] || event.queryStringParameters?.key;

  if(!process.env.EXPORT_PASSWORD || password!==process.env.EXPORT_PASSWORD) {
    return {
      statusCode:401,
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({error:'Kata sandi ekspor tidak valid.'})
    };
  }

  const token=process.env.NETLIFY_AUTH_TOKEN;
  const siteId=process.env.SITE_ID || process.env.NETLIFY_SITE_ID;

  if(!token || !siteId) {
    return {
      statusCode:500,
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({error:'Konfigurasi NETLIFY_AUTH_TOKEN atau NETLIFY_SITE_ID belum diisi di Netlify.'})
    };
  }

  try {
    const headers={
      Authorization:`Bearer ${token}`,
      'Content-Type':'application/json'
    };

    const formsRes=await fetch(
      `https://api.netlify.com/api/v1/sites/${siteId}/forms`,
      {headers}
    );

    if(!formsRes.ok) {
      throw new Error(`Gagal membaca daftar form Netlify (${formsRes.status}).`);
    }

    const forms=await formsRes.json();
    const form=forms.find(f=>f.name==='survey-responses');

    if(!form) {
      throw new Error('Form survey-responses belum ditemukan. Pastikan website sudah di-deploy dan minimal satu pengiriman uji berhasil.');
    }

    let submissions=[],page=1;

    while(true) {
      const r=await fetch(
        `https://api.netlify.com/api/v1/forms/${form.id}/submissions?per_page=100&page=${page}`,
        {headers}
      );

      if(!r.ok) {
        throw new Error(`Gagal membaca data responden (${r.status}).`);
      }

      const chunk=await r.json();

      if(!Array.isArray(chunk) || chunk.length===0) break;

      submissions.push(...chunk);

      if(chunk.length<100) break;

      page++;
      if(page>100) break;
    }

    submissions.sort((a,b)=>
      new Date(a.created_at||0)-new Date(b.created_at||0)
    );

    const wb=new ExcelJS.Workbook();

    wb.creator='Kuesioner Penelitian';
    wb.created=new Date();
    wb.modified=new Date();
    wb.subject='Data responden penelitian';
    wb.company='Kuesioner Penelitian';

    // SHEET 1: DATA RESPONDEN
    const data=wb.addWorksheet('Data Responden');

    data.views=[{state:'frozen',ySplit:1,xSplit:4}];

    const headersData=[
      'No','Waktu Pengisian','Gender','Usia','Pendidikan',
      ...Array.from({length:30},(_,i)=>`P${i+1}`)
    ];

    data.addRow(headersData);

    submissions.forEach((s,i)=>{
      const d=s.data||{};

      const row=[
        i+1,
        s.created_at||d.submitted_at||'',
        d.gender||'',
        d.age||'',
        d.education||'',
        ...Array.from({length:30},(_,j)=>d[`P${j+1}`]||'')
      ];

      data.addRow(row);
    });

    data.getRow(1).height=30;

    data.getRow(1).eachCell(c=>{
      c.font={bold:true,color:{argb:'FFFFFFFF'}};
      c.fill={
        type:'pattern',
        pattern:'solid',
        fgColor:{argb:'FF55239B'}
      };
      c.alignment={
        vertical:'middle',
        horizontal:'center',
        wrapText:true
      };
    });

    data.columns=[
      {width:7},{width:23},{width:16},{width:22},{width:22},
      ...Array(30).fill({width:24})
    ];

    data.autoFilter={from:'A1',to:'AI1'};

    data.eachRow((row,r)=>{
      if(r>1) {
        row.alignment={vertical:'middle',wrapText:true};

        if(r%2===0) {
          row.eachCell(c=>{
            c.fill={
              type:'pattern',
              pattern:'solid',
              fgColor:{argb:'FFF7F3FC'}
            };
          });
        }
      }
    });

    // SHEET 2: REKAPITULASI
    const recap=wb.addWorksheet('Rekapitulasi');

    recap.columns=[
      {width:23},{width:10},{width:75},
      {width:13},{width:13},{width:13},
      {width:13},{width:13},{width:14}
    ];

    recap.addRow([
      'Variabel','Kode','Pernyataan',
      'STS','TS','N','S','SS','Total'
    ]);

    recap.getRow(1).height=28;

    recap.getRow(1).eachCell(c=>{
      c.font={bold:true,color:{argb:'FFFFFFFF'}};
      c.fill={
        type:'pattern',
        pattern:'solid',
        fgColor:{argb:'FF55239B'}
      };
      c.alignment={
        vertical:'middle',
        horizontal:'center',
        wrapText:true
      };
    });

    const allCounts={};

    for(const g of groups) {
      const summary=[0,0,0,0,0];

      for(let q=g.start;q<=g.end;q++) {
        const counts=LIKERT.map(v=>
          submissions.reduce(
            (n,s)=>n+((s.data||{})[`P${q}`]===v?1:0),
            0
          )
        );

        counts.forEach((v,i)=>summary[i]+=v);
        allCounts[q]=counts;

        recap.addRow([
          g.name,
          `P${q}`,
          questionText[q-1],
          ...counts,
          counts.reduce((a,b)=>a+b,0)
        ]);
      }

      recap.addRow([
        `TOTAL ${g.name}`,'','',
        ...summary,
        summary.reduce((a,b)=>a+b,0)
      ]);

      recap.addRow([]);
    }

    recap.eachRow((row,r)=>{
      if(r>1) {
        row.alignment={vertical:'middle',wrapText:true};

        if(String(row.getCell(1).value||'').startsWith('TOTAL')) {
          row.eachCell(c=>{
            c.font={bold:true,color:{argb:'FF35136C'}};
            c.fill={
              type:'pattern',
              pattern:'solid',
              fgColor:{argb:'FFEDE4F8'}
            };
          });
        }
      }
    });

    // SHEET 3: DIAGRAM
    const diagram=wb.addWorksheet('Diagram');

    diagram.getCell('A1').value='DISTRIBUSI JAWABAN RESPONDEN';
    diagram.getCell('A1').font={
      bold:true,size:16,color:{argb:'FF35136C'}
    };

    diagram.getCell('A2').value=`Jumlah responden: ${submissions.length}`;
    diagram.getCell('A2').font={
      italic:true,color:{argb:'FF756B83'}
    };

    diagram.columns=[
      {width:25},{width:18},{width:18},{width:18},
      {width:18},{width:18},{width:18},{width:18}
    ];

    let row=4;

    for(const g of groups) {

      const sums=[0,0,0,0,0];

      for(let q=g.start;q<=g.end;q++) {
        allCounts[q].forEach((v,i)=>sums[i]+=v);
      }

      const total=sums.reduce((a,b)=>a+b,0);

      diagram.getCell(`A${row}`).value=g.name;
      diagram.getCell(`A${row}`).font={
        bold:true,size:14,color:{argb:'FF35136C'}
      };

      const img=await chartPng(g.name,sums);

      const imageId=wb.addImage({
        buffer:img,
        extension:'png'
      });

      diagram.addImage(imageId,{
        tl:{col:0,row:row},
        ext:{width:760,height:494}
      });

      const tableRow=row+26;

      diagram.getRow(tableRow).values=[
        'Kategori','Jumlah','Persentase'
      ];

      diagram.getRow(tableRow).eachCell(c=>{
        c.font={bold:true,color:{argb:'FFFFFFFF'}};
        c.fill={
          type:'pattern',
          pattern:'solid',
          fgColor:{argb:'FF7542AE'}
        };
      });

      for(let i=0;i<5;i++) {
        diagram.getRow(tableRow+1+i).values=[
          LIKERT[i],
          sums[i],
          total?`${(sums[i]/total*100).toFixed(1)}%`:'0.0%'
        ];
      }

      const maxIndex=sums.indexOf(Math.max(...sums));

      const maxPct=total
        ?(sums[maxIndex]/total*100).toFixed(1)
        :'0.0';

      const description=total===0
        ?`Belum terdapat jawaban yang dapat direkap untuk variabel ${g.name}.`
        :`Kategori jawaban terbanyak adalah "${LIKERT[maxIndex]}" (${sums[maxIndex]} dari ${total} jawaban; ${maxPct}%). Hasil ini menggambarkan kecenderungan jawaban responden pada variabel ${g.name}, berdasarkan seluruh butir pertanyaan P${g.start}–P${g.end}.`;

      const descRow=tableRow+8;

      diagram.getCell(`A${descRow}`).value='Deskripsi hasil';
      diagram.getCell(`A${descRow}`).font={
        bold:true,color:{argb:'FF35136C'}
      };

      diagram.mergeCells(`A${descRow+1}:H${descRow+3}`);

      diagram.getCell(`A${descRow+1}`).value=description;

      diagram.getCell(`A${descRow+1}`).alignment={
        wrapText:true,
        vertical:'top'
      };

      diagram.getCell(`A${descRow+1}`).fill={
        type:'pattern',
        pattern:'solid',
        fgColor:{argb:'FFF7F3FC'}
      };

      diagram.getRow(descRow+1).height=30;
      diagram.getRow(descRow+2).height=25;
      diagram.getRow(descRow+3).height=25;

      row=descRow+7;
    }

    // DOWNLOAD FILE EXCEL
    const buffer=await wb.xlsx.writeBuffer();

    const stamp=new Date().toISOString().slice(0,10);

    return {
      statusCode:200,
      isBase64Encoded:true,
      headers:{
        'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition':`attachment; filename="hasil-kuesioner-${stamp}.xlsx"`,
        'Cache-Control':'no-store'
      },
      body:Buffer.from(buffer).toString('base64')
    };

  } catch(err) {
    console.error(err);

    return {
      statusCode:500,
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        error:err.message||'Terjadi kesalahan saat membuat Excel.'
      })
    };
  }
};
      
