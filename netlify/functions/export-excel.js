
const ExcelJS = require('exceljs');
const sharp = require('sharp');

const LIKERT = [
  'Sangat Tidak Setuju',
  'Tidak Setuju',
  'Netral',
  'Setuju',
  'Sangat Setuju'
];

const SHORT = ['STS', 'TS', 'N', 'S', 'SS'];

const COLORS = [
  '#8b5bb8',
  '#a47acb',
  '#c3a8e4',
  '#7542ae',
  '#4b1b85'
];

const groups = [
  { name: 'Motivasi Kerja', start: 1, end: 10 },
  { name: 'Disiplin Kerja', start: 11, end: 20 },
  { name: 'Kinerja Pegawai', start: 21, end: 30 }
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

// MEMBUAT DIAGRAM LINGKARAN TANPA TEKS
async function chartPng(counts) {
  const W = 700;
  const H = 500;
  const cx = 350;
  const cy = 250;
  const radius = 190;

  const total = counts.reduce((a, b) => a + b, 0);

  let angle = -Math.PI / 2;
  let slices = '';

  if (total === 0) {
    slices = `<circle cx="${cx}" cy="${cy}" r="${radius}" fill="#eee8f5"/>`;
  } else {
    for (let i = 0; i < counts.length; i++) {
      if (counts[i] === 0) continue;

      const sweep = counts[i] / total * Math.PI * 2;

      const x1 = cx + radius * Math.cos(angle);
      const y1 = cy + radius * Math.sin(angle);

      const end = angle + sweep;

      const x2 = cx + radius * Math.cos(end);
      const y2 = cy + radius * Math.sin(end);

      const large = sweep > Math.PI ? 1 : 0;

      const path = `
        M ${cx} ${cy}
        L ${x1} ${y1}
        A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2}
        Z
      `;

      slices += `
        <path
          d="${path}"
          fill="${COLORS[i]}"
          stroke="#ffffff"
          stroke-width="5"
        />
      `;

      angle = end;
    }
  }

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg"
      width="${W}" height="${H}">
      <rect width="100%" height="100%" fill="#ffffff"/>
      ${slices}
    </svg>
  `;

  return sharp(Buffer.from(svg)).png().toBuffer();
}

// MEMBUAT KESIMPULAN DESKRIPTIF
function makeConclusion(groupName, counts, total) {
  if (total === 0) {
    return `Belum terdapat jawaban responden yang dapat dianalisis untuk variabel ${groupName}.`;
  }

  const negative = counts[0] + counts[1];
  const neutral = counts[2];
  const positive = counts[3] + counts[4];

  const negativePct = negative / total * 100;
  const neutralPct = neutral / total * 100;
  const positivePct = positive / total * 100;

  let result = '';

  if (positive > negative && positive > neutral) {
    result =
      `Berdasarkan hasil rekapitulasi, jawaban Setuju dan Sangat Setuju ` +
      `mendominasi dengan persentase ${positivePct.toFixed(1)}% dari seluruh jawaban. ` +
      `Sementara itu, jawaban Netral sebesar ${neutralPct.toFixed(1)}% dan ` +
      `jawaban negatif sebesar ${negativePct.toFixed(1)}%. ` +
      `Hal ini menunjukkan bahwa secara deskriptif tanggapan responden ` +
      `terhadap variabel ${groupName} cenderung positif.`;
  } else if (negative > positive && negative > neutral) {
    result =
      `Berdasarkan hasil rekapitulasi, jawaban Sangat Tidak Setuju dan Tidak Setuju ` +
      `mendominasi dengan persentase ${negativePct.toFixed(1)}% dari seluruh jawaban. ` +
      `Jawaban Netral sebesar ${neutralPct.toFixed(1)}% dan jawaban positif ` +
      `sebesar ${positivePct.toFixed(1)}%. ` +
      `Hal ini menunjukkan bahwa secara deskriptif tanggapan responden ` +
      `terhadap variabel ${groupName} cenderung negatif dan aspek-aspek ` +
      `yang dinilai perlu mendapat perhatian lebih lanjut.`;
  } else {
    result =
      `Berdasarkan hasil rekapitulasi, jawaban responden menunjukkan ` +
      `kecenderungan yang relatif beragam. Jawaban positif sebesar ` +
      `${positivePct.toFixed(1)}%, jawaban Netral sebesar ${neutralPct.toFixed(1)}%, ` +
      `dan jawaban negatif sebesar ${negativePct.toFixed(1)}%. ` +
      `Dengan demikian, secara deskriptif belum terdapat dominasi ` +
      `kecenderungan jawaban yang kuat pada variabel ${groupName}.`;
  }

  return result;
}

// PROSES UTAMA EKSPOR EXCEL
exports.handler = async (event) => {

  if (event.httpMethod !== 'GET') {
    return {
      statusCode: 405,
      headers: { Allow: 'GET' },
      body: 'Method not allowed'
    };
  }

  // PEMERIKSAAN PASSWORD
  const password =
    event.headers['x-export-password'] ||
    event.queryStringParameters?.key;

  if (
    !process.env.EXPORT_PASSWORD ||
    password !== process.env.EXPORT_PASSWORD
  ) {
    return {
      statusCode: 401,
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        error: 'Kata sandi ekspor tidak valid.'
      })
    };
  }

  // KONFIGURASI NETLIFY
  const token = process.env.NETLIFY_AUTH_TOKEN;

  const siteId =
    process.env.NETLIFY_SITE_ID ||
    process.env.SITE_ID;

  if (!token || !siteId) {
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        error: 'Konfigurasi NETLIFY_AUTH_TOKEN atau NETLIFY_SITE_ID belum diisi di Netlify.'
      })
    };
  }

  try {

    const headers = {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    };

    // MENGAMBIL DAFTAR FORM
    const formsRes = await fetch(
      `https://api.netlify.com/api/v1/sites/${siteId}/forms`,
      { headers }
    );

    if (!formsRes.ok) {
      throw new Error(
        `Gagal membaca daftar form Netlify (${formsRes.status}).`
      );
    }

    const forms = await formsRes.json();

    const form = forms.find(
      f => f.name === 'survey-responses'
    );

    if (!form) {
      throw new Error(
        'Form survey-responses belum ditemukan. Pastikan website sudah di-deploy dan minimal satu pengiriman uji berhasil.'
      );
    }

    // MENGAMBIL SELURUH DATA RESPONDEN
    let submissions = [];
    let page = 1;

    while (true) {
      const response = await fetch(
        `https://api.netlify.com/api/v1/forms/${form.id}/submissions?per_page=100&page=${page}`,
        { headers }
      );

      if (!response.ok) {
        throw new Error(
          `Gagal membaca data responden (${response.status}).`
        );
      }

      const chunk = await response.json();

      if (!Array.isArray(chunk) || chunk.length === 0) {
        break;
      }

      submissions.push(...chunk);

      if (chunk.length < 100) {
        break;
      }

      page++;

      if (page > 100) {
        break;
      }
    }

    submissions.sort((a, b) =>
      new Date(a.created_at || 0) -
      new Date(b.created_at || 0)
    );

    // MEMBUAT WORKBOOK
    const wb = new ExcelJS.Workbook();

    wb.creator = 'Kuesioner Penelitian';
    wb.created = new Date();
    wb.modified = new Date();
    wb.subject = 'Data responden penelitian';
    wb.company = 'Kuesioner Penelitian';

    // =====================================
    // SHEET 1: DATA RESPONDEN
    // =====================================

    const data = wb.addWorksheet('Data Responden');

    data.views = [
      { state: 'frozen', ySplit: 1, xSplit: 4 }
    ];

    const headersData = [
      'No',
      'Waktu Pengisian',
      'Gender',
      'Usia',
      'Pendidikan',
      ...Array.from({ length: 30 }, (_, i) => `P${i + 1}`)
    ];

    data.addRow(headersData);

    submissions.forEach((submission, i) => {
      const d = submission.data || {};

      const row = [
        i + 1,
        submission.created_at || d.submitted_at || '',
        d.gender || '',
        d.age || '',
        d.education || '',
        ...Array.from(
          { length: 30 },
          (_, j) => d[`P${j + 1}`] || ''
        )
      ];

      data.addRow(row);
    });

    data.getRow(1).height = 30;

    data.getRow(1).eachCell(cell => {
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' }
      };

      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF55239B' }
      };

      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true
      };
    });

    data.columns = [
      { width: 7 },
      { width: 23 },
      { width: 16 },
      { width: 22 },
      { width: 22 },
      ...Array(30).fill({ width: 24 })
    ];

    data.autoFilter = {
      from: 'A1',
      to: 'AI1'
    };

    data.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.alignment = {
          vertical: 'middle',
          wrapText: true
        };

        if (rowNumber % 2 === 0) {
          row.eachCell(cell => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFF7F3FC' }
            };
          });
        }
      }
    });

    // =====================================
    // SHEET 2: REKAPITULASI
    // =====================================

    const recap = wb.addWorksheet('Rekapitulasi');

    recap.columns = [
      { width: 23 },
      { width: 10 },
      { width: 75 },
      { width: 13 },
      { width: 13 },
      { width: 13 },
      { width: 13 },
      { width: 13 },
      { width: 14 }
    ];

    recap.addRow([
      'Variabel',
      'Kode',
      'Pernyataan',
      'STS',
      'TS',
      'N',
      'S',
      'SS',
      'Total'
    ]);

    recap.getRow(1).height = 28;

    recap.getRow(1).eachCell(cell => {
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' }
      };

      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF55239B' }
      };

      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true
      };
    });

    const allCounts = {};

    for (const group of groups) {

      const summary = [0, 0, 0, 0, 0];

      for (
        let question = group.start;
        question <= group.end;
        question++
      ) {

        const counts = LIKERT.map(value =>
          submissions.reduce((number, submission) => {
            const d = submission.data || {};

            return number +
              (d[`P${question}`] === value ? 1 : 0);
          }, 0)
        );

        counts.forEach((value, index) => {
          summary[index] += value;
        });

        allCounts[question] = counts;

        recap.addRow([
          group.name,
          `P${question}`,
          questionText[question - 1],
          ...counts,
          counts.reduce((a, b) => a + b, 0)
        ]);
      }

      recap.addRow([
        `TOTAL ${group.name}`,
        '',
        '',
        ...summary,
        summary.reduce((a, b) => a + b, 0)
      ]);

      recap.addRow([]);
    }

    recap.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.alignment = {
          vertical: 'middle',
          wrapText: true
        };

        if (
          String(row.getCell(1).value || '')
            .startsWith('TOTAL')
        ) {
          row.eachCell(cell => {
            cell.font = {
              bold: true,
              color: { argb: 'FF35136C' }
            };

            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFEDE4F8' }
            };
          });
        }
      }
    });

    // =====================================
    // SHEET 3: DIAGRAM DAN KESIMPULAN
    // =====================================

    const diagram = wb.addWorksheet('Diagram');

    diagram.getCell('A1').value =
      'HASIL ANALISIS DESKRIPTIF PENELITIAN';

    diagram.getCell('A1').font = {
      bold: true,
      size: 16,
      color: { argb: 'FF35136C' }
    };

    diagram.getCell('A2').value =
      `Jumlah responden: ${submissions.length}`;

    diagram.getCell('A2').font = {
      italic: true,
      color: { argb: 'FF756B83' }
    };

    diagram.columns = [
      { width: 27 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 16 }
    ];

    let row = 4;

    for (const group of groups) {

      const sums = [0, 0, 0, 0, 0];

      for (
        let question = group.start;
        question <= group.end;
        question++
      ) {
        const counts = allCounts[question] || [0, 0, 0, 0, 0];

        counts.forEach((value, index) => {
          sums[index] += value;
        });
      }

      const total = sums.reduce((a, b) => a + b, 0);

      // JUDUL VARIABEL
      diagram.mergeCells(`A${row}:H${row}`);

      diagram.getCell(`A${row}`).value =
        group.name.toUpperCase();

      diagram.getCell(`A${row}`).font = {
        bold: true,
        size: 14,
        color: { argb: 'FF35136C' }
      };

      // GAMBAR DIAGRAM LINGKARAN
      const imageBuffer = await chartPng(sums);

      const imageId = wb.addImage({
        buffer: imageBuffer,
        extension: 'png'
      });

      diagram.addImage(imageId, {
        tl: { col: 1, row: row + 1 },
        ext: { width: 470, height: 335 }
      });

      // TABEL PERSENTASE
      const tableRow = row + 19;

      diagram.getRow(tableRow).values = [
        'Kategori Jawaban',
        'Kode',
        'Jumlah',
        'Persentase'
      ];

      diagram.getRow(tableRow).eachCell(cell => {
        cell.font = {
          bold: true,
          color: { argb: 'FFFFFFFF' }
        };

        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF7542AE' }
        };
      });

      for (let i = 0; i < 5; i++) {

        const percentage = total
          ? (sums[i] / total * 100)
          : 0;

        const currentRow = diagram.getRow(tableRow + 1 + i);

        currentRow.values = [
          LIKERT[i],
          SHORT[i],
          sums[i],
          `${percentage.toFixed(1)}%`
        ];

        currentRow.getCell(1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: [
              'FF8B5BB8',
              'FFA47ACB',
              'FFC3A8E4',
              'FF7542AE',
              'FF4B1B85'
            ][i]
          }
        };

        currentRow.getCell(1).font = {
          color: { argb: 'FFFFFFFF' },
          bold: true
        };
      }

      // KESIMPULAN DESKRIPTIF
      const conclusion = makeConclusion(
        group.name,
        sums,
        total
      );

      const conclusionRow = tableRow + 8;

      diagram.mergeCells(
        `A${conclusionRow}:H${conclusionRow}`
      );

      diagram.getCell(`A${conclusionRow}`).value =
        'KESIMPULAN DESKRIPTIF';

      diagram.getCell(`A${conclusionRow}`).font = {
        bold: true,
        size: 12,
        color: { argb: 'FF35136C' }
      };

      diagram.mergeCells(
        `A${conclusionRow + 1}:H${conclusionRow + 4}`
      );

      diagram.getCell(`A${conclusionRow + 1}`).value =
        conclusion;

      diagram.getCell(`A${conclusionRow + 1}`).alignment = {
        wrapText: true,
        vertical: 'top'
      };

      diagram.getCell(`A${conclusionRow + 1}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF2EAF9' }
      };

      for (let i = 1; i <= 4; i++) {
        diagram.getRow(conclusionRow + i).height = 25;
      }

      row = conclusionRow + 7;
    }

    // =====================================
    // DOWNLOAD FILE EXCEL
    // =====================================

    const buffer = await wb.xlsx.writeBuffer();

    const stamp = new Date()
      .toISOString()
      .slice(0, 10);

    return {
      statusCode: 200,
      isBase64Encoded: true,

      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

        'Content-Disposition':
          `attachment; filename="hasil-kuesioner-${stamp}.xlsx"`,

        'Cache-Control': 'no-store'
      },

      body: Buffer.from(buffer).toString('base64')
    };

  } catch (error) {

    console.error(error);

    return {
      statusCode: 500,

      headers: {
        'Content-Type': 'application/json'
      },

      body: JSON.stringify({
        error:
          error.message ||
          'Terjadi kesalahan saat membuat Excel.'
      })
    };
  }
};
  
