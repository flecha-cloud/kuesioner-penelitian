export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // =========================
    // EXPORT EXCEL
    // =========================
    if (url.pathname === "/api/export-excel") {
      if (request.method !== "GET") {
        return Response.json({ success: false, error: "Method not allowed" }, { status: 405 });
      }

      if (!env.DB) {
        return Response.json({ success: false, error: "Database belum terhubung." }, { status: 500 });
      }

      if (!env.EXPORT_PASSWORD) {
        return Response.json({ success: false, error: "Password ekspor belum diatur di Cloudflare." }, { status: 500 });
      }

      const password = request.headers.get("x-export-password") || "";
      if (password !== env.EXPORT_PASSWORD) {
        return Response.json({ success: false, error: "Kata sandi ekspor salah." }, { status: 401 });
      }

      try {
        const result = await env.DB.prepare(`
          SELECT
            submission_id,
            submitted_at,
            gender,
            age,
            education,
            ${Array.from({ length: 30 }, (_, i) => `P${i + 1}`).join(", ")}
          FROM responses
          ORDER BY submitted_at ASC
        `).all();

        const rows = result.results || [];
        const ExcelJS = (await import("exceljs")).default;
        const workbook = new ExcelJS.Workbook();

        workbook.creator = "Kuesioner Penelitian";
        workbook.created = new Date();

        const choices = [
          "Sangat Tidak Setuju",
          "Tidak Setuju",
          "Netral",
          "Setuju",
          "Sangat Setuju"
        ];
        const shortChoices = ["STS", "TS", "N", "S", "SS"];

        const ratio = (value, total) => total ? value / total : 0;
        const percent = (value, total) => `${(ratio(value, total) * 100).toFixed(2)}%`;

        const countValues = values => {
          const counts = new Map();
          for (const value of values) {
            const label = String(value || "Tidak diisi");
            counts.set(label, (counts.get(label) || 0) + 1);
          }
          return counts;
        };

        const countLikert = values =>
          choices.map(choice => values.filter(value => value === choice).length);

        // Buat satu diagram pie PNG melalui QuickChart.
        // Semua gambar diagram dimasukkan ke sheet "Diagram".
        async function makePieChart(title, labels, values) {
          const total = values.reduce((sum, value) => sum + value, 0);
          const chartConfig = {
            type: "pie",
            data: {
              labels: labels.map((label, i) =>
                `${label} (${total ? (values[i] / total * 100).toFixed(1) : "0.0"}%)`
              ),
              datasets: [{
                data: values,
                backgroundColor: [
                  "#7051B8", "#9476D0", "#B9A5E6",
                  "#648EDB", "#64B99D", "#E6B85C",
                  "#DB8298", "#8BA7BF"
                ],
                borderColor: "#FFFFFF",
                borderWidth: 2
              }]
            },
            options: {
              responsive: false,
              maintainAspectRatio: false,
              title: {
                display: true,
                text: title,
                fontSize: 18,
                fontColor: "#30234D"
              },
              legend: {
                display: true,
                position: "right",
                labels: {
                  fontSize: 11,
                  boxWidth: 12,
                  padding: 10
                }
              },
              plugins: {
                datalabels: {
                  display: true,
                  color: "#FFFFFF",
                  font: { weight: "bold", size: 13 },
                  formatter: "(value, context) => { const data = context.chart.data.datasets[0].data; const total = data.reduce((a,b) => a+b, 0); return total ? (value/total*100).toFixed(1)+'%' : '0%'; }"
                }
              }
            }
          };

          const response = await fetch("https://quickchart.io/chart", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              width: 720,
              height: 440,
              devicePixelRatio: 2,
              format: "png",
              version: "2",
              backgroundColor: "#FFFFFF",
              chart: chartConfig
            })
          });

          if (!response.ok) {
            throw new Error("Gagal membuat diagram. Pastikan layanan diagram dapat diakses, lalu coba ekspor lagi.");
          }

          const contentType = response.headers.get("content-type") || "";
          if (!contentType.includes("image")) {
            throw new Error("Layanan diagram tidak mengembalikan gambar.");
          }

          return new Uint8Array(await response.arrayBuffer());
        }

        // SHEET 1: DATA RESPONDEN
        const dataSheet = workbook.addWorksheet("Data Responden");
        dataSheet.columns = [
          { header: "ID Pengiriman", key: "submission_id", width: 25 },
          { header: "Waktu Pengiriman", key: "submitted_at", width: 24 },
          { header: "Jenis Kelamin", key: "gender", width: 20 },
          { header: "Usia", key: "age", width: 18 },
          { header: "Pendidikan", key: "education", width: 24 },
          ...Array.from({ length: 30 }, (_, i) => ({
            header: `P${i + 1}`,
            key: `P${i + 1}`,
            width: 24
          }))
        ];
        dataSheet.addRows(rows);
        dataSheet.views = [{ state: "frozen", ySplit: 1 }];
        dataSheet.autoFilter = { from: "A1", to: "AI1" };

        // SHEET 2: REKAPITULASI JAWABAN
        const recap = workbook.addWorksheet("Rekapitulasi");
        recap.columns = [
          { header: "Pertanyaan", key: "question", width: 15 },
          { header: "Kategori Jawaban", key: "answer", width: 28 },
          { header: "Jumlah", key: "count", width: 14 },
          { header: "Persentase", key: "percentage", width: 16 },
          { header: "Total Jawaban", key: "total", width: 18 }
        ];

        for (let i = 1; i <= 30; i++) {
          const values = rows.map(row => row[`P${i}`]);
          const counts = countLikert(values);
          const total = counts.reduce((sum, value) => sum + value, 0);

          choices.forEach((choice, index) => {
            recap.addRow({
              question: `P${i}`,
              answer: choice,
              count: counts[index],
              percentage: ratio(counts[index], total),
              total
            });
          });
        }
        recap.getColumn("percentage").numFmt = "0.00%";
        recap.views = [{ state: "frozen", ySplit: 1 }];

        // Siapkan tepat 6 chart: 3 demografi + 3 variabel penelitian.
        const chartDefinitions = [
          {
            title: "1. Jenis Kelamin",
            values: rows.map(row => row.gender),
            anchor: { col: 0, row: 4 }
          },
          {
            title: "2. Kelompok Usia",
            values: rows.map(row => row.age),
            anchor: { col: 8, row: 4 }
          },
          {
            title: "3. Pendidikan",
            values: rows.map(row => row.education),
            anchor: { col: 16, row: 4 }
          }
        ];

        const variables = [
          { title: "4. Motivasi Kerja (P1–P10)", start: 1, end: 10, anchor: { col: 0, row: 27 } },
          { title: "5. Disiplin Kerja (P11–P20)", start: 11, end: 20, anchor: { col: 8, row: 27 } },
          { title: "6. Kinerja Pegawai (P21–P30)", start: 21, end: 30, anchor: { col: 16, row: 27 } }
        ];

        for (const variable of variables) {
          const values = [];
          for (let question = variable.start; question <= variable.end; question++) {
            for (const row of rows) values.push(row[`P${question}`]);
          }

          chartDefinitions.push({
            title: variable.title,
            values,
            anchor: variable.anchor,
            likert: true
          });
        }

        // SHEET 3: SEMUA CHART DI SINI SAJA
        const diagram = workbook.addWorksheet("Diagram");
        diagram.mergeCells("A1:W1");
        diagram.getCell("A1").value = "DIAGRAM HASIL PENELITIAN";
        diagram.getCell("A1").font = {
          bold: true,
          size: 19,
          color: { argb: "FFFFFFFF" }
        };
        diagram.getCell("A1").fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF7051B8" }
        };
        diagram.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
        diagram.getRow(1).height = 34;

        diagram.mergeCells("A2:W2");
        diagram.getCell("A2").value =
          `Jumlah responden: ${rows.length} | Seluruh enam diagram berada pada sheet ini`;
        diagram.getCell("A2").font = { italic: true, color: { argb: "FF555555" } };
        diagram.getCell("A2").alignment = { horizontal: "center" };

        for (let col = 1; col <= 23; col++) diagram.getColumn(col).width = 12;

        // Tabel angka dan persentase diletakkan di bawah setiap diagram.
        for (let index = 0; index < chartDefinitions.length; index++) {
          const item = chartDefinitions[index];
          const isLikert = Boolean(item.likert);

          let labels;
          let values;

          if (isLikert) {
            labels = shortChoices;
            values = countLikert(item.values);
          } else {
            const counts = countValues(item.values);
            labels = [...counts.keys()];
            values = labels.map(label => counts.get(label));
          }

          const total = values.reduce((sum, value) => sum + value, 0);
          const startRow = index < 3 ? 22 : 45;
          const startCol = item.anchor.col + 1;

          diagram.getCell(startRow, startCol).value = item.title;
          diagram.getCell(startRow, startCol).font = {
            bold: true,
            color: { argb: "FF7051B8" },
          };

          const headerRow = startRow + 1;
          diagram.getCell(headerRow, startCol).value = "Kategori";
          diagram.getCell(headerRow, startCol + 1).value = "Jumlah";
          diagram.getCell(headerRow, startCol + 2).value = "Persentase";

          [startCol, startCol + 1, startCol + 2].forEach(col => {
            diagram.getCell(headerRow, col).font = {
              bold: true,
              color: { argb: "FFFFFFFF" }
            };
            diagram.getCell(headerRow, col).fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FF7051B8" }
            };
          });

          labels.forEach((label, i) => {
            const rowNumber = headerRow + 1 + i;
            diagram.getCell(rowNumber, startCol).value = label;
            diagram.getCell(rowNumber, startCol + 1).value = values[i];
            diagram.getCell(rowNumber, startCol + 2).value = ratio(values[i], total);
            diagram.getCell(rowNumber, startCol + 2).numFmt = "0.00%";
          });

          const imageBytes = await makePieChart(item.title, labels, values);
          const imageId = workbook.addImage({
            buffer: imageBytes,
            extension: "png"
          });

          diagram.addImage(imageId, {
            tl: {
              col: item.anchor.col,
              row: item.anchor.row
            },
            ext: {
              width: 440,
              height: 270
            }
          });
        }

        // SHEET 4: PENJELASAN DAN KESIMPULAN
        const explanation = workbook.addWorksheet("Penjelasan dan Kesimpulan");
        explanation.columns = [
          { width: 30 },
          { width: 100 }
        ];

        explanation.mergeCells("A1:B1");
        explanation.getCell("A1").value = "PENJELASAN HASIL PENELITIAN";
        explanation.getCell("A1").font = {
          bold: true,
          size: 18,
          color: { argb: "FFFFFFFF" }
        };
        explanation.getCell("A1").fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF7051B8" }
        };
        explanation.getCell("A1").alignment = { horizontal: "center" };
        explanation.getRow(1).height = 32;

        const demographicDescriptions = [
          {
            title: "Jenis Kelamin",
            values: rows.map(row => row.gender)
          },
          {
            title: "Kelompok Usia",
            values: rows.map(row => row.age)
          },
          {
            title: "Pendidikan",
            values: rows.map(row => row.education)
          }
        ];

        let currentRow = 3;

        explanation.getCell(`A${currentRow}`).value = "A. DEMOGRAFI RESPONDEN";
        explanation.getCell(`A${currentRow}`).font = { bold: true, size: 13 };
        currentRow += 1;

        for (const item of demographicDescriptions) {
          const counts = countValues(item.values);
          const total = item.values.length;
          const description = [...counts.entries()]
            .map(([label, count]) => `${label}: ${count} responden (${percent(count, total)})`)
            .join("; ");

          explanation.getCell(`A${currentRow}`).value = item.title;
          explanation.getCell(`B${currentRow}`).value =
            total
              ? `Distribusi responden berdasarkan ${item.title.toLowerCase()} adalah ${description}.`
              : `Belum terdapat data responden untuk kategori ${item.title.toLowerCase()}.`;

          currentRow++;
        }

        currentRow += 1;
        explanation.getCell(`A${currentRow}`).value = "B. HASIL VARIABEL PENELITIAN";
        explanation.getCell(`A${currentRow}`).font = { bold: true, size: 13 };
        currentRow++;

        const variableDescriptions = [
          { title: "Motivasi Kerja", start: 1, end: 10 },
          { title: "Disiplin Kerja", start: 11, end: 20 },
          { title: "Kinerja Pegawai", start: 21, end: 30 }
        ];

        for (const variable of variableDescriptions) {
          const values = [];

          for (let question = variable.start; question <= variable.end; question++) {
            for (const row of rows) {
              const answer = row[`P${question}`];
              if (choices.includes(answer)) values.push(answer);
            }
          }

          const counts = countLikert(values);
          const total = counts.reduce((sum, value) => sum + value, 0);

          const distribution = choices
            .map((choice, index) =>
              `${choice}: ${counts[index]} (${percent(counts[index], total)})`
            )
            .join("; ");

          const positive = counts[3] + counts[4];
          const positivePercentage = percent(positive, total);

          explanation.getCell(`A${currentRow}`).value = variable.title;
          explanation.getCell(`B${currentRow}`).value =
            total
              ? `Distribusi jawaban pada variabel ${variable.title} menunjukkan bahwa jawaban Setuju dan Sangat Setuju berjumlah ${positive} jawaban (${positivePercentage}). Rincian seluruh jawaban: ${distribution}.`
              : `Belum terdapat jawaban yang dapat dianalisis untuk variabel ${variable.title}.`;

          currentRow++;
        }

        currentRow += 1;
        explanation.getCell(`A${currentRow}`).value = "C. KESIMPULAN DESKRIPTIF";
        explanation.getCell(`A${currentRow}`).font = { bold: true, size: 13 };
        currentRow++;

        explanation.getCell(`A${currentRow}`).value = "Kesimpulan";
        explanation.getCell(`B${currentRow}`).value =
          `Berdasarkan data yang terkumpul dari ${rows.length} responden, hasil penelitian dapat dideskripsikan melalui distribusi jawaban pada variabel Motivasi Kerja, Disiplin Kerja, dan Kinerja Pegawai. Persentase pada diagram menunjukkan proporsi jawaban responden terhadap setiap kategori.`;

        currentRow += 1;
        explanation.getCell(`A${currentRow}`).value = "Catatan Analisis";
        explanation.getCell(`B${currentRow}`).value =
          "Hasil persentase dan diagram ini merupakan analisis deskriptif. Untuk menyimpulkan apakah Motivasi Kerja dan Disiplin Kerja berpengaruh terhadap Kinerja Pegawai, diperlukan pengujian statistik lanjutan seperti uji validitas, reliabilitas, asumsi klasik, regresi linear berganda, uji t, uji F, dan koefisien determinasi sesuai rancangan penelitian.";

        explanation.getColumn(2).alignment = {
          wrapText: true,
          vertical: "top"
        };

        for (let row = 3; row <= currentRow; row++) {
          explanation.getRow(row).height = 42;
        }

        // Format umum seluruh worksheet.
        for (const sheet of [dataSheet, recap, diagram, explanation]) {
          sheet.eachRow(row => {
            row.eachCell(cell => {
              cell.alignment = {
                ...cell.alignment,
                vertical: "middle"
              };
            });
          });
        }

        const buffer = await workbook.xlsx.writeBuffer();

        return new Response(buffer, {
          status: 200,
          headers: {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": 'attachment; filename="Hasil_Kuesioner_Penelitian.xlsx"',
            "Cache-Control": "no-store"
          }
        });
      } catch (error) {
        return Response.json({
          success: false,
          error: error.message || "Terjadi kesalahan saat membuat file Excel."
        }, { status: 500 });
      }
    }

    // =========================
    // SUBMIT KUESIONER
    // =========================
    if (url.pathname === "/api/submit") {
      if (request.method !== "POST") {
        return Response.json({ success: false, error: "Method not allowed" }, { status: 405 });
      }

      if (!env.DB) {
        return Response.json({ success: false, error: "Database belum terhubung." }, { status: 500 });
      }

      try {
        const body = await request.json();

        const submissionId = body.submission_id || crypto.randomUUID();
        const submittedAt = body.submitted_at || new Date().toISOString();

        const gender = body.gender || "";
        const age = body.age || "";
        const education = body.education || "";

        const answers = body.answers || body;

        const columns = [
          "submission_id",
          "submitted_at",
          "gender",
          "age",
          "education",
          ...Array.from({ length: 30 }, (_, i) => `P${i + 1}`)
        ];

        const values = [
          submissionId,
          submittedAt,
          gender,
          age,
          education,
          ...Array.from({ length: 30 }, (_, i) => {
            const key = `P${i + 1}`;
            return answers[key] || "";
          })
        ];

        const placeholders = columns.map(() => "?").join(", ");

        await env.DB.prepare(`
          INSERT INTO responses (${columns.join(", ")})
          VALUES (${placeholders})
        `).bind(...values).run();

        return Response.json({
          success: true,
          message: "Jawaban berhasil dikirim.",
          submission_id: submissionId
        });
      } catch (error) {
        return Response.json({
          success: false,
          error: error.message || "Gagal menyimpan jawaban."
        }, { status: 500 });
      }
    }

    // =========================
    // STATIC ASSETS
    // =========================
    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Not Found", { status: 404 });
  }
};

            bold: true,
            color: { argb: "FF7051B8" },
          };
                    const headerRow = startRow + 1;
          diagram.getCell(headerRow, startCol).value = "Kategori";
          diagram.getCell(headerRow, startCol + 1).value = "Jumlah";
          diagram.getCell(headerRow, startCol + 2).value = "Persentase";

          [startCol, startCol + 1, startCol + 2].forEach(col => {
            diagram.getCell(headerRow, col).font = {
              bold: true,
              color: { argb: "FFFFFFFF" }
            };
            diagram.getCell(headerRow, col).fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FF7051B8" }
            };
          });

          labels.forEach((label, i) => {
            const rowNumber = headerRow + 1 + i;
            diagram.getCell(rowNumber, startCol).value = label;
            diagram.getCell(rowNumber, startCol + 1).value = values[i];
            diagram.getCell(rowNumber, startCol + 2).value = ratio(values[i], total);
            diagram.getCell(rowNumber, startCol + 2).numFmt = "0.00%";
          });

          const imageBytes = await makePieChart(item.title, labels, values);
          const imageId = workbook.addImage({
            buffer: imageBytes,
            extension: "png"
          });

          diagram.addImage(imageId, {
            tl: {
              col: item.anchor.col,
              row: item.anchor.row
            },
            ext: {
              width: 440,
              height: 270
            }
          });
        }

        // SHEET 4: PENJELASAN DAN KESIMPULAN
        const explanation = workbook.addWorksheet("Penjelasan dan Kesimpulan");
        explanation.columns = [
          { width: 30 },
          { width: 100 }
        ];

        explanation.mergeCells("A1:B1");
        explanation.getCell("A1").value = "PENJELASAN HASIL PENELITIAN";
        explanation.getCell("A1").font = {
          bold: true,
          size: 18,
          color: { argb: "FFFFFFFF" }
        };
        explanation.getCell("A1").fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF7051B8" }
        };
        explanation.getCell("A1").alignment = { horizontal: "center" };
        explanation.getRow(1).height = 32;

        const demographicDescriptions = [
          {
            title: "Jenis Kelamin",
            values: rows.map(row => row.gender)
          },
          {
            title: "Kelompok Usia",
            values: rows.map(row => row.age)
          },
          {
            title: "Pendidikan",
            values: rows.map(row => row.education)
          }
        ];

        let currentRow = 3;

        explanation.getCell(`A${currentRow}`).value = "A. DEMOGRAFI RESPONDEN";
        explanation.getCell(`A${currentRow}`).font = { bold: true, size: 13 };
        currentRow += 1;

        for (const item of demographicDescriptions) {
          const counts = countValues(item.values);
          const total = item.values.length;
          const description = [...counts.entries()]
            .map(([label, count]) => `${label}: ${count} responden (${percent(count, total)})`)
            .join("; ");

          explanation.getCell(`A${currentRow}`).value = item.title;
          explanation.getCell(`B${currentRow}`).value =
            total
              ? `Distribusi responden berdasarkan ${item.title.toLowerCase()} adalah ${description}.`
              : `Belum terdapat data responden untuk kategori ${item.title.toLowerCase()}.`;

          currentRow++;
        }

        currentRow += 1;
        explanation.getCell(`A${currentRow}`).value = "B. HASIL VARIABEL PENELITIAN";
        explanation.getCell(`A${currentRow}`).font = { bold: true, size: 13 };
        currentRow++;

        const variableDescriptions = [
          { title: "Motivasi Kerja", start: 1, end: 10 },
          { title: "Disiplin Kerja", start: 11, end: 20 },
          { title: "Kinerja Pegawai", start: 21, end: 30 }
        ];

        for (const variable of variableDescriptions) {
          const values = [];

          for (let question = variable.start; question <= variable.end; question++) {
            for (const row of rows) {
              const answer = row[`P${question}`];
              if (choices.includes(answer)) values.push(answer);
            }
          }

          const counts = countLikert(values);
          const total = counts.reduce((sum, value) => sum + value, 0);

          const distribution = choices
            .map((choice, index) =>
              `${choice}: ${counts[index]} (${percent(counts[index], total)})`
            )
            .join("; ");

          const positive = counts[3] + counts[4];
          const positivePercentage = percent(positive, total);

          explanation.getCell(`A${currentRow}`).value = variable.title;
          explanation.getCell(`B${currentRow}`).value =
            total
              ? `Distribusi jawaban pada variabel ${variable.title} menunjukkan bahwa jawaban Setuju dan Sangat Setuju berjumlah ${positive} jawaban (${positivePercentage}). Rincian seluruh jawaban: ${distribution}.`
              : `Belum terdapat jawaban yang dapat dianalisis untuk variabel ${variable.title}.`;

          currentRow++;
        }

        currentRow += 1;
        explanation.getCell(`A${currentRow}`).value = "C. KESIMPULAN DESKRIPTIF";
        explanation.getCell(`A${currentRow}`).font = { bold: true, size: 13 };
        currentRow++;

        explanation.getCell(`A${currentRow}`).value = "Kesimpulan";
        explanation.getCell(`B${currentRow}`).value =
          `Berdasarkan data yang terkumpul dari ${rows.length} responden, hasil penelitian dapat dideskripsikan melalui distribusi jawaban pada variabel Motivasi Kerja, Disiplin Kerja, dan Kinerja Pegawai. Persentase pada diagram menunjukkan proporsi jawaban responden terhadap setiap kategori.`;

        currentRow += 1;
        explanation.getCell(`A${currentRow}`).value = "Catatan Analisis";
        explanation.getCell(`B${currentRow}`).value =
          "Hasil persentase dan diagram ini merupakan analisis deskriptif. Untuk menyimpulkan apakah Motivasi Kerja dan Disiplin Kerja berpengaruh terhadap Kinerja Pegawai, diperlukan pengujian statistik lanjutan seperti uji validitas, reliabilitas, asumsi klasik, regresi linear berganda, uji t, uji F, dan koefisien determinasi sesuai rancangan penelitian.";

        explanation.getColumn(2).alignment = {
          wrapText: true,
          vertical: "top"
        };

        for (let row = 3; row <= currentRow; row++) {
          explanation.getRow(row).height = 42;
        }

        // Format umum seluruh worksheet.
        for (const sheet of [dataSheet, recap, diagram, explanation]) {
          sheet.eachRow(row => {
            row.eachCell(cell => {
              cell.alignment = {
                ...cell.alignment,
                vertical: "middle"
              };
            });
          });
        }

        const buffer = await workbook.xlsx.writeBuffer();

        return new Response(buffer, {
          status: 200,
          headers: {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": 'attachment; filename="Hasil_Kuesioner_Penelitian.xlsx"',
            "Cache-Control": "no-store"
          }
        });
      } catch (error) {
        return Response.json({
          success: false,
          error: error.message || "Terjadi kesalahan saat membuat file Excel."
        }, { status: 500 });
      }
    }

    // =========================
    // SUBMIT KUESIONER
    // =========================
    if (url.pathname === "/api/submit") {
      if (request.method !== "POST") {
        return Response.json({ success: false, error: "Method not allowed" }, { status: 405 });
      }

      if (!env.DB) {
        return Response.json({ success: false, error: "Database belum terhubung." }, { status: 500 });
      }

      try {
        const body = await request.json();

        const submissionId = body.submission_id || crypto.randomUUID();
        const submittedAt = body.submitted_at || new Date().toISOString();

        const gender = body.gender || "";
        const age = body.age || "";
        const education = body.education || "";

        const answers = body.answers || body;

        const columns = [
          "submission_id",
          "submitted_at",
          "gender",
          "age",
          "education",
          ...Array.from({ length: 30 }, (_, i) => `P${i + 1}`)
        ];

        const values = [
          submissionId,
          submittedAt,
          gender,
          age,
          education,
          ...Array.from({ length: 30 }, (_, i) => {
            const key = `P${i + 1}`;
            return answers[key] || "";
          })
        ];

        const placeholders = columns.map(() => "?").join(", ");

        await env.DB.prepare(`
          INSERT INTO responses (${columns.join(", ")})
          VALUES (${placeholders})
        `).bind(...values).run();

        return Response.json({
          success: true,
          message: "Jawaban berhasil dikirim.",
          submission_id: submissionId
        });
      } catch (error) {
        return Response.json({
          success: false,
          error: error.message || "Gagal menyimpan jawaban."
        }, { status: 500 });
      }
    }

    // =========================
    // STATIC ASSETS
    // =========================
    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Not Found", { status: 404 });
  }
};
