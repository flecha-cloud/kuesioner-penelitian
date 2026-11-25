
export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // =========================
    // EXPORT EXCEL
    // =========================
    if (url.pathname === "/api/export-excel") {
      if (request.method !== "GET") {
        return Response.json({ error: "Method not allowed" }, { status: 405 });
      }

      if (!env.DB) {
        return Response.json({ error: "Database belum terhubung." }, { status: 500 });
      }

      if (!env.EXPORT_PASSWORD) {
        return Response.json({ error: "Password ekspor belum diatur." }, { status: 500 });
      }

      if (request.headers.get("x-export-password") !== env.EXPORT_PASSWORD) {
        return Response.json({ error: "Password salah." }, { status: 401 });
      }

      try {
        const result = await env.DB.prepare(`
          SELECT submission_id, submitted_at, gender, age, education,
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

        const shortChoices = ["STS", "TS", "Netral", "S", "SS"];

        const pct = (n, total) => total ? n / total : 0;
        const percentText = (n, total) =>
          `${(pct(n, total) * 100).toFixed(2)}%`;

        const countValues = values => {
          const result = {};
          for (const value of values) {
            const label = String(value || "Tidak diisi");
            result[label] = (result[label] || 0) + 1;
          }
          return result;
        };

        const countAnswers = values =>
          choices.map(choice => values.filter(v => v === choice).length);

        // =========================
        // FUNGSI MEMBUAT DIAGRAM
        // =========================
        async function createPieChart(labels, values, title) {
          const chartConfig = {
            type: "pie",
            data: {
              labels,
              datasets: [{
                data: values,
                backgroundColor: [
                  "#8064C8",
                  "#A58ADD",
                  "#D0BDF0",
                  "#6D9EEB",
                  "#79C9B0",
                  "#F3C46B",
                  "#E98B9B",
                  "#8CABC7"
                ],
                borderColor: "#FFFFFF",
                borderWidth: 2
              }]
            },
            options: {
              responsive: false,
              title: {
                display: true,
                text: title,
                fontSize: 18
              },
              legend: {
                position: "right"
              },
              plugins: {
                datalabels: {
                  display: true,
                  color: "#FFFFFF",
                  font: { weight: "bold", size: 13 },
                  formatter: function(value, context) {
                    const data = context.chart.data.datasets[0].data;
                    const total = data.reduce((a, b) => a + b, 0);
                    return total ? (value / total * 100).toFixed(1) + "%" : "0%";
                  }
                }
              }
            }
          };

          const response = await fetch("https://quickchart.io/chart", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              width: 650,
              height: 400,
              format: "png",
              version: "2",
              backgroundColor: "white",
              chart: chartConfig
            })
          });

          if (!response.ok) {
            throw new Error("Gagal membuat diagram. Silakan coba ekspor kembali.");
          }

          return new Uint8Array(await response.arrayBuffer());
        }

        // =========================
        // SHEET 1: DATA RESPONDEN
        // =========================
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

        // =========================
        // SHEET 2: REKAPITULASI
        // =========================
        const recap = workbook.addWorksheet("Rekapitulasi");

        recap.columns = [
          { header: "Pertanyaan", key: "question", width: 15 },
          { header: "Jawaban", key: "answer", width: 28 },
          { header: "Jumlah", key: "count", width: 14 },
          { header: "Persentase", key: "percentage", width: 16 },
          { header: "Total Jawaban", key: "total", width: 18 }
        ];

        for (let i = 1; i <= 30; i++) {
          const values = rows.map(row => row[`P${i}`]);
          const counts = countAnswers(values);
          const total = counts.reduce((a, b) => a + b, 0);

          choices.forEach((choice, j) => {
            recap.addRow({
              question: `P${i}`,
              answer: choice,
              count: counts[j],
              percentage: pct(counts[j], total),
              total
            });
          });
        }

        recap.getColumn("percentage").numFmt = "0.00%";
        recap.views = [{ state: "frozen", ySplit: 1 }];

        // =========================
        // SHEET 3: DIAGRAM
        // =========================
        const charts = workbook.addWorksheet("Diagram dan Persentase");

        charts.mergeCells("A1:U1");
        charts.getCell("A1").value = "LAPORAN DIAGRAM HASIL PENELITIAN";
        charts.getCell("A1").font = {
          bold: true,
          size: 17,
          color: { argb: "FFFFFFFF" }
        };
        charts.getCell("A1").fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF6D4BC3" }
        };
        charts.getCell("A1").alignment = { horizontal: "center" };
        charts.getRow(1).height = 32;

        charts.mergeCells("A2:U2");
        charts.getCell("A2").value =
          `Jumlah Responden: ${rows.length}`;
        charts.getCell("A2").font = { bold: true, size: 12 };

        for (let c = 1; c <= 21; c++) {
          charts.getColumn(c).width = 13;
        }

        const chartItems = [
          {
            title: "Jenis Kelamin",
            values: rows.map(r => r.gender),
            position: "A4"
          },
          {
            title: "Kelompok Usia",
            values: rows.map(r => r.age),
            position: "H4"
          },
          {
            title: "Pendidikan",
            values: rows.map(r => r.education),
            position: "O4"
          }
        ];

        const groups = [
          { title: "Motivasi Kerja", start: 1, end: 10, position: "A25" },
          { title: "Disiplin Kerja", start: 11, end: 20, position: "H25" },
          { title: "Kinerja Pegawai", start: 21, end: 30, position: "O25" }
        ];

        for (const group of groups) {
          const values = [];

          for (let i = group.start; i <= group.end; i++) {
            for (const row of rows) {
              values.push(row[`P${i}`]);
            }
          }

          chartItems.push({
            title: group.title,
            values,
            position: group.position
          });
        }

        // Tabel persentase dan diagram
        for (let index = 0; index < chartItems.length; index++) {
          const item = chartItems[index];
          const counts = countValues(item.values);
          const labels = Object.keys(counts).filter(
            label => label !== "Tidak diisi"
          );
          const values = labels.map(label => counts[label]);
          const total = values.reduce((a, b) => a + b, 0);

          const baseRow = index < 3 ? 5 : 26;
          const baseCol = index % 3 === 0 ? 1 :
                          index % 3 === 1 ? 8 : 15;

          charts.getCell(baseRow, baseCol).value = item.title;
          charts.getCell(baseRow, baseCol).font = {
            bold: true,
            size: 12,
            color: { argb: "FF6D4BC3" }
          };

          charts.getCell(baseRow + 1, baseCol).value = "Kategori";
          charts.getCell(baseRow + 1, baseCol + 1).value = "Jumlah";
          charts.getCell(baseRow + 1, baseCol + 2).value = "Persentase";

          labels.forEach((label, i) => {
            charts.getCell(baseRow + 2 + i, baseCol).value = label;
            charts.getCell(baseRow + 2 + i, baseCol + 1).value = values[i];
            charts.getCell(baseRow + 2 + i, baseCol + 2).value = pct(values[i], total);
            charts.getCell(baseRow + 2 + i, baseCol + 2).numFmt = "0.00%";
          });

          if (total > 0) {
            const imageBuffer = await createPieChart(
              labels,
              values,
              item.title
            );

            const imageId = workbook.addImage({
              buffer: imageBuffer,
              extension: "png"
            });

            charts.addImage(imageId, {
              tl: {
                col: baseCol - 1,
                row: index < 3 ? 11 : 32
              },
              ext: { width: 420, height: 260 }
            });
          }
        }

        // =========================
        // SHEET 4: PENJELASAN
        // =========================
        const explanation = workbook.addWorksheet("Penjelasan dan Kesimpulan");

        explanation.columns = [
          { header: "Bagian", key: "section", width: 30 },
          { header: "Penjelasan Hasil", key: "description", width: 110 }
        ];

        explanation.addRow({
          section: "Jumlah Responden",
          description: `${rows.length} orang responden mengisi kuesioner.`
        });

        // Demografi
        const demographics = [
          { title: "Jenis Kelamin", key: "gender" },
          { title: "Usia", key: "age" },
          { title: "Pendidikan", key: "education" }
        ];

        for (const demo of demographics) {
          const counts = countValues(rows.map(r => r[demo.key]));
          const total = Object.values(counts).reduce((a, b) => a + b, 0);

          const description = Object.entries(counts)
            .map(([label, count]) =>
              `${label}: ${count} responden (${percentText(count, total)})`
            )
            .join("; ");

          explanation.addRow({
            section: demo.title,
            description: description || "Belum ada data."
          });
        }

        // Analisis variabel
        const variableResults = [];

        for (const group of groups) {
          const values = [];

          for (let i = group.start; i <= group.end; i++) {
            for (const row of rows) {
              values.push(row[`P${i}`]);
            }
          }

          const counts = countAnswers(values);
          const total = counts.reduce((a, b) => a + b, 0);

          const positive = counts[3] + counts[4];
          const neutral = counts[2];
          const negative = counts[0] + counts[1];

          explanation.addRow({
            section: group.title,
            description: `Variabel ini diukur melalui pertanyaan P${group.start} sampai P${group.end}.`
          });

          explanation.addRow({
            section: "Distribusi Jawaban",
            description: choices.map((choice, i) =>
              `${choice}: ${counts[i]} jawaban (${percentText(counts[i], total)})`
            ).join("; ")
          });

          explanation.addRow({
            section: "Interpretasi " + group.title,
            description:
              `Jawaban positif (Setuju dan Sangat Setuju) sebesar ${percentText(positive, total)}. ` +
              `Jawaban netral sebesar ${percentText(neutral, total)}, ` +
              `sedangkan jawaban negatif sebesar ${percentText(negative, total)}.`
          });

          variableResults.push({
            title: group.title,
            positive: pct(positive, total),
            neutral: pct(neutral, total),
            negative: pct(negative, total)
          });
        }

        // Kesimpulan otomatis
        const conclusion = rows.length
          ? variableResults.map(v =>
              `${v.title}: jawaban positif ${percentText(v.positive * 100, 100)}, ` +
              `netral ${percentText(v.neutral * 100, 100)}, ` +
              `dan negatif ${percentText(v.negative * 100, 100)}.`
            ).join(" ") +
            " Secara deskriptif, hasil ini menggambarkan kecenderungan jawaban responden pada masing-masing variabel. " +
            "Kesimpulan mengenai pengaruh Motivasi Kerja dan Disiplin Kerja terhadap Kinerja Pegawai " +
            "belum dapat ditetapkan hanya dari persentase. Diperlukan pengujian statistik seperti " +
            "uji validitas, reliabilitas, regresi, dan uji hipotesis."
          : "Belum ada data responden sehingga kesimpulan belum dapat dibuat.";

        explanation.addRow({
          section: "KESIMPULAN",
          description: conclusion
        });

        // =========================
        // FORMAT EXCEL
        // =========================
        for (const sheet of [dataSheet, recap, explanation]) {
          sheet.getRow(1).height = 28;
          sheet.getRow(1).font = {
            bold: true,
            color: { argb: "FFFFFFFF" }
          };
          sheet.getRow(1).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF6D4BC3" }
          };
          sheet.getRow(1).alignment = {
            horizontal: "center",
            vertical: "middle",
            wrapText: true
          };
        }

        explanation.eachRow(row => {
          row.alignment = {
            vertical: "top",
            wrapText: true
          };
          row.height = 36;
        });

        const buffer = await workbook.xlsx.writeBuffer();

        return new Response(buffer, {
          headers: {
            "Content-Type":
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition":
              'attachment; filename="laporan-hasil-penelitian.xlsx"',
            "Cache-Control": "no-store"
          }
        });

      } catch (error) {
        return Response.json({
          success: false,
          error: "Gagal membuat laporan Excel: " +
            String(error?.message || error)
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
        return Response.json({ success: false, error: "Database binding DB belum dipasang." }, { status: 500 });
      }

      try {
        const data = await request.json();

        const required = [
          "submission_id",
          "submitted_at",
          "gender",
          "age",
          "education"
        ];

        for (const key of required) {
          if (typeof data[key] !== "string" || !data[key].trim()) {
            return Response.json({
              success: false,
              error: `Data ${key} belum lengkap.`
            }, { status: 400 });
          }
        }

        const answers = [];

        for (let i = 1; i <= 30; i++) {
          const value = data[`P${i}`];

          if (![
            "Sangat Tidak Setuju",
            "Tidak Setuju",
            "Netral",
            "Setuju",
            "Sangat Setuju"
          ].includes(value)) {
            return Response.json({
              success: false,
              error: `Jawaban pertanyaan ${i} tidak valid.`
            }, { status: 400 });
          }

          answers.push(value);
        }

        const columns = [
          "submission_id",
          "submitted_at",
          "gender",
          "age",
          "education",
          ...Array.from({ length: 30 }, (_, i) => `P${i + 1}`)
        ];

        const values = [
          data.submission_id,
          data.submitted_at,
          data.gender,
          data.age,
          data.education,
          ...answers
        ];

        const placeholders = columns.map(() => "?").join(", ");

        await env.DB.prepare(
          `INSERT INTO responses (${columns.join(", ")}) VALUES (${placeholders})`
        ).bind(...values).run();

        return Response.json({ success: true });

      } catch (error) {
        const message = String(error?.message || error);
        const status = message.toLowerCase().includes("unique") ? 409 : 500;

        return Response.json({
          success: false,
          error: status === 409
            ? "Jawaban ini sudah pernah dikirim."
            : "Gagal menyimpan jawaban. Periksa tabel dan binding database."
        }, { status });
      }
    }

    return env.ASSETS.fetch(request);
  }
};
       { header: "Hasil Analisis", key: "analysis", width: 110 }
        ];

        const respondentCount = rows.length;

        explanation.addRow({
          section: "Jumlah Responden",
          analysis: `${respondentCount} responden`
        });

        for (const demo of demographics) {
          const counts = countValues(rows.map(r => r[demo.key]));
          const total = Object.values(counts).reduce((a, b) => a + b, 0);
          const description = Object.entries(counts)
            .map(([label, count]) =>
              `${label}: ${count} responden (${formatPercent(percent(count, total))})`
            ).join("; ");

          explanation.addRow({
            section: demo.title,
            analysis: description || "Belum ada data."
          });
        }

        const variableResults = [];

        for (const group of variableGroups) {
          const values = [];
          for (let i = group.start; i <= group.end; i++) {
            for (const row of rows) values.push(row[`P${i}`]);
          }

          const counts = countChoices(values);
          const total = counts.reduce((a, b) => a + b, 0);
          const positive = counts[3] + counts[4];
          const negative = counts[0] + counts[1];
          const neutral = counts[2];

          const description = choices.map((choice, i) =>
            `${choice}: ${counts[i]} (${formatPercent(percent(counts[i], total))})`
          ).join("; ");

          explanation.addRow({
            section: group.title,
            analysis: `Distribusi jawaban dari P${group.start} sampai P${group.end}. ${description}.`
          });

          explanation.addRow({
            section: `Interpretasi ${group.title}`,
            analysis: `Jawaban positif (Setuju dan Sangat Setuju) sebesar ${formatPercent(percent(positive, total))}; netral ${formatPercent(percent(neutral, total))}; negatif (Tidak Setuju dan Sangat Tidak Setuju) ${formatPercent(percent(negative, total))}.`
          });

          variableResults.push({
            title: group.title,
            positive: percent(positive, total),
            negative: percent(negative, total),
            neutral: percent(neutral, total)
          });
        }

        explanation.addRow({
          section: "KESIMPULAN",
          analysis: respondentCount
            ? variableResults.map(v =>
                `${v.title} memperoleh jawaban positif ${formatPercent(v.positive)}, netral ${formatPercent(v.neutral)}, dan negatif ${formatPercent(v.negative)}.`
              ).join(" ") +
              " Hasil ini merupakan gambaran deskriptif jawaban responden. Untuk menyimpulkan pengaruh antarvariabel secara statistik, diperlukan pengujian lebih lanjut seperti uji validitas, reliabilitas, regresi, dan uji hipotesis."
            : "Belum ada responden yang mengisi kuesioner sehingga kesimpulan belum dapat dibuat."
        });

        // Format semua header
        for (const sheet of [dataSheet, recap, explanation]) {
          sheet.getRow(1).height = 26;
          sheet.getRow(1).font = {
            bold: true,
            color: { argb: "FFFFFFFF" }
          };
          sheet.getRow(1).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF6D4BC3" }
          };
          sheet.getRow(1).alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true
          };
        }

        explanation.eachRow(row => {
          row.alignment = { vertical: "top", wrapText: true };
        });

        const finalBuffer = await workbook.xlsx.writeBuffer();

        // Tambahkan kembali diagram setelah semua sheet selesai
        const finalWithCharts = chartSpecs.length
          ? await addCharts(finalBuffer, chartSpecs, { sheet: "Ringkasan dan Diagram" })
          : finalBuffer;

        return new Response(finalWithCharts, {
          headers: {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": 'attachment; filename="laporan-hasil-penelitian.xlsx"',
            "Cache-Control": "no-store"
          }
        });

      } catch (error) {
        return Response.json({
          success: false,
          error: "Gagal membuat laporan Excel: " + String(error?.message || error)
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
        return Response.json({ success: false, error: "Database binding DB belum dipasang." }, { status: 500 });
      }

      try {
        const data = await request.json();

        const required = [
          "submission_id",
          "submitted_at",
          "gender",
          "age",
          "education"
        ];

        for (const key of required) {
          if (typeof data[key] !== "string" || !data[key].trim()) {
            return Response.json({
              success: false,
              error: `Data ${key} belum lengkap.`
            }, { status: 400 });
          }
        }

        const answers = [];

        for (let i = 1; i <= 30; i++) {
          const value = data[`P${i}`];

          if (![
            "Sangat Tidak Setuju",
            "Tidak Setuju",
            "Netral",
            "Setuju",
            "Sangat Setuju"
          ].includes(value)) {
            return Response.json({
              success: false,
              error: `Jawaban pertanyaan ${i} tidak valid.`
            }, { status: 400 });
          }

          answers.push(value);
        }

        const columns = [
          "submission_id",
          "submitted_at",
          "gender",
          "age",
          "education",
          ...Array.from({ length: 30 }, (_, i) => `P${i + 1}`)
        ];

        const values = [
          data.submission_id,
          data.submitted_at,
          data.gender,
          data.age,
          data.education,
          ...answers
        ];

        const placeholders = columns.map(() => "?").join(", ");

        await env.DB.prepare(
          `INSERT INTO responses (${columns.join(", ")}) VALUES (${placeholders})`
        ).bind(...values).run();

        return Response.json({ success: true });

      } catch (error) {
        const message = String(error?.message || error);
        const status = message.toLowerCase().includes("unique") ? 409 : 500;

        return Response.json({
          success: false,
          error: status === 409
            ? "Jawaban ini sudah pernah dikirim."
            : "Gagal menyimpan jawaban. Periksa tabel dan binding database."
        }, { status });
      }
    }

    return env.ASSETS.fetch(request);
  }
};
