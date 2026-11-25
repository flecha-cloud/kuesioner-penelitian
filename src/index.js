
export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // =========================
    // EXPORT EXCEL + DIAGRAM
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
          SELECT submission_id, submitted_at, gender, age, education,
          ${Array.from({ length: 30 }, (_, i) => `P${i + 1}`).join(", ")}
          FROM responses ORDER BY submitted_at ASC
        `).all();

        const rows = result.results || [];
        const ExcelJS = (await import("exceljs")).default;
        const chartModule = await import("chartsheet");
        const addCharts = chartModule.addCharts || chartModule.default?.addCharts;

        if (!addCharts) {
          throw new Error("Fitur diagram gagal dimuat. Periksa instalasi chartsheet.");
        }

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

        const percent = (n, total) =>
          total ? n / total : 0;

        const formatPercent = (n) =>
          `${(n * 100).toFixed(2).replace(".", ",")}%`;

        const countValues = (values) => {
          const counts = {};
          for (const value of values) {
            const key = String(value || "Tidak diisi");
            counts[key] = (counts[key] || 0) + 1;
          }
          return counts;
        };

        const countChoices = (values) =>
          choices.map(choice => values.filter(v => v === choice).length);

        const getDominant = (counts, labels) => {
          if (!counts.length || !counts.some(v => v > 0)) return "Belum ada data";
          const max = Math.max(...counts);
          return labels.filter((label, i) => counts[i] === max).join(", ");
        };

        // =========================
        // SHEET 1: DATA RESPONDEN
        // =========================
        const dataSheet = workbook.addWorksheet("Data Responden");

        dataSheet.columns = [
          { header: "ID Pengiriman", key: "submission_id", width: 25 },
          { header: "Waktu Pengiriman", key: "submitted_at", width: 24 },
          { header: "Jenis Kelamin", key: "gender", width: 20 },
          { header: "Usia", key: "age", width: 16 },
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
          { header: "Pertanyaan", key: "question", width: 16 },
          { header: "Jawaban", key: "answer", width: 28 },
          { header: "Jumlah", key: "count", width: 14 },
          { header: "Persentase", key: "percentage", width: 16 },
          { header: "Total Responden", key: "total", width: 20 }
        ];

        for (let i = 1; i <= 30; i++) {
          const values = rows.map(row => row[`P${i}`]);
          const counts = countChoices(values);
          const total = counts.reduce((a, b) => a + b, 0);

          choices.forEach((choice, j) => {
            recap.addRow({
              question: `P${i}`,
              answer: choice,
              count: counts[j],
              percentage: percent(counts[j], total),
              total
            });
          });
        }

        recap.getColumn("percentage").numFmt = "0.00%";
        recap.views = [{ state: "frozen", ySplit: 1 }];

        // =========================
        // SHEET 3: RINGKASAN + DIAGRAM
        // =========================
        const summary = workbook.addWorksheet("Ringkasan dan Diagram");

        summary.mergeCells("A1:X1");
        summary.getCell("A1").value = "LAPORAN HASIL KUESIONER PENELITIAN";
        summary.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
        summary.getCell("A1").fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF6D4BC3" }
        };
        summary.getCell("A1").alignment = { horizontal: "center" };
        summary.getRow(1).height = 32;

        summary.mergeCells("A2:X2");
        summary.getCell("A2").value =
          "Pengaruh Motivasi Kerja dan Disiplin Kerja terhadap Kinerja Pegawai";
        summary.getCell("A2").font = { italic: true, size: 11 };
        summary.getCell("A2").alignment = { horizontal: "center" };

        const demographics = [
          { title: "Jenis Kelamin", key: "gender" },
          { title: "Kelompok Usia", key: "age" },
          { title: "Pendidikan", key: "education" }
        ];

        const chartSpecs = [];
        const chartBlocks = [
          { col: 1, title: "Jenis Kelamin", values: rows.map(r => r.gender) },
          { col: 5, title: "Kelompok Usia", values: rows.map(r => r.age) },
          { col: 9, title: "Pendidikan", values: rows.map(r => r.education) }
        ];

        const variableGroups = [
          { title: "Motivasi Kerja", start: 1, end: 10 },
          { title: "Disiplin Kerja", start: 11, end: 20 },
          { title: "Kinerja Pegawai", start: 21, end: 30 }
        ];

        for (const group of variableGroups) {
          const values = [];
          for (let i = group.start; i <= group.end; i++) {
            for (const row of rows) values.push(row[`P${i}`]);
          }
          chartBlocks.push({
            col: group.title === "Motivasi Kerja" ? 13 :
                 group.title === "Disiplin Kerja" ? 17 : 21,
            title: group.title,
            values
          });
        }

        // Buat tabel ringkasan untuk enam diagram
        for (const block of chartBlocks) {
          const counts = countValues(block.values);
          const labels = Object.keys(counts).filter(k => k !== "Tidak diisi");
          const total = labels.reduce((sum, label) => sum + counts[label], 0);

          const col = block.col;
          const colLetter = n => {
            let s = "";
            while (n > 0) {
              const r = (n - 1) % 26;
              s = String.fromCharCode(65 + r) + s;
              n = Math.floor((n - 1) / 26);
            }
            return s;
          };

          const labelCol = colLetter(col);
          const countCol = colLetter(col + 1);
          const pctCol = colLetter(col + 2);

          summary.getCell(4, col).value = block.title;
          summary.getCell(4, col).font = { bold: true, size: 12 };

          summary.getCell(5, col).value = "Kategori";
          summary.getCell(5, col + 1).value = "Jumlah";
          summary.getCell(5, col + 2).value = "Persentase";

          for (let i = 0; i < labels.length; i++) {
            const rowNum = 6 + i;
            summary.getCell(rowNum, col).value = labels[i];
            summary.getCell(rowNum, col + 1).value = counts[labels[i]];
            summary.getCell(rowNum, col + 2).value = percent(counts[labels[i]], total);
            summary.getCell(rowNum, col + 2).numFmt = "0.00%";
          }

          const lastRow = 5 + labels.length;

          if (labels.length && total > 0) {
            chartSpecs.push({
              type: "pie",
              title: block.title,
              legend: "b",
              dataLabels: true,
              categories: `'Ringkasan dan Diagram'!$${labelCol}$6:$${labelCol}$${lastRow}`,
              series: [{
                nameRef: `'Ringkasan dan Diagram'!$${pctCol}$5`,
                ref: `'Ringkasan dan Diagram'!$${pctCol}$6:$${pctCol}$${lastRow}`
              }],
              anchor: {
                col: block.col <= 9 ? (block.col - 1) * 2 : (block.col - 13) * 2,
                row: block.col <= 9 ? 12 : 30
              },
              width: 480,
              height: 300
            });
          }
        }

        // Format tabel ringkasan
        for (let r = 5; r <= 12; r++) {
          for (let c = 1; c <= 24; c++) {
            const cell = summary.getCell(r, c);
            if (r === 5 && cell.value) {
              cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FF8064C8" }
              };
            }
          }
        }

        // Tambahkan diagram ke file Excel
        let buffer = await workbook.xlsx.writeBuffer();

        if (chartSpecs.length) {
          buffer = await addCharts(buffer, chartSpecs, {
            sheet: "Ringkasan dan Diagram"
          });
        }

        // =========================
        // SHEET 4: PENJELASAN & KESIMPULAN
        // =========================
        const explanation = workbook.addWorksheet("Penjelasan dan Kesimpulan");

        explanation.columns = [
          { header: "Bagian", key: "section", width: 28 },
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
