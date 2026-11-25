
export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // =========================
    // EXPORT EXCEL
    // =========================
    if (url.pathname === "/api/export-excel") {
      if (request.method !== "GET") {
        return Response.json(
          { success: false, error: "Method not allowed" },
          { status: 405 }
        );
      }

      if (!env.DB) {
        return Response.json(
          { success: false, error: "Database belum terhubung." },
          { status: 500 }
        );
      }

      if (!env.EXPORT_PASSWORD) {
        return Response.json(
          { success: false, error: "Password ekspor belum diatur di Cloudflare." },
          { status: 500 }
        );
      }

      const password = request.headers.get("x-export-password") || "";

      if (password !== env.EXPORT_PASSWORD) {
        return Response.json(
          { success: false, error: "Kata sandi ekspor salah." },
          { status: 401 }
        );
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

        // SHEET 1: DATA RESPONDEN
        const dataSheet = workbook.addWorksheet("Data Responden");

        dataSheet.columns = [
          { header: "ID Pengiriman", key: "submission_id", width: 25 },
          { header: "Waktu Pengiriman", key: "submitted_at", width: 24 },
          { header: "Jenis Kelamin", key: "gender", width: 18 },
          { header: "Usia", key: "age", width: 12 },
          { header: "Pendidikan", key: "education", width: 22 },
          ...Array.from({ length: 30 }, (_, i) => ({
            header: `P${i + 1}`,
            key: `P${i + 1}`,
            width: 20
          }))
        ];

        dataSheet.addRows(rows);
        dataSheet.views = [{ state: "frozen", ySplit: 1 }];
        dataSheet.autoFilter = {
          from: "A1",
          to: "AI1"
        };

        // SHEET 2: REKAPITULASI
        const recapSheet = workbook.addWorksheet("Rekapitulasi");

        const choices = [
          "Sangat Tidak Setuju",
          "Tidak Setuju",
          "Netral",
          "Setuju",
          "Sangat Setuju"
        ];

        recapSheet.columns = [
          { header: "Pertanyaan", key: "question", width: 18 },
          { header: choices[0], key: "sts", width: 24 },
          { header: choices[1], key: "ts", width: 18 },
          { header: choices[2], key: "n", width: 15 },
          { header: choices[3], key: "s", width: 15 },
          { header: choices[4], key: "ss", width: 20 },
          { header: "Total Jawaban", key: "total", width: 18 }
        ];

        for (let i = 1; i <= 30; i++) {
          const key = `P${i}`;
          const counts = choices.map(choice =>
            rows.filter(row => row[key] === choice).length
          );

          recapSheet.addRow({
            question: key,
            sts: counts[0],
            ts: counts[1],
            n: counts[2],
            s: counts[3],
            ss: counts[4],
            total: counts.reduce((a, b) => a + b, 0)
          });
        }

        recapSheet.insertRow(1, [
          "Pertanyaan",
          ...choices,
          "Total Jawaban"
        ]);

        recapSheet.views = [{ state: "frozen", ySplit: 1 }];

        // FORMAT HEADER
        for (const sheet of [dataSheet, recapSheet]) {
          sheet.getRow(1).height = 25;
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
            horizontal: "center"
          };
        }

        const buffer = await workbook.xlsx.writeBuffer();

        return new Response(buffer, {
          headers: {
            "Content-Type":
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition":
              'attachment; filename="hasil-kuesioner.xlsx"',
            "Cache-Control": "no-store"
          }
        });

      } catch (error) {
        return Response.json(
          {
            success: false,
            error: "Gagal membuat Excel: " + String(error?.message || error)
          },
          { status: 500 }
        );
      }
    }

    // =========================
    // SUBMIT KUESIONER
    // =========================
    if (url.pathname === "/api/submit") {
      if (request.method !== "POST") {
        return Response.json(
          { success: false, error: "Method not allowed" },
          { status: 405 }
        );
      }

      if (!env.DB) {
        return Response.json(
          { success: false, error: "Database binding DB belum dipasang." },
          { status: 500 }
        );
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
          if (
            typeof data[key] !== "string" ||
            !data[key].trim()
          ) {
            return Response.json(
              { success: false, error: `Data ${key} belum lengkap.` },
              { status: 400 }
            );
          }
        }

        const answers = [];

        for (let i = 1; i <= 30; i++) {
          const value = data[`P${i}`];

          if (
            ![
              "Sangat Tidak Setuju",
              "Tidak Setuju",
              "Netral",
              "Setuju",
              "Sangat Setuju"
            ].includes(value)
          ) {
            return Response.json(
              {
                success: false,
                error: `Jawaban pertanyaan ${i} tidak valid.`
              },
              { status: 400 }
            );
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

        return Response.json(
          {
            success: false,
            error: status === 409
              ? "Jawaban ini sudah pernah dikirim."
              : "Gagal menyimpan jawaban. Periksa tabel dan binding database."
          },
          { status }
        );
      }
    }

    return env.ASSETS.fetch(request);
  }
};
      
