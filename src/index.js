export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/submit") {
      if (request.method !== "POST") {
        return Response.json({ success: false, error: "Method not allowed" }, { status: 405 });
      }
      if (!env.DB) {
        return Response.json({ success: false, error: "Database binding DB belum dipasang." }, { status: 500 });
      }
      try {
        const data = await request.json();
        const required = ["submission_id", "submitted_at", "gender", "age", "education"];
        for (const key of required) {
          if (typeof data[key] !== "string" || !data[key].trim()) {
            return Response.json({ success: false, error: `Data ${key} belum lengkap.` }, { status: 400 });
          }
        }
        const answers = [];
        for (let i = 1; i <= 30; i++) {
          const value = data[`P${i}`];
          if (!["Sangat Tidak Setuju", "Tidak Setuju", "Netral", "Setuju", "Sangat Setuju"].includes(value)) {
            return Response.json({ success: false, error: `Jawaban pertanyaan ${i} tidak valid.` }, { status: 400 });
          }
          answers.push(value);
        }
        const columns = ["submission_id", "submitted_at", "gender", "age", "education", ...Array.from({ length: 30 }, (_, i) => `P${i + 1}`)];
        const values = [data.submission_id, data.submitted_at, data.gender, data.age, data.education, ...answers];
        const placeholders = columns.map(() => "?").join(", ");
        await env.DB.prepare(
          `INSERT INTO responses (${columns.join(", ")}) VALUES (${placeholders})`
        ).bind(...values).run();

        return Response.json({ success: true });
      } catch (error) {
        const message = String(error?.message || error);
        const status = message.toLowerCase().includes("unique") ? 409 : 500;
        return Response.json({ success: false, error: status === 409 ? "Jawaban ini sudah pernah dikirim." : "Gagal menyimpan jawaban. Periksa tabel dan binding database." }, { status });
      }
    }

    return env.ASSETS.fetch(request);
  }
};
