/*
 * PATCH LMS - PROGRESS LINTAS BROWSER / PERANGKAT
 * Dipakai bersama Apps Script CAT + LMS yang sekarang.
 *
 * 1. Tambahkan fungsi lmsLoadProgress_() di bawah fungsi LMS lain.
 * 2. Di dalam doGet(e), SETELAH setup(); tambahkan router:
 *
 *   if (e && e.parameter && e.parameter.action === "lms_progress") {
 *     return lmsLoadProgress_(e.parameter.code, e.parameter.grade);
 *   }
 *
 * 3. Deploy sebagai versi baru Web App.
 */

function lmsLoadProgress_(rawCode, rawGrade) {
  try {
    const code = String(rawCode || "").trim().toUpperCase();
    const requestedGrade = String(rawGrade || "").trim().toUpperCase();

    if (!code) {
      return lmsJson_({ok:false, found:false, error:"Kode peserta LMS kosong"});
    }

    const peserta = getPesertaByCode(code);
    if (!peserta) {
      return lmsJson_({ok:false, found:false, error:"Kode peserta LMS tidak valid"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("LMS_PROGRESS");

    if (!sh || sh.getLastRow() <= 1) {
      return lmsJson_({
        ok:true, found:false, code:peserta.code, name:peserta.name,
        grade:requestedGrade, done:[], doneCount:0, total:0, percent:0
      });
    }

    const values = sh.getDataRange().getValues();

    // Baca dari bawah supaya selalu mengambil progress TERAKHIR.
    for (let i = values.length - 1; i >= 1; i--) {
      const rowCode = String(values[i][1] || "").trim().toUpperCase();
      if (rowCode !== peserta.code) continue;

      const rowGrade = String(values[i][3] || "").trim().toUpperCase();
      if (requestedGrade && rowGrade && rowGrade !== requestedGrade) continue;

      let done = [];
      try {
        const parsed = JSON.parse(values[i][7] || "[]");
        if (Array.isArray(parsed)) done = parsed.map(String);
      } catch (err) {
        done = [];
      }

      return lmsJson_({
        ok:true,
        found:true,
        code:peserta.code,
        name:peserta.name,
        grade:rowGrade || requestedGrade,
        done:done,
        doneCount:Number(values[i][5]) || done.length,
        total:Number(values[i][6]) || 0,
        percent:Number(values[i][4]) || 0,
        timestamp:values[i][0]
      });
    }

    return lmsJson_({
      ok:true, found:false, code:peserta.code, name:peserta.name,
      grade:requestedGrade, done:[], doneCount:0, total:0, percent:0
    });

  } catch (err) {
    return lmsJson_({ok:false, found:false, error:String(err)});
  }
}

/*
 * PERBAIKAN FRONTEND
 * Jika memakai lms-v8-persistence-wrapper.html, bagian ini tidak perlu
 * ditempel ke HTML karena wrapper sudah melakukan load server saat login.
 *
 * Alur setelah patch:
 * Browser A selesai 2 subbab
 * -> POST lms_progress
 * -> baris masuk LMS_PROGRESS
 * -> Browser B login kode yang sama
 * -> GET ?action=lms_progress&code=...
 * -> progress 2/29 dipulihkan.
 */
