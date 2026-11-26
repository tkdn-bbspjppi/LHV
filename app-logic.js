/**
 * app-logic.js -- Pengolah data form, gambar, PDF, nomor LHV,
 * penyimpanan lokal (IndexedDB), serta pemanggil engine Word.
 * Tidak membutuhkan backend/server (berjalan 100% di browser).
 */
(function (global) {
  "use strict";

  // Pastikan PDF.js worker diset ke CDN yang sama dengan index.html
  if (global.pdfjsLib && !global.pdfjsLib.GlobalWorkerOptions.workerSrc) {
    global.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }

  // =========================================================================
  // 1. NOMOR LHV OTOMATIS, FORMAT TANGGAL & TERBILANG INDONESIA
  // =========================================================================
  const BULAN_ID = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  const ROMAWI = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"];

  function formatTanggalIndo(dateStr) {
    if (!dateStr) return "";
    const parts = String(dateStr).split("-");
    if (parts.length !== 3) return dateStr;
    const [y, m, d] = parts;
    const bln = BULAN_ID[parseInt(m, 10)] || m;
    return `${parseInt(d, 10)} ${bln} ${y}`;
  }

  function angkaKeTerbilang(n) {
    const satuan = [
      "", "Satu", "Dua", "Tiga", "Empat", "Lima",
      "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"
    ];
    n = Math.floor(Math.abs(n));
    if (n === 0) return "Nol";
    if (n < 12) return satuan[n];
    if (n < 20) return angkaKeTerbilang(n - 10) + " Belas";
    if (n < 100) {
      return satuan[Math.floor(n / 10)] + " Puluh" + (n % 10 !== 0 ? " " + angkaKeTerbilang(n % 10) : "");
    }
    if (n < 200) {
      return "Seratus" + (n - 100 !== 0 ? " " + angkaKeTerbilang(n - 100) : "");
    }
    if (n < 1000) {
      return satuan[Math.floor(n / 100)] + " Ratus" + (n % 100 !== 0 ? " " + angkaKeTerbilang(n % 100) : "");
    }
    return String(n);
  }

  function nilaiKeTerbilangPersen(nilaiStr) {
    if (nilaiStr === null || nilaiStr === undefined) return "";
    const cleaned = String(nilaiStr).replace(/%/g, "").trim();
    if (!cleaned) return "";

    const num = parseFloat(cleaned.replace(",", "."));
    if (isNaN(num)) return "";

    const fixed = num.toFixed(2);
    const [bulatStr, desimalStr] = fixed.split(".");
    const bulat = parseInt(bulatStr, 10);

    let hasil = angkaKeTerbilang(bulat);
    if (desimalStr && desimalStr !== "00") {
      const d1 = parseInt(desimalStr[0], 10);
      const d2 = parseInt(desimalStr[1], 10);
      const kataDesimal = [
        d1 === 0 ? "Nol" : angkaKeTerbilang(d1),
        d2 === 0 ? "Nol" : angkaKeTerbilang(d2),
      ].join(" ");
      hasil += " Koma " + kataDesimal;
    }
    return hasil + " Persen";
  }

  function buatInisial(nama) {
    if (!nama || !String(nama).trim()) return "";
    return String(nama)
      .trim()
      .split(/\s+/)
      .map((w) => w[0].toUpperCase())
      .join("");
  }

  function kodeSkemaNomor(jenisLhv) {
    if (jenisLhv === "Kerjasama") return "TKDN-K";
    if (jenisLhv === "BMP") return "BMP";
    if (jenisLhv === "Jasa") return "TKDN-J";
    if (jenisLhv === "Gabungan") return "TKDN-G";
    return "TKDN";
  }

  function buildNoLhv(state) {
    const tgl = state.tanggalLhv || "";
    const parts = tgl.split("-");
    const tahun = parts[0] || String(new Date().getFullYear());
    const bulanNum = parseInt(parts[1], 10) || (new Date().getMonth() + 1);
    const romawi = ROMAWI[bulanNum] || "I";
    const kode = kodeSkemaNomor(state.jenisLhv);
    const id = (state.idBerkas || "").trim() || "00000";
    const inisial = buatInisial(state.namaVerifikator);
    return `${id}/LVI-BBSPJPPI/${kode}/${romawi}/${tahun}${inisial ? "/" + inisial : ""}`;
  }

  // Format nama file:
  //   TKDN (Produksi Sendiri / Kerjasama): ID Berkas_LHV TKDN Nama Perusahaan_Jenis Barang_Inisial Verifikator.docx
  //   BMP                                : ID Berkas_LHV BMP Nama Perusahaan_Inisial Verifikator.docx
  //   Jasa Industri                      : ID Berkas_LHV TKDN Jasa Nama Perusahaan_Jenis Barang/Jasa_Inisial Verifikator.docx
  //   Gabungan Barang & Jasa             : ID Berkas_LHV TKDN Gabungan Nama Perusahaan_Jenis Barang_Inisial Verifikator.docx
  // Bagian yang kosong dilewati (tanpa garis bawah ganda).
  const LABEL_FILE = {
    "Produksi Sendiri": "LHV TKDN",
    "Kerjasama": "LHV TKDN",
    "BMP": "LHV BMP",
    "Jasa": "LHV TKDN Jasa",
    "Gabungan": "LHV TKDN Gabungan",
  };

  function bersihkanNamaFile(str) {
    return String(str || "")
      .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "-")
      .replace(/\s+/g, " ")
      .trim();
  }

  function buildFileName(state) {
    const jenis = state.jenisLhv || "Produksi Sendiri";
    const label = LABEL_FILE[jenis] || "LHV TKDN";
    const idBerkas = bersihkanNamaFile(state.idBerkas);
    const perusahaan = bersihkanNamaFile(state.namaPerusahaan);
    const jenisBarang = jenis === "BMP" ? "" : bersihkanNamaFile(state.jenisBarang);
    const inisial = bersihkanNamaFile(buatInisial(state.namaVerifikator));

    const parts = [
      idBerkas,
      [label, perusahaan].filter(Boolean).join(" "),
      jenisBarang,
      inisial,
    ].filter(Boolean);

    let base = parts.join("_").replace(/[. ]+$/g, "");
    if (base.length > 150) base = base.slice(0, 150).replace(/[. _]+$/g, "");
    return `${base || label}.docx`;
  }

  // =========================================================================
  // 2. PEMROSESAN GAMBAR & PDF (AUTO-TRIM + PROPORTIONAL SCALING)
  // =========================================================================
  let tokenCounter = 0;
  function nextToken(prefix) {
    tokenCounter += 1;
    return `§§IMG:${prefix}_${tokenCounter}§§`;
  }

  // Pangkas margin putih/kosong di sekeliling kanvas agar mengikuti pola isi dokumen.
  // Pada mode formulir: abaikan 3% garis bingkai luar & 6% footer bawah ("Halaman X dari"),
  // serta berhenti saat menjumpai celah kosong vertikal lebar setelah tabel utama.
  function trimCanvasWhitespace(sourceCanvas, isFormulir = false, padding = 16) {
    const ctx = sourceCanvas.getContext("2d");
    const { width, height } = sourceCanvas;
    const data = ctx.getImageData(0, 0, width, height).data;

    const mX = isFormulir ? Math.round(width * 0.03) : 4;
    const mTop = isFormulir ? Math.round(height * 0.03) : 4;
    const mBot = isFormulir ? Math.round(height * 0.94) : height - 4;

    const rowInk = new Array(height).fill(0);
    let minX = width, maxX = 0, hasContent = false;

    for (let y = mTop; y < mBot; y++) {
      let count = 0;
      for (let x = mX; x < width - mX; x++) {
        const i = (y * width + x) * 4;
        if (data[i + 3] > 10 && (data[i] < 240 || data[i + 1] < 240 || data[i + 2] < 240)) {
          count++;
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          hasContent = true;
        }
      }
      rowInk[y] = count;
    }

    if (!hasContent) return sourceCanvas;

    let minY = mTop;
    while (minY < mBot && rowInk[minY] < 3) minY++;

    let maxY = minY;
    if (isFormulir) {
      let emptyRun = 0;
      const gapThreshold = Math.round(height * 0.055);
      for (let y = minY; y < mBot; y++) {
        if (rowInk[y] >= 3) {
          maxY = y;
          emptyRun = 0;
        } else {
          emptyRun++;
          if (emptyRun >= gapThreshold && maxY > minY + 100) break;
        }
      }
    } else {
      for (let y = mBot - 1; y >= minY; y--) {
        if (rowInk[y] >= 3) { maxY = y; break; }
      }
    }

    minX = Math.max(0, minX - padding);
    minY = Math.max(0, minY - padding);
    maxX = Math.min(width, maxX + padding);
    maxY = Math.min(height, maxY + padding);

    const w = maxX - minX;
    const h = maxY - minY;
    const trimmed = document.createElement("canvas");
    trimmed.width = w;
    trimmed.height = h;
    trimmed.getContext("2d").drawImage(sourceCanvas, minX, minY, w, h, 0, 0, w, h);
    return trimmed;
  }

  async function extractPdfPagesAsBlobs(file, isFormulir = false) {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const blobs = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: 2.0 });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext("2d");
      await page.render({ canvasContext: ctx, viewport }).promise;

      const cleanCanvas = trimCanvasWhitespace(canvas, isFormulir, 16);
      const blob = await new Promise((res) => cleanCanvas.toBlob(res, "image/png"));
      blobs.push(blob);
    }
    return blobs;
  }

  // Perkecil resolusi secara proporsional TANPA kotak persegi 1:1 (tanpa ruang kosong putih di kiri-kanan)
  async function normalizeImageProportional(blob, maxSidePx = 1200) {
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
    const img = await new Promise((resolve, reject) => {
      const im = new Image();
      im.onload = () => resolve(im);
      im.onerror = reject;
      im.src = dataUrl;
    });

    const scale = Math.min(1, maxSidePx / Math.max(img.width, img.height));
    const targetW = Math.round(img.width * scale);
    const targetH = Math.round(img.height * scale);

    const canvas = document.createElement("canvas");
    canvas.width = targetW;
    canvas.height = targetH;
    canvas.getContext("2d").drawImage(img, 0, 0, targetW, targetH);

    return new Promise((resolve) => {
      canvas.toBlob((b) => resolve(b || blob), "image/png");
    });
  }

  function isExcelFile(file) {
    if (!file) return false;
    const name = (file.name || "").toLowerCase();
    return (
      name.endsWith(".xlsx") ||
      name.endsWith(".xls") ||
      file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.type === "application/vnd.ms-excel"
    );
  }

  async function processSingleImage(fileOrBlob, customBlobs, widthMm, imageJobs, label, tick, maxHeightMm) {
    // Hasil crop manual: semua halaman dipakai (token digabung spasi, seperti rekap bahan baku)
    if (Array.isArray(customBlobs) && customBlobs.length > 0) {
      tick && tick(`Memproses gambar: ${label}...`);
      const tokens = [];
      for (const b of customBlobs) {
        const token = nextToken(label);
        imageJobs.set(token, { blob: b, widthMm, maxHeightMm });
        tokens.push(token);
      }
      return tokens.join(" ");
    }
    const source = fileOrBlob;
    if (!source) return "";
    tick && tick(`Memproses gambar: ${label}...`);
    let blob = source;
    if (source.type === "application/pdf") {
      const pages = await extractPdfPagesAsBlobs(source, false);
      if (pages.length === 0) return "";
      blob = pages[0];
    }
    const token = nextToken(label);
    imageJobs.set(token, { blob, widthMm, maxHeightMm });
    return token;
  }

  async function buildDynamicList(list, defaultWidthMm, shape, imageJobs, tick) {
    const result = [];
    for (const item of list || []) {
      // Ukuran khusus dari modal crop (cropRects[0].widthMm) menggantikan lebar bawaan layout
      const widthMm = (item.cropRects && item.cropRects[0] && item.cropRects[0].widthMm) || defaultWidthMm;
      const hasCustom = Array.isArray(item.customBlobs) && item.customBlobs.length > 0;
      if (!item.file && !hasCustom) continue;
      const noKet = !!item.tanpaKeterangan && shape !== "formulir";
      const keterangan = noKet ? "" : shape === "formulir" ? item.judul || "" : item.keterangan || "";

      // Prioritas 1: Hasil crop manual dari modal pratinjau
      if (hasCustom) {
        tick && tick(`Memuat potongan manual: ${item.file?.name || keterangan}...`);
        for (let idx = 0; idx < item.customBlobs.length; idx++) {
          let blob = item.customBlobs[idx];
          if (shape === "std") blob = await normalizeImageProportional(blob, 1200);
          const token = nextToken("dyn");
          imageJobs.set(token, { blob, widthMm });
          const caption = noKet
            ? ""
            : item.customBlobs.length > 1
              ? keterangan
                ? `${keterangan} (Halaman ${idx + 1})`
                : `Halaman ${idx + 1}`
              : keterangan;
          result.push({ judul: caption, keterangan: caption, gambar: token });
        }
        continue;
      }

      // Prioritas 2: Ekstraksi & auto-trim otomatis
      if (item.file.type === "application/pdf") {
        tick && tick(`Mengekstrak halaman PDF: ${item.file.name}...`);
        const pages = await extractPdfPagesAsBlobs(item.file, shape === "formulir");
        for (let idx = 0; idx < pages.length; idx++) {
          let blob = pages[idx];
          if (shape === "std") blob = await normalizeImageProportional(blob, 1200);
          const token = nextToken("dyn");
          imageJobs.set(token, { blob, widthMm });
          const caption = noKet ? "" : keterangan ? `${keterangan} (Halaman ${idx + 1})` : `Halaman ${idx + 1}`;
          result.push({ judul: caption, keterangan: caption, gambar: token });
        }
      } else if (isExcelFile(item.file)) {
        tick && tick(`Menggambar tabel dari Excel: ${item.file.name}...`);
        let blob = await ExcelRender.renderExcelToImage(item.file);
        if (shape === "std") blob = await normalizeImageProportional(blob, 1200);
        const token = nextToken("dyn");
        imageJobs.set(token, { blob, widthMm });
        result.push({ judul: keterangan, keterangan: keterangan, gambar: token });
      } else {
        let blob = item.file;
        tick && tick(`Memproses berkas: ${item.file.name || keterangan}...`);
        if (shape === "std") blob = await normalizeImageProportional(blob, 1200);
        const token = nextToken("dyn");
        imageJobs.set(token, { blob, widthMm });
        result.push({ judul: keterangan, keterangan: keterangan, gambar: token });
      }
    }
    return result;
  }

  function groupIntoRows(list, perBaris) {
    const out = [];
    if (perBaris === 1) {
      for (let i = 0; i < list.length; i++) {
        out.push({
          gambar_1: list[i].gambar,
          keterangan_1: list[i].keterangan || "",
          gambar_2: "",
          keterangan_2: "",
        });
      }
      return out;
    }
    for (let i = 0; i < list.length; i += 2) {
      const a = list[i];
      const b = list[i + 1];
      out.push({
        gambar_1: a.gambar,
        keterangan_1: a.keterangan || "",
        gambar_2: b ? b.gambar : "",
        keterangan_2: b ? b.keterangan || "" : "",
      });
    }
    return out;
  }

  // =========================================================================
  // 3. KATEGORI DOKUMEN DINAMIS
  // =========================================================================
  const CATEGORY_CONFIG = [
    ["formulirVerifikasi", "list_formulir", 165, "formulir", false],
    ["fileBuktiPabrik", "list_bukti_pabrik", 70, "std", true],
    ["fileBom", "list_bom", 70, "std", true],
    ["fileSertifikatTkdn", "list_sertifikat_tkdn", 70, "std", true],
    ["fileBuktiBeli", "list_bukti_beli", 70, "std", true],
    ["fileKtp", "list_ktp", 70, "std", true],
    ["fileBuktiKerjasama", "list_bukti_kerjasama", 70, "std", true],
    ["fileTenagaKerjaBmp", "list_tenaga_kerja_bmp", 70, "std", true],
    ["fileInvestasiBmp", "list_investasi_bmp", 70, "std", true],
    ["fileKemitraanBmp", "list_kemitraan_bmp", 70, "std", true],
    ["fileSubstitusiBmp", "list_substitusi_bmp", 70, "std", true],
    ["fileMesinDnBmp", "list_mesin_dn_bmp", 70, "std", true],
    ["fileLokasiBmp", "list_lokasi_bmp", 70, "std", true],
    ["fileI40Bmp", "list_i40_bmp", 70, "std", true],
    ["fileSdmBmp", "list_sdm_bmp", 70, "std", true],
    ["fileSertifikatBmp", "list_sertifikat_bmp", 70, "std", true],
    ["fileHijauBmp", "list_hijau_bmp", 70, "std", true],
    ["fileEksporBmp", "list_ekspor_bmp", 70, "std", true],
    ["fileMerekDnBmp", "list_merek_dn_bmp", 70, "std", true],
    ["fileEsgBmp", "list_esg_bmp", 70, "std", true],
    ["fileAwardsBmp", "list_awards_bmp", 70, "std", true],
    ["fileSiinasBmp", "list_siinas_bmp", 70, "std", true],
    ["fileDokumenPengembangan", "list_dokumen_pengembangan", 70, "std", true],
    ["fileAktaPendirian", "list_akta_pendirian", 70, "std", true],
    ["fileNibRba", "list_nib_rba", 70, "std", true],
    ["fileSertifikatStandar", "list_sertifikat_standar", 70, "std", true],
    ["fileIzinUsaha", "list_izin_usaha", 70, "std", true],
    ["fileNpwpLampiran", "list_npwp_lampiran", 70, "std", true],
    ["fileSertifikatMerek", "list_sert_merek", 70, "std", true],
    ["fileSertifikatProduk", "list_sert_produk", 70, "std", true],
    ["fileNie", "list_nie", 70, "std", true],
    ["fileBpom", "list_bpom", 70, "std", true],
    ["fileFotoProduk", "list_foto_produk", 70, "std", true],
    ["fileFotoBahanBaku", "list_foto_bahan_baku", 70, "std", true],
    ["fileInvoiceBahanBaku", "list_invoice_bahan_baku", 70, "std", true],
    ["fileAlurProsesLampiran", "list_alur_proses_lampiran", 70, "std", true],
    ["fileDaftarGaji", "list_daftar_gaji", 70, "std", true],
    ["fileSampelKtp", "list_sampel_ktp", 70, "std", true],
    ["fileStrukturPabrik", "list_struktur_pabrik", 70, "std", true],
    ["fileFotoMesin", "list_foto_mesin", 70, "std", true],
    ["fileDaftarPenyusutan", "list_daftar_penyusutan", 70, "std", true],
    ["fileBuktiListrik", "list_bukti_listrik", 70, "std", true],
    ["fileAktaSewa", "list_akta_sewa", 70, "std", true],
    ["fileGeotagging", "list_geotagging", 70, "std", true],
  ];

  const TEMPLATE_MAP = {
    "Produksi Sendiri": "templates/template_lhv_produksi_sendiri.docx",
    "Kerjasama": "templates/template_lhv_kerjasama.docx",
    "BMP": "templates/template_lhv_bmp.docx",
    "Jasa": "templates/template_lhv_jasa.docx",
    "Gabungan": "templates/template_lhv_gabungan.docx",
  };

  // =========================================================================
  // 4. GENERATOR UTAMA
  // =========================================================================
  async function generateAndDownload(state, onProgress) {
    tokenCounter = 0;
    const imageJobs = new Map();

    const layoutChoice = String(state.layoutGambar || "2");
    const stdWidthMm = layoutChoice === "1" ? 130 : layoutChoice === "4" ? 45 : 70;
    const perBaris = layoutChoice === "1" ? 1 : 2;

    const activeCats = CATEGORY_CONFIG.filter(([key]) =>
      (state[key] || []).some((it) => it && (it.file || (it.customBlobs && it.customBlobs.length > 0)))
    );
    const totalSteps = 1 + activeCats.length + 3 + 1;
    let doneSteps = 0;
    const step = (msg) => {
      doneSteps += 1;
      const pct = Math.min(99, Math.round((doneSteps / totalSteps) * 100));
      onProgress && onProgress(msg, pct);
    };

    const tick = (msg) => {
      const pct = Math.min(99, Math.round((doneSteps / totalSteps) * 100));
      onProgress && onProgress(msg, pct);
    };

    const jenisLhv = state.jenisLhv || "Produksi Sendiri";
    const templateUrl = TEMPLATE_MAP[jenisLhv] || TEMPLATE_MAP["Produksi Sendiri"];
    const tplResp = await fetch(templateUrl);
    if (!tplResp.ok) {
      throw new Error(`Template tidak ditemukan (${templateUrl} -> HTTP ${tplResp.status})`);
    }
    const templateBuffer = await tplResp.arrayBuffer();
    step(`Memuat template ${jenisLhv}...`);

    const logoToken = await processSingleImage(state.fileLogo, null, 20, imageJobs, "logo", tick);
    const ttdToken = await processSingleImage(state.fileTtdVerifikator, null, 20, imageJobs, "ttd", tick);
    const strukturToken = await processSingleImage(
      state.fileStruktur, state.customBlobsStruktur, (state.widthOverrides && state.widthOverrides.struktur) || 70, imageJobs, "struktur", tick
    );
    const alurToken = await processSingleImage(
      state.fileAlurProduksi, state.customBlobsAlurProduksi, (state.widthOverrides && state.widthOverrides.alurProduksi) || 70, imageJobs, "alur", tick
    );
    const strukturIndustriToken = await processSingleImage(
      state.fileStrukturIndustri, state.customBlobsStrukturIndustri, (state.widthOverrides && state.widthOverrides.strukturIndustri) || 70, imageJobs, "struktur_ind", tick
    );
    const coverToken = await processSingleImage(state.fileCover, null, 210, imageJobs, "cover", tick, 297);
    const fotoBarangToken = await processSingleImage(state.fileFotoBarang, null, 70, imageJobs, "foto_barang", tick);
    const fotoProdukUtamaToken = await processSingleImage(state.fileFotoProdukUtama, null, 70, imageJobs, "foto_produk_utama", tick);
    step("Memproses cover, logo & gambar utama...");

    const dynamicContext = {};
    const flatContext = {}; // daftar datar {judul, keterangan, gambar} per kategori (untuk template yang mem-batch sendiri)
    for (const [stateKey, tplKey, widthMm, shape, groupTwo] of CATEGORY_CONFIG) {
      const list = state[stateKey] || [];
      const hasFile = list.some((it) => it && (it.file || (it.customBlobs && it.customBlobs.length > 0)));
      const effectiveWidth = shape === "std" ? stdWidthMm : widthMm;
      const built = await buildDynamicList(list, effectiveWidth, shape, imageJobs, tick);
      flatContext[tplKey] = built;
      dynamicContext[tplKey] = groupTwo ? groupIntoRows(built, perBaris) : built;
      if (hasFile) step(`Memproses ${tplKey}...`);
    }

    const acuanFiltered = (state.acuanPeraturan || [])
      .filter((item) => {
        if (item.id === 6 && !(state.permenperin || "").includes("35 Tahun 2025")) return false;
        return Boolean(item.aturan && item.aturan.trim());
      })
      .map((item) => item.aturan.trim());

    const aspekChecked = (state.aspekBmp || [])
      .filter((a) => a.checked)
      .map((a) => a.teks);

    const rekapClean = (state.rekapBahanBaku || [])
      .filter((r) => (r.nama_bahan && r.nama_bahan.trim()) || (r.produsen && r.produsen.trim()))
      .map((r, i) => ({
        no: i + 1,
        nama_bahan: r.nama_bahan || "",
        produsen: r.produsen || "",
        asal: r.asal || "DN",
      }));

    let rekapToken = "";
    if (state.modeRekapBahanBaku === "upload" && (state.fileRekapBahanBaku || (state.customBlobsRekap && state.customBlobsRekap.length > 0))) {
      tick("Memproses dokumen rekapitulasi bahan baku...");
      const rekapList = await buildDynamicList(
        [{ file: state.fileRekapBahanBaku, customBlobs: state.customBlobsRekap, keterangan: "" }],
        (state.widthOverrides && state.widthOverrides.rekapBahanBaku) || 165,
        "formulir",
        imageJobs,
        tick
      );
      rekapToken = rekapList.map((r) => r.gambar).join(" ");
    }

    // ---- Alias nama variabel untuk template gaya lama (mis. template_lhv_produksi_sendiri.docx) ----
    // Template tsb mem-batch sendiri: {% for baris in xxx | batch(2, None) %} + baris[0].gambar / baris[1].gambar.
    // Untuk layout "1 foto per baris", sisipkan null setelah tiap item agar tiap batch hanya berisi 1 foto.
    const forBatch = (list) => {
      if (perBaris !== 1) return list;
      const out = [];
      list.forEach((it) => { out.push(it, null); });
      return out;
    };
    const BATCHED_ALIASES = {
      akta_perusahaan: "list_akta_pendirian", akta_sewa: "list_akta_sewa", bpom: "list_bpom",
      bukti_beli: "list_bukti_beli", bukti_listrik: "list_bukti_listrik", bukti_pabrik: "list_bukti_pabrik",
      daftar_gaji: "list_daftar_gaji", daftar_penyusutan: "list_daftar_penyusutan",
      foto_bahan_baku: "list_foto_bahan_baku", foto_mesin: "list_foto_mesin", gambar_bom: "list_bom",
      geotagging: "list_geotagging", invoice_bahan_baku: "list_invoice_bahan_baku",
      izin_usaha: "list_izin_usaha", ktp_karyawan: "list_ktp", nib_rba: "list_nib_rba", nie: "list_nie",
      sampel_ktp: "list_sampel_ktp", sert_merek: "list_sert_merek", sert_produk: "list_sert_produk",
      sert_standar: "list_sertifikat_standar", sertifikat_tkdn: "list_sertifikat_tkdn",
    };
    const FLAT_ALIASES = {
      alur_proses_lampiran: "list_alur_proses_lampiran", npwp_lampiran: "list_npwp_lampiran",
      struktur_pabrik: "list_struktur_pabrik", formulir_tkdn: "list_formulir",
    };
    const aliasContext = {};
    for (const [alias, key] of Object.entries(BATCHED_ALIASES)) aliasContext[alias] = forBatch(flatContext[key] || []);
    for (const [alias, key] of Object.entries(FLAT_ALIASES)) aliasContext[alias] = flatContext[key] || [];

    // "foto_produk" dipakai template sebagai gambar tunggal ({{ foto_produk }}) SEKALIGUS daftar ({% for ... in foto_produk %}):
    // daftar yang bila dicetak langsung menghasilkan token gambar utama.
    const fotoProdukSingle = fotoProdukUtamaToken || (flatContext.list_foto_produk?.[0]?.gambar || "");
    const fotoProdukList = forBatch(flatContext.list_foto_produk || []).slice();
    Object.defineProperty(fotoProdukList, "toString", { value: () => fotoProdukSingle, enumerable: false });
    aliasContext.foto_produk = fotoProdukList;

    const noLhv = buildNoLhv(state);
    const context = {
      ...aliasContext,
      ...dynamicContext,
      ...aliasContext,
      cover_laporan: coverToken,
      struktur_organisasi: strukturToken,
      alur_produksi: alurToken,
      tanggal_dibuat: formatTanggalIndo(state.tanggalLhv),
      daftar_acuan_peraturan: acuanFiltered.map((a) => ({ aturan: a })),
      no_lhv: noLhv,
      nama_perusahaan: state.namaPerusahaan || "",
      tanggal_lhv: formatTanggalIndo(state.tanggalLhv),
      id_berkas: state.idBerkas || "",
      permenperin: state.permenperin || "",
      tgl_verifikasi_dok: formatTanggalIndo(state.tglVerifikasiDok),
      tgl_verifikasi_lapangan: formatTanggalIndo(state.tglVerifikasiLapangan),
      kbli: state.kbli || "",
      kapasitas_produksi: state.kapasitasProduksi || "",
      jenis_barang: state.jenisBarang || "",
      tipe_barang: state.tipeBarang || "",
      spesifikasi_barang: state.spesifikasiBarang || "",
      kode_hs: state.kodeHs || "",
      merek_barang: state.merekBarang || "",
      kelompok_barang: state.kelompokBarang || "",
      nilai_tkdn: state.nilaiTkdn || "",
      terbilang_tkdn: state.terbilangTkdn || "",
      nilai_brainware: state.nilaiBrainware || "",
      terbilang_brainware: state.terbilangBrainware || "",
      nilai_bmp: state.nilaiBmp || "",
      terbilang_bmp: state.terbilangBmp || "",
      skala_perusahaan: state.skalaPerusahaan || "",
      no_izin: state.noIzin || "",
      nama_perusahaan_industri: state.namaPerusahaanIndustri || "",
      alamat_perusahaan_industri: state.alamatPerusahaanIndustri || "",
      skala_perusahaan_industri: state.skalaPerusahaanIndustri || "",
      no_izin_perusahaan_industri: state.noIzinPerusahaanIndustri || "",
      alamat_kantor: state.alamatKantor || "",
      telepon_kantor: state.teleponKantor || "",
      fax_kantor: state.faxKantor || "",
      email_kantor: state.emailKantor || "",
      website_kantor: state.websiteKantor || "",
      status_kantor: state.statusKantor || "",
      pic_kantor: state.picKantor || "",
      akta_kantor: state.aktaKantor || "",
      npwp_kantor: state.npwpKantor || "",
      alamat_pabrik: state.alamatPabrik || "",
      telepon_pabrik: state.teleponPabrik || "",
      fax_pabrik: state.faxPabrik || "",
      email_pabrik: state.emailPabrik || "",
      website_pabrik: state.websitePabrik || "",
      status_pabrik: state.statusPabrik || "",
      pic_pabrik: state.picPabrik || "",
      akta_pabrik: state.aktaPabrik || "",
      npwp_pabrik: state.npwpPabrik || "",
      telepon_kantor_industri: state.teleponKantorIndustri || "",
      fax_kantor_industri: state.faxKantorIndustri || "",
      email_kantor_industri: state.emailKantorIndustri || "",
      website_kantor_industri: state.websiteKantorIndustri || "",
      status_kantor_industri: state.statusKantorIndustri || "",
      pic_kantor_industri: state.picKantorIndustri || "",
      akta_kantor_industri: state.aktaKantorIndustri || "",
      npwp_kantor_industri: state.npwpKantorIndustri || "",
      nama_verifikator: state.namaVerifikator || "",
      nip_verifikator: state.nipVerifikator || "",
      pejabat_mengetahui: state.pejabatMengetahui || "",
      nama_pejabat: state.namaPejabat || "",
      nip_pejabat: state.nipPejabat || "",
      acuan_peraturan: acuanFiltered,
      aspek_bmp: aspekChecked,
      rekap_bahan_baku: rekapClean,
      mode_rekap_bahan_baku: state.modeRekapBahanBaku || "tabel",
      rekap_bahan_baku_gambar: rekapToken,
      logo_perusahaan: logoToken,
      ttd_verifikator: ttdToken,
      gambar_struktur: strukturToken,
      gambar_alur_produksi: alurToken,
      gambar_struktur_industri: strukturIndustriToken,
      gambar_cover: coverToken,
      foto_barang: fotoBarangToken,
      // foto_produk: lihat aliasContext (daftar dengan toString = token gambar utama)
      foto_produk_utama: fotoProdukSingle,
    };

    step("Menyusun dokumen Word (.docx)...");
    const docxBlob = await DocxEngine.generateDocx(templateBuffer, context, imageJobs);

    step("Menyiapkan unduhan...");
    const filename = buildFileName(state);

    const url = URL.createObjectURL(docxBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 1500);

    return filename;
  }

  // =========================================================================
  // 5. PENYIMPANAN PROYEK LOKAL (INDEXEDDB)
  // =========================================================================
  const DB_NAME = "lhv_generator_db";
  const DB_VERSION = 5;
  const STORE_PROJECTS = "projects";
  const STORE_DRAFT = "draft";

  let dbPromise = null;
  function openDb() {
    if (!dbPromise) {
      dbPromise = openDbRaw().catch((e) => { dbPromise = null; throw e; });
    }
    return dbPromise;
  }

  function openDbRaw() {
    return new Promise((resolve, reject) => {
      if (!global.indexedDB) {
        reject(new Error("IndexedDB tidak didukung oleh browser ini."));
        return;
      }
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE_PROJECTS)) {
          db.createObjectStore(STORE_PROJECTS, { keyPath: "id", autoIncrement: true });
        }
        if (!db.objectStoreNames.contains(STORE_DRAFT)) {
          db.createObjectStore(STORE_DRAFT, { keyPath: "key" });
        }
      };
      req.onsuccess = () => {
        const db = req.result;
        db.onversionchange = () => { db.close(); dbPromise = null; };
        resolve(db);
      };
      req.onerror = () => {
        // Jika versi di browser masih lebih tinggi, buka tanpa parameter versi
        if (req.error && req.error.name === "VersionError") {
          const fallbackReq = indexedDB.open(DB_NAME);
          fallbackReq.onsuccess = () => resolve(fallbackReq.result);
          fallbackReq.onerror = () => reject(fallbackReq.error);
        } else {
          reject(req.error);
        }
      };
    });
  }
  const MAX_PROJECTS = 10;

  async function dbSaveProject(project) {
    const db = await openDb();
    const savedId = await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_PROJECTS, "readwrite");
      const req = project.id
        ? tx.objectStore(STORE_PROJECTS).put(project)
        : tx.objectStore(STORE_PROJECTS).add(project);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    // Batasi riwayat agar tidak melewati quota browser (File/Blob tersimpan utuh di tiap proyek)
    try {
      const all = await dbListProjects();
      for (const old of all.slice(MAX_PROJECTS)) await dbDeleteProject(old.id);
    } catch (e) {
      console.warn("Gagal merapikan riwayat proyek:", e);
    }
    return savedId;
  }

  async function dbListProjects() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_PROJECTS, "readonly");
      const req = tx.objectStore(STORE_PROJECTS).getAll();
      req.onsuccess = () => resolve((req.result || []).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0)));
      req.onerror = () => reject(req.error);
    });
  }

  async function dbDeleteProject(id) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_PROJECTS, "readwrite");
      const req = tx.objectStore(STORE_PROJECTS).delete(id);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  }

  async function dbSaveDraft(stateObj) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_DRAFT, "readwrite");
      const req = tx.objectStore(STORE_DRAFT).put({ key: "current", savedAt: Date.now(), state: stateObj });
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  }

  async function dbLoadDraft() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_DRAFT, "readonly");
      const req = tx.objectStore(STORE_DRAFT).get("current");
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }

  async function dbClearDraft() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_DRAFT, "readwrite");
      const req = tx.objectStore(STORE_DRAFT).delete("current");
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  }

  global.LHVLogic = {
    generateAndDownload,
    buildNoLhv,
    buildFileName,
    formatTanggalIndo,
    nilaiKeTerbilangPersen,
    dbSaveProject,
    dbListProjects,
    dbDeleteProject,
    dbSaveDraft,
    dbLoadDraft,
    dbClearDraft,
  };
})(window);
