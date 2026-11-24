/**
 * docx-engine.js -- Engine pembuat file .docx langsung di browser dari template
 * ber-tag Jinja/docxtpl ({{ var }}, {% for %}, {% if %}, {% tr for %}).
 *
 * Menggunakan:
 *   - JSZip    (membaca/menulis paket ZIP .docx)
 *   - Nunjucks (mengevaluasi tag Jinja2 pada XML Word)
 *   - Canvas   (mengonversi gambar WebP/lainnya ke PNG & membaca ukuran piksel)
 */
(function (global) {
  "use strict";

  const EMU_PER_MM = 36000;
  const EMU_PER_PX_96DPI = 9525; // 1 px @ 96 DPI = 9525 EMU

  // =========================================================================
  // 1. UTILITAS GAMBAR & XML
  // =========================================================================
  function escapeXml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  /**
   * Pastikan gambar berformat PNG atau JPEG (Word tidak mendukung WebP)
   * dan kembalikan { buffer, ext, mime }.
   */
  async function normalizeImageForWord(blob) {
    const type = (blob && blob.type) || "";
    if (type === "image/png") {
      return { buffer: await blob.arrayBuffer(), ext: "png", mime: "image/png" };
    }
    if (type === "image/jpeg" || type === "image/jpg") {
      return { buffer: await blob.arrayBuffer(), ext: "jpeg", mime: "image/jpeg" };
    }
    const url = URL.createObjectURL(blob);
    try {
      const img = await loadHtmlImage(url);
      const canvas = document.createElement("canvas");
      canvas.width = img.naturalWidth || img.width || 1;
      canvas.height = img.naturalHeight || img.height || 1;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0);
      const pngBlob = await new Promise((res) => canvas.toBlob(res, "image/png"));
      return { buffer: await pngBlob.arrayBuffer(), ext: "png", mime: "image/png" };
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  function loadHtmlImage(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = (e) => reject(new Error("Gagal memuat gambar: " + e));
      img.src = url;
    });
  }

  async function getImageNaturalSize(blob) {
    const url = URL.createObjectURL(blob);
    try {
      const img = await loadHtmlImage(url);
      return {
        width: img.naturalWidth || img.width || 300,
        height: img.naturalHeight || img.height || 300,
      };
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  function buildDrawingXml(rId, docPrId, name, cx, cy) {
    return (
      `<w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0"` +
      ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${docPrId}" name="${escapeXml(name)}"/>` +
      `<wp:cNvGraphicFramePr>` +
      `<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>` +
      `</wp:cNvGraphicFramePr>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:nvPicPr>` +
      `<pic:cNvPr id="${docPrId}" name="${escapeXml(name)}"/>` +
      `<pic:cNvPicPr/>` +
      `</pic:nvPicPr>` +
      `<pic:blipFill>` +
      `<a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
      `<a:stretch><a:fillRect/></a:stretch>` +
      `</pic:blipFill>` +
      `<pic:spPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `</pic:spPr>` +
      `</pic:pic>` +
      `</a:graphicData>` +
      `</a:graphic>` +
      `</wp:inline>` +
      `</w:drawing>`
    );
  }

  /**
   * Membangun XML tabel Word untuk lampiran Rekapitulasi Bahan Baku.
   * Menggunakan <w:tblHeader/> pada baris judul dan <w:cantSplit/> pada tiap baris
   * agar tabel rapi jika melintasi batas halaman cetak.
   */
  function buildBahanBakuTableXml(rows) {
    const border = `<w:tcBorders>` +
      `<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `</w:tcBorders>`;
    const cell = (text, w, bold) =>
      `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${border}</w:tcPr>` +
      `<w:p><w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:tc>`;
    const row = (cells, bold, isHeader = false) => {
      const trPr = `<w:trPr><w:cantSplit/>${isHeader ? "<w:tblHeader/>" : ""}</w:trPr>`;
      return `<w:tr>${trPr}` + cells.map((c, i) => cell(c, [700, 3200, 3200, 1400][i], bold)).join("") + `</w:tr>`;
    };

    let xml =
      `<w:tbl><w:tblPr><w:tblW w:w="8500" w:type="dxa"/>` +
      `<w:tblBorders>` +
      `<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `</w:tblBorders></w:tblPr>`;
    xml += row(["No", "Bahan Baku", "Produsen/Pemasok", "Asal"], true, true);
    rows.forEach((r) => {
      xml += row([String(r.no), r.nama_bahan, r.produsen, r.asal], false, false);
    });
    xml += `</w:tbl><w:p/>`;
    return xml;
  }

  // =========================================================================
  // 2. PRE-PROCESSING XML WORD (GABUNG RUN TERPECAH + NORMALISASI TAG JINJA)
  // =========================================================================
  function mergeSplitTagRuns(xml) {
    return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paraXml) => {
      const tMatches = [...paraXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)];
      if (tMatches.length <= 1) return paraXml;
      const joined = tMatches.map((m) => m[1]).join("");
      if (!/\{\{|\{%|\{#/.test(joined)) return paraXml;

      let first = true;
      return paraXml.replace(/<w:t\b([^>]*)>[\s\S]*?<\/w:t>/g, (full, attrs) => {
        if (first) {
          first = false;
          const cleanAttrs = attrs.includes("xml:space") ? attrs : attrs + ' xml:space="preserve"';
          return `<w:t${cleanAttrs}>${joined}</w:t>`;
        }
        return `<w:t${attrs}></w:t>`;
      });
    });
  }

  function decodeEntities(s) {
    return s
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }

  function stripInlineXml(s) {
    return decodeEntities(s.replace(/<[^>]+>/g, ""));
  }

  function preprocessDocxXml(xml) {
    let out = mergeSplitTagRuns(xml);

    // 1. Tag loop baris tabel docxtpl: {% tr for ... %} ... {% tr endfor %}
    out = out.replace(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g, (trXml) => {
      const mStart = trXml.match(/\{%-?\s*tr\s+(for\b[\s\S]*?)-?%\}/);
      if (mStart) {
        const expr = stripInlineXml(mStart[1]);
        const cleanedTr = trXml.replace(mStart[0], "");
        return `{% ${expr} %}${cleanedTr}`;
      }
      const mEnd = trXml.match(/\{%-?\s*tr\s+endfor\s*-?%\}/);
      if (mEnd) {
        const cleanedTr = trXml.replace(mEnd[0], "");
        return `${cleanedTr}{% endfor %}`;
      }
      return trXml;
    });

    // 2. Paragraf yang hanya berisi tag kontrol ({% ... %})
    out = out.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (pXml) => {
      const plain = stripInlineXml(pXml).trim();
      if (/^\{%-?[\s\S]*?-?%\}$/.test(plain)) {
        return plain;
      }
      return pXml;
    });

    // 3. Bersihkan XML di dalam {{ ... }} dan {% ... %}
    out = out.replace(/(\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\})/g, (tag) => stripInlineXml(tag));

    // 4. Placeholder khusus tabel Rekapitulasi Bahan Baku ({{ tabel_rekap }})
    out = out.replace(/<w:p\b[^>]*>[\s\S]*?\{\{\s*tabel_rekap\s*\}\}[\s\S]*?<\/w:p>/g, "{{ tabel_rekap | safe }}");

    return out;
  }

  // =========================================================================
  // 3. ESCAPE DATA KONTEKS & PENGUNCIAN TATA LETAK CETAK
  // =========================================================================
  function escapeContext(val) {
    if (val === null || val === undefined) return "";
    if (typeof val === "string") {
      if (val.startsWith("§§RAW:")) return val.slice(6);
      return escapeXml(val).replace(/\r?\n/g, "</w:t><w:br/><w:t xml:space=\"preserve\">");
    }
    if (Array.isArray(val)) return val.map(escapeContext);
    if (typeof val === "object") {
      const out = {};
      for (const k of Object.keys(val)) out[k] = escapeContext(val[k]);
      return out;
    }
    return val;
  }

  /**
   * Pastikan setiap baris tabel (<w:tr>) yang berisi gambar memiliki properti
   * <w:cantSplit/> agar gambar dan keterangannya tidak terbelah di antara dua halaman.
   */
  function enforceTableCantSplit(xmlString) {
    return xmlString.replace(/<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g, (match, inner) => {
      if (!inner.includes("§§IMG:")) return match;
      if (inner.includes("<w:trPr>")) {
        return match.replace("<w:trPr>", "<w:trPr><w:cantSplit/>");
      }
      return match.replace(/(<w:tr\b[^>]*>)/, "$1<w:trPr><w:cantSplit/></w:trPr>");
    });
  }

  // =========================================================================
  // 4. FUNGSI UTAMA GENERATE DOCX
  // =========================================================================
  async function generateDocx(templateBuffer, context, imageJobs) {
    const zip = await JSZip.loadAsync(templateBuffer);

    const env = new nunjucks.Environment(null, { autoescape: false, throwOnUndefined: false });
    env.addFilter("safe", (s) => s);

    const safeCtx = escapeContext(context);
    if (Array.isArray(context.rekap_bahan_baku) && context.rekap_bahan_baku.length > 0) {
      safeCtx.tabel_rekap = buildBahanBakuTableXml(context.rekap_bahan_baku);
    } else {
      safeCtx.tabel_rekap = "";
    }

    const relsPath = "word/_rels/document.xml.rels";
    let relsXml = await zip.file(relsPath).async("string");
    let contentTypesXml = await zip.file("[Content_Types].xml").async("string");

    const idNums = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
    let nextRelNum = (idNums.length ? Math.max(...idNums) : 10) + 1;
    let docPrCounter = 100;
    const addedExts = new Set();

    const xmlTargets = Object.keys(zip.files).filter(
      (p) => p === "word/document.xml" || /^word\/(header|footer)\d+\.xml$/.test(p)
    );

    for (const xmlPath of xmlTargets) {
      const rawXml = await zip.file(xmlPath).async("string");
      const prepared = preprocessDocxXml(rawXml);
      let rendered;
      try {
        rendered = env.renderString(prepared, safeCtx);
      } catch (err) {
        throw new Error(`Gagal mengevaluasi tag template pada ${xmlPath}: ${err.message}`);
      }

      // Kunci baris tabel yang berisi gambar agar tidak terpotong antar halaman
      if (xmlPath === "word/document.xml") {
        rendered = enforceTableCantSplit(rendered);
      }

      if (xmlPath === "word/document.xml" && imageJobs && imageJobs.size > 0) {
        for (const [token, job] of imageJobs.entries()) {
          if (!rendered.includes(token)) continue;
          const norm = await normalizeImageForWord(job.blob);
          const natural = await getImageNaturalSize(job.blob).catch(() => ({ width: 400, height: 300 }));

          let cx, cy;
          // Batas tinggi aman per sel/halaman A4 (105 mm untuk grid 2 kolom, 210 mm untuk 1 kolom penuh)
          const maxHeightMm = job.widthMm && job.widthMm <= 80 ? 105 : 210;
          const maxCy = Math.round(maxHeightMm * EMU_PER_MM);

          if (job.widthMm) {
            cx = Math.round(job.widthMm * EMU_PER_MM);
            cy = Math.round(cx * (natural.height / natural.width));
            if (cy > maxCy) {
              const ratio = maxCy / cy;
              cx = Math.round(cx * ratio);
              cy = maxCy;
            }
          } else {
            cx = Math.round(natural.width * EMU_PER_PX_96DPI);
            cy = Math.round(natural.height * EMU_PER_PX_96DPI);
            const maxCx = 165 * EMU_PER_MM;
            if (cx > maxCx) {
              const ratio = maxCx / cx;
              cx = Math.round(cx * ratio);
              cy = Math.round(cy * ratio);
            }
            if (cy > maxCy) {
              const ratio = maxCy / cy;
              cx = Math.round(cx * ratio);
              cy = maxCy;
            }
          }

          const rId = `rId${nextRelNum++}`;
          const docPrId = docPrCounter++;
          const mediaFile = `media/img_${docPrId}.${norm.ext}`;
          zip.file(`word/${mediaFile}`, norm.buffer);

          relsXml = relsXml.replace(
            "</Relationships>",
            `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${mediaFile}"/></Relationships>`
          );
          addedExts.add(norm.ext);

          const drawingXml = buildDrawingXml(rId, docPrId, `img_${docPrId}.${norm.ext}`, cx, cy);
          const replacement = `</w:t></w:r><w:r>${drawingXml}</w:r><w:r><w:t xml:space="preserve">`;
          rendered = rendered.split(token).join(replacement);
        }
      }

      zip.file(xmlPath, rendered);
    }

    for (const ext of addedExts) {
      const mime = ext === "png" ? "image/png" : "image/jpeg";
      const re = new RegExp(`Extension="${ext}"`, "i");
      if (!re.test(contentTypesXml)) {
        contentTypesXml = contentTypesXml.replace(
          "</Types>",
          `<Default Extension="${ext}" ContentType="${mime}"/></Types>`
        );
      }
    }

    zip.file(relsPath, relsXml);
    zip.file("[Content_Types].xml", contentTypesXml);

    return await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
    });
  }

  global.DocxEngine = {
    generateDocx,
  };
})(window);
