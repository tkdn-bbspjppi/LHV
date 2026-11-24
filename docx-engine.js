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
  function decodeEntities(s) {
    return s
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }

  function stripInlineXml(s) {
    return decodeEntities(s.replace(/<[^>]+>/g, ""))
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/\u00A0/g, " ");
  }

  function normalizeJinjaTag(rawTag) {
    let cleanTag = stripInlineXml(rawTag);
    cleanTag = cleanTag.replace(/\.length\b/g, " | length");
    return cleanTag;
  }

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

  function preprocessDocxXml(xml) {
    let out = mergeSplitTagRuns(xml);

    // 1. Tag loop baris tabel docxtpl: {% tr for ... %} ... {% tr endfor %}
    out = out.replace(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g, (trXml) => {
      let currentTr = trXml;
      let prefix = "";
      let suffix = "";

      const mStart = currentTr.match(/\{%-?\s*tr\s+(for\b[\s\S]*?)-?%\}/);
      if (mStart) {
        const expr = stripInlineXml(mStart[1]);
        currentTr = currentTr.replace(mStart[0], "");
        prefix = `{% ${expr} %}`;
      }

      const mEnd = currentTr.match(/\{%-?\s*tr\s+endfor\s*-?%\}/);
      if (mEnd) {
        currentTr = currentTr.replace(mEnd[0], "");
        suffix = `{% endfor %}`;
      }

      return `${prefix}${currentTr}${suffix}`;
    });

    // 2. Proses setiap paragraf <w:p> secara terisolasi (tidak pernah lintas paragraf)
    out = out.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (pXml) => {
      const plain = stripInlineXml(pXml).trim();

      // Jika seluruh paragraf hanya berisi satu atau lebih tag kontrol {% ... %}
      if (/^(?:\{%-?[\s\S]*?-?%\}\s*)+$/.test(plain)) {
        return plain.replace(/\{%[\s\S]*?%\}/g, (t) => normalizeJinjaTag(t));
      }

      // Placeholder khusus tabel Rekapitulasi Bahan Baku ({{ tabel_rekap }})
      if (/\{\{\s*tabel_rekap\s*\}\}/.test(plain)) {
        return "{{ tabel_rekap | safe }}";
      }

      // Bersihkan tag Jinja di dalam teks paragraf ini saja
      return pXml.replace(/(\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\})/g, (tag) => normalizeJinjaTag(tag));
    });

    return out;
  }

  // =========================================================================
  // 3. ESCAPE DATA KONTEKS, SANITASI OPENXML & PENGUNCIAN TATA LETAK CETAK
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

  function ensureAllTemplateVarsDefined(preparedXml, ctxObj) {
    const keywords = new Set([
      "if", "elif", "else", "endif", "for", "in", "endfor", "set", "block", "endblock",
      "extends", "include", "import", "from", "as", "macro", "endmacro", "call", "endcall",
      "filter", "endfilter", "raw", "endraw", "and", "or", "not", "is", "true", "false",
      "none", "null", "undefined", "loop", "super", "self", "length", "safe"
    ]);
    const tags = preparedXml.match(/(\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\})/g) || [];
    for (const tag of tags) {
      const inner = tag.slice(2, -2);
      const idents = inner.match(/\b[a-zA-Z_][a-zA-Z0-9_]*\b/g) || [];
      for (const id of idents) {
        if (!keywords.has(id) && !(id in ctxObj)) {
          ctxObj[id] = "";
        }
      }
    }
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

  /**
   * Pastikan struktur XML valid menurut spesifikasi OpenXML Word:
   * - Setiap sel tabel <w:tc> wajib memiliki minimal 1 <w:p> di dalamnya.
   * - Setiap baris tabel <w:tr> wajib memiliki minimal 1 <w:tc> (jika kosong akibat loop, hapus baris).
   * - Hapus run teks kosong <w:t xml:space="preserve"></w:t> sisa pemecahan gambar.
   */
  function sanitizeWordXml(xmlString) {
    let out = xmlString;

    // 1. Hapus <w:tr> yang tidak memiliki satu pun <w:tc>
    out = out.replace(/<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g, (match, inner) => {
      return /<w:tc\b/.test(inner) ? match : "";
    });

    // 2. Pastikan setiap <w:tc> memiliki minimal satu <w:p> (syarat wajib skema Word)
    out = out.replace(/<w:tc\b[^>]*>([\s\S]*?)<\/w:tc>/g, (match, inner) => {
      if (/<w:p\b/.test(inner)) return match;
      return match.replace("</w:tc>", "<w:p/></w:tc>");
    });

    // 3. Bersihkan <w:r> kosong yang hanya berisi <w:t></w:t> tanpa teks
    out = out.replace(/<w:r>\s*<w:t(?:\s+xml:space="preserve")?><\/w:t>\s*<\/w:r>/g, "");

    return out;
  }

  /**
   * Validasi XML menggunakan DOMParser browser sebelum file .docx diunduh.
   */
  function validateXmlStructure(xmlString, xmlPath) {
    if (typeof DOMParser === "undefined") return;
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, "application/xml");
    const parseErr = doc.querySelector("parsererror");
    if (parseErr) {
      const msg = (parseErr.textContent || "").replace(/\s+/g, " ").trim().slice(0, 260);
      throw new Error(`Struktur XML tidak valid pada ${xmlPath}: ${msg}`);
    }
  }

  // =========================================================================
  // 4. FUNGSI UTAMA GENERATE DOCX
  // =========================================================================
  async function generateDocx(templateBuffer, context, imageJobs) {
    const zip = await JSZip.loadAsync(templateBuffer);

    const env = new nunjucks.Environment(null, { autoescape: false, throwOnUndefined: false });
    env.addFilter("safe", (s) => s);
    env.addFilter("length", (v) => {
      if (v === null || v === undefined) return 0;
      if (typeof v.length === "number") return v.length;
      if (typeof v === "object") return Object.keys(v).length;
      return 0;
    });

    const safeCtx = escapeContext(context);
    if (Array.isArray(context.rekap_bahan_baku) && context.rekap_bahan_baku.length > 0) {
      safeCtx.tabel_rekap = buildBahanBakuTableXml(context.rekap_bahan_baku);
    } else {
      safeCtx.tabel_rekap = "<w:p/>";
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

      const renderCtx = Object.assign({}, safeCtx);
      ensureAllTemplateVarsDefined(prepared, renderCtx);

      let rendered;
      try {
        rendered = env.renderString(prepared, renderCtx);
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
          // Sisipkan gambar di dalam <w:r> secara aman baik berada di dalam <w:t> maupun langsung di <w:p>
          rendered = rendered.replace(new RegExp(`(<w:t\\b[^>]*>)([^<]*)${token}([^<]*)(</w:t>)`, "g"), (m, tOpen, before, after, tClose) => {
            const partBefore = before ? `${tOpen}${before}${tClose}` : "";
            const partAfter = after ? `${tOpen}${after}${tClose}` : "";
            return `${partBefore}</w:r><w:r>${drawingXml}</w:r><w:r>${partAfter}`;
          });
          // Fallback jika token berada di luar <w:t>
          rendered = rendered.split(token).join(`<w:r>${drawingXml}</w:r>`);
        }
      }

      rendered = sanitizeWordXml(rendered);
      validateXmlStructure(rendered, xmlPath);
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
