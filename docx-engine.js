/**
 * docx-engine.js -- Engine pembuat file .docx langsung di browser dari template
 * ber-tag Jinja/docxtpl ({{ var }}, {% for %}, {% if %}, {% tr for %}).
 */
(function (global) {
  "use strict";

  const EMU_PER_MM = 36000;
  const EMU_PER_PX_96DPI = 9525;
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
      `<w:drawing xmlns:w="${W_NS}">` +
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
  // 2. PRE-PROCESSING XML WORD BERBASIS DOM
  // =========================================================================
  function cleanText(s) {
    return String(s || "")
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/\u00A0/g, " ");
  }

  function normalizeJinjaExpression(tagStr) {
    return cleanText(tagStr).replace(/\.length\b/g, " | length");
  }

  /**
   * Lepas pembungkus Content Control Word (<w:sdt> / <w:sdtContent>) sambil
   * mempertahankan seluruh isi di dalamnya agar tidak pernah terpotong oleh tag Jinja.
   */
  function unwrapContentControls(doc) {
    let sdts = Array.from(doc.getElementsByTagNameNS(W_NS, "sdt"));
    while (sdts.length > 0) {
      for (const sdt of sdts) {
        const parent = sdt.parentNode;
        if (!parent) continue;
        const sdtContent = sdt.getElementsByTagNameNS(W_NS, "sdtContent")[0];
        if (sdtContent) {
          while (sdtContent.firstChild) {
            parent.insertBefore(sdtContent.firstChild, sdt);
          }
        }
        parent.removeChild(sdt);
      }
      sdts = Array.from(doc.getElementsByTagNameNS(W_NS, "sdt"));
    }
  }

  function preprocessDocxDom(rawXml) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(rawXml, "application/xml");
    const parseErr = doc.querySelector("parsererror");
    if (parseErr) {
      throw new Error(`Template XML rusak: ${parseErr.textContent}`);
    }

    // 0. Lepas semua <w:sdt> agar tidak ada mismatch <w:sdtContent> vs <w:tc>
    unwrapContentControls(doc);

    // 1. Gabungkan teks terpecah (<w:t>) di dalam tiap paragraf <w:p> dan bersihkan tag Jinja
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    for (const p of paragraphs) {
      const tNodes = Array.from(p.getElementsByTagNameNS(W_NS, "t"));
      if (tNodes.length === 0) continue;

      const fullText = cleanText(tNodes.map((t) => t.textContent || "").join(""));
      if (!/\{\{|\{%|\{#/.test(fullText)) continue;

      if (/^\{\{\s*tabel_rekap\s*\}\}$/.test(fullText.trim())) {
        const placeholderNode = doc.createTextNode("{{ tabel_rekap | safe }}");
        p.parentNode.replaceChild(placeholderNode, p);
        continue;
      }

      const trimmed = fullText.trim();
      if (
        /^(?:\{%-?[\s\S]*?-?%\}\s*)+$/.test(trimmed) &&
        !/\{%-?\s*tr\b/.test(trimmed)
      ) {
        const normalizedCtrl = trimmed.replace(/\{%[\s\S]*?%\}/g, (t) => normalizeJinjaExpression(t));
        const ctrlNode = doc.createTextNode(normalizedCtrl);
        p.parentNode.replaceChild(ctrlNode, p);
        continue;
      }

      const normalizedParaText = fullText.replace(/(\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\})/g, (tag) =>
        normalizeJinjaExpression(tag)
      );

      tNodes[0].textContent = normalizedParaText;
      tNodes[0].setAttribute("xml:space", "preserve");
      for (let i = 1; i < tNodes.length; i++) {
        tNodes[i].textContent = "";
      }
    }

    // 2. Tangani tag baris tabel docxtpl: {% tr for ... %} dan {% tr endfor %}
    const rows = Array.from(doc.getElementsByTagNameNS(W_NS, "tr"));
    for (const tr of rows) {
      const tNodes = Array.from(tr.getElementsByTagNameNS(W_NS, "t"));
      let prefix = "";
      let suffix = "";

      for (const t of tNodes) {
        let txt = t.textContent || "";
        const mStart = txt.match(/\{%-?\s*tr\s+(for\b[\s\S]*?)-?%\}/);
        if (mStart) {
          prefix += `{% ${normalizeJinjaExpression(mStart[1])} %}`;
          txt = txt.replace(mStart[0], "");
        }
        const mEnd = txt.match(/\{%-?\s*tr\s+endfor\s*-?%\}/);
        if (mEnd) {
          suffix += `{% endfor %}`;
          txt = txt.replace(mEnd[0], "");
        }
        t.textContent = txt;
      }

      if (prefix) {
        tr.parentNode.insertBefore(doc.createTextNode(prefix), tr);
      }
      if (suffix) {
        tr.parentNode.insertBefore(doc.createTextNode(suffix), tr.nextSibling);
      }
    }

    return new XMLSerializer().serializeToString(doc);
  }

  // =========================================================================
  // 3. ESCAPE KONTEKS, INJEKSI GAMBAR & SANITASI AKHIR OPENXML
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

  function enforceTableCantSplit(xmlString) {
    return xmlString.replace(/<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g, (match, inner) => {
      if (!inner.includes("§§IMG:")) return match;
      if (inner.includes("<w:trPr>")) {
        return match.replace("<w:trPr>", "<w:trPr><w:cantSplit/>");
      }
      return match.replace(/(<w:tr\b[^>]*>)/, "$1<w:trPr><w:cantSplit/></w:trPr>");
    });
  }

  function replaceImageTokenInRuns(xmlString, token, drawingXml) {
    return xmlString.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (runXml) => {
      if (!runXml.includes(token)) return runXml;

      const rPrMatch = runXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
      const rPr = rPrMatch ? rPrMatch[0] : "";

      const textMatches = [...runXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)];
      const fullText = textMatches.map((m) => m[1]).join("");
      const parts = fullText.split(token);

      let outRuns = "";
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) {
          outRuns += `<w:r>${rPr}<w:t xml:space="preserve">${parts[i]}</w:t></w:r>`;
        }
        if (i < parts.length - 1) {
          outRuns += `<w:r>${drawingXml}</w:r>`;
        }
      }
      return outRuns;
    });
  }

  function sanitizeWordDom(xmlString, xmlPath) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, "application/xml");
    const parseErr = doc.querySelector("parsererror");
    if (parseErr) {
      const msg = (parseErr.textContent || "").replace(/\s+/g, " ").trim().slice(0, 260);
      throw new Error(`Struktur XML tidak valid pada ${xmlPath}: ${msg}`);
    }

    // 1. Hapus <w:tr> yang tidak memiliki <w:tc>
    const trList = Array.from(doc.getElementsByTagNameNS(W_NS, "tr"));
    for (const tr of trList) {
      if (tr.getElementsByTagNameNS(W_NS, "tc").length === 0) {
        tr.parentNode.removeChild(tr);
      }
    }

    // 2. Setiap <w:tc> wajib memiliki minimal satu <w:p> dan tidak boleh diakhiri <w:tbl> tanpa <w:p>
    const tcList = Array.from(doc.getElementsByTagNameNS(W_NS, "tc"));
    for (const tc of tcList) {
      if (tc.getElementsByTagNameNS(W_NS, "p").length === 0) {
        tc.appendChild(doc.createElementNS(W_NS, "w:p"));
      } else {
        const lastEl = tc.lastElementChild;
        if (lastEl && lastEl.localName === "tbl") {
          tc.appendChild(doc.createElementNS(W_NS, "w:p"));
        }
      }
    }

    return new XMLSerializer().serializeToString(doc);
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

    const usedRelIds = new Set([...relsXml.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
    let nextRelNum = 100;
    const allocRelId = () => {
      while (usedRelIds.has(`rId${nextRelNum}`)) nextRelNum++;
      const id = `rId${nextRelNum++}`;
      usedRelIds.add(id);
      return id;
    };

    const addedExts = new Set();

    const xmlTargets = Object.keys(zip.files).filter(
      (p) => p === "word/document.xml" || /^word\/(header|footer)\d+\.xml$/.test(p)
    );

    for (const xmlPath of xmlTargets) {
      const rawXml = await zip.file(xmlPath).async("string");

      const existingDocPrIds = [...rawXml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map((m) => parseInt(m[1], 10));
      let docPrCounter = (existingDocPrIds.length ? Math.max(...existingDocPrIds) : 100) + 10;

      const prepared = preprocessDocxDom(rawXml);
      const renderCtx = Object.assign({}, safeCtx);
      ensureAllTemplateVarsDefined(prepared, renderCtx);

      let rendered;
      try {
        rendered = env.renderString(prepared, renderCtx);
      } catch (err) {
        throw new Error(`Gagal mengevaluasi tag template pada ${xmlPath}: ${err.message}`);
      }

      if (xmlPath === "word/document.xml") {
        rendered = enforceTableCantSplit(rendered);
      }

      if (xmlPath === "word/document.xml" && imageJobs && imageJobs.size > 0) {
        for (const [token, job] of imageJobs.entries()) {
          if (!rendered.includes(token)) continue;
          const norm = await normalizeImageForWord(job.blob);
          const natural = await getImageNaturalSize(job.blob).catch(() => ({ width: 400, height: 300 }));

          let cx, cy;
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

          const rId = allocRelId();
          const docPrId = docPrCounter++;
          const mediaFile = `media/img_${docPrId}.${norm.ext}`;
          zip.file(`word/${mediaFile}`, norm.buffer);

          relsXml = relsXml.replace(
            "</Relationships>",
            `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${mediaFile}"/></Relationships>`
          );
          addedExts.add(norm.ext);

          const drawingXml = buildDrawingXml(rId, docPrId, `img_${docPrId}.${norm.ext}`, cx, cy);
          rendered = replaceImageTokenInRuns(rendered, token, drawingXml);
        }
      }

      rendered = sanitizeWordDom(rendered, xmlPath);
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
