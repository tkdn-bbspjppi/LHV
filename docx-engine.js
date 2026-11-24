/**
 * docx-engine.js -- Engine pembuat file .docx langsung di browser dari template
 * ber-tag Jinja/docxtpl ({{ var }}, {% for %}, {% if %}, {% tr for %}, {% tc for %}, {% p if %}).
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
      `<w:tbl xmlns:w="${W_NS}"><w:tblPr><w:tblW w:w="8500" w:type="dxa"/>` +
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
    xml += `</w:tbl>`;
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

  function findAncestor(node, localName) {
    let cur = node && node.parentNode;
    while (cur) {
      if (cur.localName === localName && cur.namespaceURI === W_NS) return cur;
      cur = cur.parentNode;
    }
    return null;
  }

  function preprocessDocxDom(rawXml) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(rawXml, "application/xml");
    const parseErr = doc.querySelector("parsererror");
    if (parseErr) {
      throw new Error(`Template XML rusak: ${parseErr.textContent}`);
    }

    unwrapContentControls(doc);

    // Tahap 1: Gabungkan run terpecah <w:t> dalam tiap paragraf <w:p>
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    for (const p of paragraphs) {
      const tNodes = Array.from(p.getElementsByTagNameNS(W_NS, "t"));
      if (tNodes.length === 0) continue;

      const fullText = cleanText(tNodes.map((t) => t.textContent || "").join(""));
      if (!/\{\{|\{%|\{#/.test(fullText)) continue;

      if (/^\{\{\s*tabel_rekap\s*\}\}$/.test(fullText.trim())) {
        p.parentNode.replaceChild(doc.createTextNode("§§TABEL_REKAP§§"), p);
        continue;
      }

      const normalized = fullText.replace(/(\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\})/g, (tag) =>
        normalizeJinjaExpression(tag)
      );

      tNodes[0].textContent = normalized;
      tNodes[0].setAttribute("xml:space", "preserve");
      for (let i = 1; i < tNodes.length; i++) {
        tNodes[i].textContent = "";
      }
    }

    // Tahap 2: Ekstrak tag kontrol docxtpl (tr, tc, p, r)
    const allT = Array.from(doc.getElementsByTagNameNS(W_NS, "t"));
    for (const t of allT) {
      let txt = t.textContent || "";
      if (!txt.includes("{%")) continue;

      const trEl = findAncestor(t, "tr");
      if (trEl) {
        txt = txt.replace(/\{%-?\s*tr\s+(for\b[\s\S]*?|if\b[\s\S]*?)-?%\}/g, (_, expr) => {
          trEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), trEl);
          return "";
        });
        txt = txt.replace(/\{%-?\s*tr\s+(endfor|endif|else|elif\b[\s\S]*?)-?%\}/g, (_, expr) => {
          trEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), trEl.nextSibling);
          return "";
        });
      }

      const tcEl = findAncestor(t, "tc");
      if (tcEl) {
        txt = txt.replace(/\{%-?\s*tc\s+(for\b[\s\S]*?|if\b[\s\S]*?)-?%\}/g, (_, expr) => {
          tcEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), tcEl);
          return "";
        });
        txt = txt.replace(/\{%-?\s*tc\s+(endfor|endif|else|elif\b[\s\S]*?)-?%\}/g, (_, expr) => {
          tcEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), tcEl.nextSibling);
          return "";
        });
      }

      const pEl = findAncestor(t, "p");
      if (pEl) {
        txt = txt.replace(/\{%-?\s*p\s+(for\b[\s\S]*?|if\b[\s\S]*?)-?%\}/g, (_, expr) => {
          pEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), pEl);
          return "";
        });
        txt = txt.replace(/\{%-?\s*p\s+(endfor|endif|else|elif\b[\s\S]*?)-?%\}/g, (_, expr) => {
          pEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), pEl.nextSibling);
          return "";
        });
      }

      const rEl = findAncestor(t, "r");
      if (rEl) {
        txt = txt.replace(/\{%-?\s*r\s+(for\b[\s\S]*?|if\b[\s\S]*?)-?%\}/g, (_, expr) => {
          rEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), rEl);
          return "";
        });
        txt = txt.replace(/\{%-?\s*r\s+(endfor|endif|else|elif\b[\s\S]*?)-?%\}/g, (_, expr) => {
          rEl.parentNode.insertBefore(doc.createTextNode(`{% ${expr.trim()} %}`), rEl.nextSibling);
          return "";
        });
      }

      t.textContent = txt;
    }

    // Tahap 3: Paragraf yang hanya berisi tag kontrol dipindahkan ke luar <w:p>
    const paragraphsAfter = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    for (const p of paragraphsAfter) {
      const tNodes = Array.from(p.getElementsByTagNameNS(W_NS, "t"));
      const fullText = cleanText(tNodes.map((t) => t.textContent || "").join("")).trim();
      if (fullText && /^(?:\{%-?[\s\S]*?-?%\}\s*)+$/.test(fullText)) {
        p.parentNode.replaceChild(doc.createTextNode(fullText), p);
      }
    }

    // Tahap 4: Pisahkan tag kontrol {% ... %} yang bercampur teks dalam satu <w:r>
    const runsWithCtrl = Array.from(doc.getElementsByTagNameNS(W_NS, "r"));
    for (const r of runsWithCtrl) {
      const tNodes = Array.from(r.getElementsByTagNameNS(W_NS, "t"));
      if (tNodes.length === 0) continue;
      const combined = tNodes.map((t) => t.textContent || "").join("");
      if (!combined.includes("{%")) continue;

      const rPr = r.getElementsByTagNameNS(W_NS, "rPr")[0];
      const segments = combined.split(/(\{%[\s\S]*?%\})/g).filter(Boolean);
      const frag = doc.createDocumentFragment();

      for (const seg of segments) {
        if (/^\{%[\s\S]*?%\}$/.test(seg)) {
          frag.appendChild(doc.createTextNode(seg));
        } else {
          const newR = doc.createElementNS(W_NS, "w:r");
          if (rPr) newR.appendChild(rPr.cloneNode(true));
          const newT = doc.createElementNS(W_NS, "w:t");
          newT.setAttribute("xml:space", "preserve");
          newT.textContent = seg;
          newR.appendChild(newT);
          frag.appendChild(newR);
        }
      }
      r.parentNode.replaceChild(frag, r);
    }

    return new XMLSerializer().serializeToString(doc);
  }

  // =========================================================================
  // 3. ESCAPE KONTEKS, PEMULIH TAG XML & SANITASI OPENXML
  // =========================================================================
  function escapeContext(val) {
    if (val === null || val === undefined) return "";
    if (typeof val === "string") {
      if (val.startsWith("§§RAW:")) return val.slice(6);
      return escapeXml(val).replace(/\r?\n/g, "§§BR§§");
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

  function repairMismatchedXmlTags(xmlString) {
    const tokenRegex = /<(\/?)([a-zA-Z0-9_:-]+)([^>]*)>/g;
    const stack = [];
    let out = "";
    let lastIndex = 0;
    let m;

    while ((m = tokenRegex.exec(xmlString)) !== null) {
      const [full, slash, tagName, attrs] = m;
      out += xmlString.slice(lastIndex, m.index);
      lastIndex = m.index + full.length;

      if (tagName.startsWith("?") || tagName.startsWith("!") || attrs.endsWith("/")) {
        out += full;
        continue;
      }

      if (!slash) {
        stack.push(tagName);
        out += full;
      } else {
        const idx = stack.lastIndexOf(tagName);
        if (idx === -1) continue;
        while (stack.length - 1 > idx) {
          out += `</${stack.pop()}>`;
        }
        stack.pop();
        out += full;
      }
    }
    out += xmlString.slice(lastIndex);
    while (stack.length > 0) {
      out += `</${stack.pop()}>`;
    }
    return out;
  }

  async function finalizeWordXml(xmlString, xmlPath, zip, imageJobs, allocRelId, docPrStart, addedExts, rekapXml) {
    const repaired = repairMismatchedXmlTags(xmlString);
    const parser = new DOMParser();
    const doc = parser.parseFromString(repaired, "application/xml");
    const parseErr = doc.querySelector("parsererror");
    if (parseErr) {
      const msg = (parseErr.textContent || "").replace(/\s+/g, " ").trim().slice(0, 260);
      throw new Error(`Struktur XML tidak valid pada ${xmlPath}: ${msg}`);
    }

    // 1. Ganti placeholder §§TABEL_REKAP§§ dengan elemen <w:tbl>
    const walker = doc.createTreeWalker(doc.documentElement, NodeFilter.SHOW_TEXT);
    const textNodes = [];
    let curNode;
    while ((curNode = walker.nextNode())) {
      if (curNode.nodeValue && curNode.nodeValue.includes("§§TABEL_REKAP§§")) {
        textNodes.push(curNode);
      }
    }
    for (const tn of textNodes) {
      const tblDoc = parser.parseFromString(rekapXml, "application/xml");
      const importedTbl = doc.importNode(tblDoc.documentElement, true);
      tn.parentNode.replaceChild(importedTbl, tn);
    }

    // 2. Konversi token §§BR§§ menjadi <w:br/>
    const allTForBr = Array.from(doc.getElementsByTagNameNS(W_NS, "t"));
    for (const t of allTForBr) {
      const txt = t.textContent || "";
      if (!txt.includes("§§BR§§")) continue;
      const lines = txt.split("§§BR§§");
      const parentRun = t.parentNode;
      t.textContent = lines[0];
      t.setAttribute("xml:space", "preserve");
      let refNode = t.nextSibling;
      for (let i = 1; i < lines.length; i++) {
        const br = doc.createElementNS(W_NS, "w:br");
        const nextT = doc.createElementNS(W_NS, "w:t");
        nextT.setAttribute("xml:space", "preserve");
        nextT.textContent = lines[i];
        parentRun.insertBefore(br, refNode);
        parentRun.insertBefore(nextT, refNode);
      }
    }

    // 3. Kunci baris tabel yang berisi gambar (§§IMG:) agar tidak terbelah antar halaman
    const trList = Array.from(doc.getElementsByTagNameNS(W_NS, "tr"));
    for (const tr of trList) {
      if (tr.getElementsByTagNameNS(W_NS, "tc").length === 0) {
        tr.parentNode.removeChild(tr);
        continue;
      }
      if ((tr.textContent || "").includes("§§IMG:")) {
        let trPr = tr.getElementsByTagNameNS(W_NS, "trPr")[0];
        if (!trPr) {
          trPr = doc.createElementNS(W_NS, "w:trPr");
          tr.insertBefore(trPr, tr.firstChild);
        }
        if (trPr.getElementsByTagNameNS(W_NS, "cantSplit").length === 0) {
          trPr.appendChild(doc.createElementNS(W_NS, "w:cantSplit"));
        }
      }
    }

    // 4. Pastikan setiap <w:tc> memenuhi aturan skema OpenXML Word
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

    // 5. Hapus seluruh text node liar di luar <w:t> (wajib bersih untuk skema Word)
    const cleanWalker = doc.createTreeWalker(doc.documentElement, NodeFilter.SHOW_TEXT);
    const orphanTexts = [];
    let n;
    while ((n = cleanWalker.nextNode())) {
      const pLocal = n.parentNode && n.parentNode.localName;
      if (pLocal !== "t" && pLocal !== "instrText" && pLocal !== "delText") {
        orphanTexts.push(n);
      }
    }
    for (const ot of orphanTexts) {
      ot.parentNode.removeChild(ot);
    }

    // Serialisasi DOM ke string XML dulu, BARU sisipkan <w:drawing> murni agar namespace tidak diubah XMLSerializer
    let finalXml = new XMLSerializer().serializeToString(doc);
    let docPrCounter = docPrStart;

    if (xmlPath === "word/document.xml" && imageJobs && imageJobs.size > 0) {
      for (const [token, job] of imageJobs.entries()) {
        if (!finalXml.includes(token)) continue;
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
        addedExts.add(norm.ext);

        const relTag = `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${mediaFile}"/>`;
        imageJobs._newRels.push(relTag);

        const drawingXml = buildDrawingXml(rId, docPrId, `img_${docPrId}.${norm.ext}`, cx, cy);
        // Ganti <w:t ...>...token...</w:t> di dalam <w:r> secara presisi
        finalXml = finalXml.replace(
          new RegExp(`(<w:t\\b[^>]*>)([^<]*)${token}([^<]*)(</w:t>)`, "g"),
          (_, tOpen, before, after, tClose) => {
            const bPart = before ? `${tOpen}${before}${tClose}` : "";
            const aPart = after ? `${tOpen}${after}${tClose}` : "";
            return `${bPart}${drawingXml}${aPart}`;
          }
        );
      }
    }

    return finalXml;
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
    const rekapXml =
      Array.isArray(context.rekap_bahan_baku) && context.rekap_bahan_baku.length > 0
        ? buildBahanBakuTableXml(context.rekap_bahan_baku)
        : `<w:p xmlns:w="${W_NS}"/>`;

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
    const jobs = imageJobs || new Map();
    jobs._newRels = [];

    const xmlTargets = Object.keys(zip.files).filter(
      (p) => p === "word/document.xml" || /^word\/(header|footer)\d+\.xml$/.test(p)
    );

    for (const xmlPath of xmlTargets) {
      const rawXml = await zip.file(xmlPath).async("string");

      const existingDocPrIds = [...rawXml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map((m) => parseInt(m[1], 10));
      const docPrStart = (existingDocPrIds.length ? Math.max(...existingDocPrIds) : 100) + 10;

      const prepared = preprocessDocxDom(rawXml);
      const renderCtx = Object.assign({}, safeCtx);
      ensureAllTemplateVarsDefined(prepared, renderCtx);

      let rendered;
      try {
        rendered = env.renderString(prepared, renderCtx);
      } catch (err) {
        throw new Error(`Gagal mengevaluasi tag template pada ${xmlPath}: ${err.message}`);
      }

      const finalized = await finalizeWordXml(
        rendered,
        xmlPath,
        zip,
        jobs,
        allocRelId,
        docPrStart,
        addedExts,
        rekapXml
      );
      zip.file(xmlPath, finalized);
    }

    if (jobs._newRels.length > 0) {
      relsXml = relsXml.replace("</Relationships>", `${jobs._newRels.join("")}</Relationships>`);
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
