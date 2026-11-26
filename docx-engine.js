/**
 * docx-engine.js -- Engine pembuat file .docx langsung di browser dari template
 * ber-tag Jinja/docxtpl ({{ var }}, {% for %}, {% if %}, {% tr for %}, {% tc for %}, {% p if %}).
 */
(function (global) {
  "use strict";

  const EMU_PER_MM = 36000;
  const EMU_PER_PX_96DPI = 9525;
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

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
      `</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="700"/><w:gridCol w:w="3200"/><w:gridCol w:w="3200"/><w:gridCol w:w="1400"/></w:tblGrid>`;
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
      .replace(/\u00A0/g, " ")
      .replace(/[\u201C\u201D\u201E\u201F]/g, '"')
      .replace(/[\u2018\u2019\u201A\u201B]/g, "'");
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

      let normalized = fullText.replace(/(\{\{[\s\S]*?\}\}|\{%[\s\S]*?%\})/g, (tag) =>
        normalizeJinjaExpression(tag)
      );

      // Label "Keterangan:" di depan tag keterangan dibuang: yang tercetak langsung isinya
      // (mis. "IMB dan Sertifikat", bukan "Keterangan: IMB dan Sertifikat").
      normalized = normalized.replace(
        /Keterangan\s*:\s*(\{\{\s*[\w\[\]\.]*keterangan[\w\[\]\.]*\s*\}\})/gi,
        "$1"
      );

      // Paragraf yang isinya HANYA tag keterangan: tampilkan hanya jika keterangannya terisi,
      // supaya tidak tersisa baris kosong ketika pengguna memilih tanpa keterangan.
      const capOnly = normalized.trim().match(/^\{\{\s*([\w\[\]\.]*keterangan[\w\[\]\.]*)\s*\}\}$/i);
      if (capOnly) {
        p.parentNode.insertBefore(doc.createTextNode(`{% if ${capOnly[1]} %}`), p);
        p.parentNode.insertBefore(doc.createTextNode(`{% endif %}`), p.nextSibling);
      }

      tNodes[0].textContent = normalized;
      tNodes[0].setAttribute("xml:space", "preserve");
      for (let i = 1; i < tNodes.length; i++) {
        tNodes[i].textContent = "";
      }
    }

    // Tahap 1a: kolom "No" kosong pada baris perulangan (sel pertama hanya berisi tag for/if,
    // baris di atasnya berjudul "No") -> isi nomor urut otomatis dengan {{ loop.index }}.
    for (const tr of Array.from(doc.getElementsByTagNameNS(W_NS, "tr"))) {
      const tcs = Array.from(tr.childNodes).filter((n) => n.localName === "tc" && n.namespaceURI === W_NS);
      if (tcs.length < 2) continue;
      const rowText = tcs.map((tc) =>
        Array.from(tc.getElementsByTagNameNS(W_NS, "t")).map((t) => t.textContent || "").join("")
      );
      if (!/\{%-?\s*(?:tr\s+)?for\b/.test(rowText.join(""))) continue;
      const prev = tr.previousSibling && tr.previousSibling.nodeType === 1 ? tr.previousSibling : null;
      const prevFirstTc = prev && prev.localName === "tr"
        ? Array.from(prev.childNodes).find((n) => n.localName === "tc" && n.namespaceURI === W_NS)
        : null;
      const prevFirst = prevFirstTc
        ? Array.from(prevFirstTc.getElementsByTagNameNS(W_NS, "t")).map((t) => t.textContent || "").join("").trim()
        : "";
      if (!/^no\.?$/i.test(prevFirst)) continue;
      if (rowText[0].replace(/\{%[\s\S]*?%\}/g, "").trim() !== "") continue;
      const ts = Array.from(tcs[0].getElementsByTagNameNS(W_NS, "t"));
      const last = ts[ts.length - 1];
      if (last) {
        last.textContent = (last.textContent || "") + "{{ loop.index }}";
        last.setAttribute("xml:space", "preserve");
      }
    }

    // Tahap 1b: {% for/if %} dan {% endfor/endif %} yang terpisah di sel (tc) berbeda dalam SATU baris
    // (mis. loop tabel rekap tanpa penanda "tr") diangkat ke level baris, supaya tabel tidak rusak.
    {
      const CTRL_RE = /\{%-?\s*(for|if|endfor|endif)\b[\s\S]*?-?%\}/g;
      for (const tr of Array.from(doc.getElementsByTagNameNS(W_NS, "tr"))) {
        const tcs = Array.from(tr.childNodes).filter((n) => n.localName === "tc" && n.namespaceURI === W_NS);
        const openers = [];
        const closers = [];
        for (const tc of tcs) {
          const stack = [];
          for (const t of Array.from(tc.getElementsByTagNameNS(W_NS, "t"))) {
            const txt = t.textContent || "";
            CTRL_RE.lastIndex = 0;
            let m;
            while ((m = CTRL_RE.exec(txt))) {
              const item = { t, start: m.index, end: m.index + m[0].length, text: m[0] };
              if (m[1] === "for" || m[1] === "if") stack.push(item);
              else if (stack.length) stack.pop();
              else closers.push(item);
            }
          }
          openers.push(...stack);
        }
        if (openers.length === 0 || openers.length !== closers.length) continue;

        const all = [...openers, ...closers].sort((a, b) => (a.t === b.t ? b.start - a.start : 0));
        for (const it of all) {
          const cur = it.t.textContent || "";
          it.t.textContent = cur.slice(0, it.start) + cur.slice(it.end);
        }
        for (const it of openers) {
          tr.parentNode.insertBefore(doc.createTextNode(it.text.replace(/\s+/g, " ")), tr);
        }
        const ref = tr.nextSibling;
        for (const it of closers) {
          tr.parentNode.insertBefore(doc.createTextNode(it.text.replace(/\s+/g, " ")), ref);
        }
      }
    }

    // Tahap 1b: {% for/if %} dibuka di satu sel dan ditutup di sel lain pada BARIS yang sama
    // (tanpa awalan tr). Jika dirender apa adanya, struktur sel rusak dan Word gagal membuka file.
    // Perlakukan sebagai tag baris: {%tr for ... %} ... {%tr endfor %}.
    const bal = (str) =>
      (str.match(/\{%-?\s*(?:for|if)\b/g) || []).length -
      (str.match(/\{%-?\s*end(?:for|if)\b/g) || []).length;
    for (const tr of Array.from(doc.getElementsByTagNameNS(W_NS, "tr"))) {
      const tcs = Array.from(tr.childNodes).filter((n) => n.localName === "tc" && n.namespaceURI === W_NS);
      if (tcs.length < 2) continue;
      const cellTexts = tcs.map((tc) =>
        Array.from(tc.getElementsByTagNameNS(W_NS, "t")).map((t) => t.textContent || "").join("")
      );
      const bals = cellTexts.map(bal);
      if (bals.some((x) => x !== 0) && bals.reduce((a, c) => a + c, 0) === 0) {
        tcs.forEach((tc, i) => {
          if (bals[i] === 0) return;
          for (const t of Array.from(tc.getElementsByTagNameNS(W_NS, "t"))) {
            const v = t.textContent || "";
            if (!v.includes("{%")) continue;
            t.textContent = v.replace(/\{%(-?)\s*(for|if|elif|else|endfor|endif)\b/g, "{%$1tr $2");
          }
        });
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
        txt = txt.replace(/\{%-?\s*tr\s+(endfor|endif|else|elif\b[\s\S]*?)\s*-?%\}/g, (_, expr) => {
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
        txt = txt.replace(/\{%-?\s*tc\s+(endfor|endif|else|elif\b[\s\S]*?)\s*-?%\}/g, (_, expr) => {
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
        txt = txt.replace(/\{%-?\s*p\s+(endfor|endif|else|elif\b[\s\S]*?)\s*-?%\}/g, (_, expr) => {
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
        txt = txt.replace(/\{%-?\s*r\s+(endfor|endif|else|elif\b[\s\S]*?)\s*-?%\}/g, (_, expr) => {
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
      if (fullText && /^(?:\{%(?:(?!%\})[\s\S])*%\}\s*)+$/.test(fullText)) {
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

    // XMLSerializer mengubah > < & " ' di dalam tag Jinja menjadi entity,
    // padahal Nunjucks butuh karakter aslinya (mis. {% if x > 0 %}).
    const serialized = new XMLSerializer().serializeToString(doc);
    return serialized.replace(/\{%[\s\S]*?%\}|\{\{[\s\S]*?\}\}/g, (tag) =>
      tag
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, "&")
    );
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
    if (Array.isArray(val)) {
      const out = val.map(escapeContext);
      if (Object.prototype.hasOwnProperty.call(val, "toString")) {
        const str = escapeContext(val.toString());
        Object.defineProperty(out, "toString", { value: () => str, enumerable: false });
      }
      return out;
    }
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

    // Gambar diproses di document.xml DAN header/footer (relasi ditulis per-bagian oleh generateDocx)
    if (imageJobs && imageJobs.size > 0) {
      const partRelIds = new Map(); // token -> rId untuk bagian XML ini
      for (const [token, job] of imageJobs.entries()) {
        if (!finalXml.includes(token)) continue;

        // Normalisasi + tulis file media sekali saja per token (dipakai ulang antar bagian)
        if (!job._media) {
          const norm = await normalizeImageForWord(job.blob);
          const natural = await getImageNaturalSize(job.blob).catch(() => ({ width: 400, height: 300 }));

          let cx, cy;
          const maxHeightMm =
            job.maxHeightMm || (job.widthMm && job.widthMm <= 80 ? 105 : 210);
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

          const mediaIndex = imageJobs._nextMedia++;
          const mediaFile = `media/img_${mediaIndex}.${norm.ext}`;
          zip.file(`word/${mediaFile}`, norm.buffer, { createFolders: false });
          addedExts.add(norm.ext);
          job._media = { mediaFile, ext: norm.ext, cx, cy };
        }

        const media = job._media;
        let rId = partRelIds.get(token);
        if (!rId) {
          rId = allocRelId();
          partRelIds.set(token, rId);
          imageJobs._newRels.push(
            `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${media.mediaFile}"/>`
          );
        }

        // Setiap kemunculan token mendapat docPr id BARU (id ganda membuat Word protes)
        finalXml = finalXml.replace(
          new RegExp(`(<w:t\\b[^>]*>)([^<]*)${token}([^<]*)(</w:t>)`, "g"),
          (_, tOpen, before, after, tClose) => {
            const docPrId = imageJobs._nextDocPr++;
            const drawingXml = buildDrawingXml(
              rId, docPrId, `img_${docPrId}.${media.ext}`, media.cx, media.cy
            );
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
    let contentTypesXml = await zip.file("[Content_Types].xml").async("string");

    const xmlTargets = Object.keys(zip.files).filter(
      (p) => p === "word/document.xml" || /^word\/(header|footer)\d+\.xml$/.test(p)
    );
    const relsPathFor = (xmlPath) =>
      xmlPath === "word/document.xml"
        ? relsPath
        : xmlPath.replace(/^word\//, "word/_rels/") + ".rels";

    // Kumpulkan semua rId yang sudah dipakai (document + header/footer) & id docPr tertinggi
    const usedRelIds = new Set();
    let maxDocPr = 100;
    const rawByPath = {};
    for (const xmlPath of xmlTargets) {
      rawByPath[xmlPath] = await zip.file(xmlPath).async("string");
      for (const m of rawByPath[xmlPath].matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)) {
        maxDocPr = Math.max(maxDocPr, parseInt(m[1], 10));
      }
      const rf = zip.file(relsPathFor(xmlPath));
      if (rf) {
        const txt = await rf.async("string");
        for (const m of txt.matchAll(/Id="([^"]+)"/g)) usedRelIds.add(m[1]);
      }
    }
    let nextRelNum = 100;
    const allocRelId = () => {
      while (usedRelIds.has(`rId${nextRelNum}`)) nextRelNum++;
      const id = `rId${nextRelNum++}`;
      usedRelIds.add(id);
      return id;
    };

    const addedExts = new Set();
    const jobs = imageJobs || new Map();
    jobs._nextMedia = 1;
    jobs._nextDocPr = maxDocPr + 10;
    const newRelsByPart = {};

    for (const xmlPath of xmlTargets) {
      const rawXml = rawByPath[xmlPath];

      const prepared = preprocessDocxDom(rawXml);
      const renderCtx = Object.assign({}, safeCtx);
      ensureAllTemplateVarsDefined(prepared, renderCtx);

      let rendered;
      try {
        rendered = env.renderString(prepared, renderCtx);
      } catch (err) {
        throw new Error(`Gagal mengevaluasi tag template pada ${xmlPath}: ${err.message}`);
      }

      jobs._newRels = [];
      const finalized = await finalizeWordXml(
        rendered,
        xmlPath,
        zip,
        jobs,
        allocRelId,
        0,
        addedExts,
        rekapXml
      );
      // XMLSerializer di browser bisa sudah menyertakan deklarasi sendiri -> pastikan hanya ada satu
      const finalizedBody = finalized.replace(/^\s*<\?xml[^>]*\?>\s*/, "");
      zip.file(xmlPath, XML_DECL + finalizedBody, { createFolders: false });
      if (jobs._newRels.length > 0) newRelsByPart[xmlPath] = jobs._newRels.slice();
    }

    // Tulis relasi gambar ke file .rels milik masing-masing bagian (buat baru jika belum ada)
    for (const [xmlPath, rels] of Object.entries(newRelsByPart)) {
      const rp = relsPathFor(xmlPath);
      const existing = zip.file(rp);
      let rx = existing
        ? await existing.async("string")
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
      rx = rx.replace("</Relationships>", `${rels.join("")}</Relationships>`);
      zip.file(rp, rx, { createFolders: false });
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

    zip.file("[Content_Types].xml", contentTypesXml, { createFolders: false });

    // Bangun ulang ZIP TANPA entri folder (word/, word/media/, ...). Word menolak membuka .docx
    // yang berisi entri folder ("Word experienced an error trying to open the file").
    const outZip = new JSZip();
    const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
    names.sort((a, b) => (a === "[Content_Types].xml" ? -1 : b === "[Content_Types].xml" ? 1 : 0));
    for (const name of names) {
      outZip.file(name, await zip.file(name).async("uint8array"), { createFolders: false });
    }

    return await outZip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
    });
  }

  global.DocxEngine = {
    generateDocx,
  };
})(window);
