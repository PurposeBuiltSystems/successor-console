/*
 * Successor Console — minimal .docx writer.
 *
 * A .docx is a ZIP of OOXML parts. This module builds the XML parts as
 * strings (pure, Node-testable); the task pane zips them with JSZip.
 * Direct run formatting only (no styles.xml) keeps the file minimal and
 * valid. Research finding that forces .docx: .md is not usable agent
 * knowledge (absent from the supported-type table; not citable from
 * SharePoint libraries).
 *
 * Doc model: { title, blocks: [
 *   { type: 'h1'|'h2'|'p'|'li'|'quote', text },
 *   { type: 'link', text, url },        // paragraph that IS a hyperlink
 *   { type: 'kv', key, value },         // "Key: value" line
 * ] }
 */
(function (root) {
  "use strict";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;")
      // strip control chars that make OOXML invalid
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  }

  function run(text, o) {
    o = o || {};
    var props = "";
    if (o.bold) { props += "<w:b/>"; }
    if (o.italic) { props += "<w:i/>"; }
    if (o.size) { props += '<w:sz w:val="' + o.size + '"/>'; }
    if (o.color) { props += '<w:color w:val="' + o.color + '"/>'; }
    // xml:space="preserve" so leading/trailing spaces in excerpts survive
    return "<w:r>" + (props ? "<w:rPr>" + props + "</w:rPr>" : "") +
      '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>";
  }

  function para(inner, pPr) {
    return "<w:p>" + (pPr || "") + inner + "</w:p>";
  }

  /**
   * Render a doc model to OOXML parts.
   * Returns { parts: {path: xmlString}, relCount } — feed parts to JSZip.
   */
  function buildDocx(doc) {
    var body = [];
    var rels = [];
    var relId = function () { return "rId" + (rels.length + 1000); };

    body.push(para(run(doc.title || "", { bold: true, size: 36 })));

    (doc.blocks || []).forEach(function (b) {
      switch (b.type) {
        case "h1":
          body.push(para(run(b.text, { bold: true, size: 30 }),
            '<w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr>'));
          break;
        case "h2":
          body.push(para(run(b.text, { bold: true, size: 26 }),
            '<w:pPr><w:spacing w:before="160" w:after="60"/></w:pPr>'));
          break;
        case "li":
          body.push(para(run("• " + b.text)));
          break;
        case "quote":
          body.push(para(run(b.text, { italic: true, color: "444444" }),
            '<w:pPr><w:ind w:left="360"/></w:pPr>'));
          break;
        case "kv":
          body.push(para(run(b.key + ": ", { bold: true }) + run(b.value)));
          break;
        case "link": {
          var id = relId();
          rels.push({ id: id, url: b.url });
          body.push(para(
            '<w:hyperlink r:id="' + id + '">' +
            run(b.text, { color: "0563C1" }) + "</w:hyperlink>"
          ));
          break;
        }
        default:
          body.push(para(run(b.text || "")));
      }
    });

    var documentXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      "<w:body>" + body.join("") +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>' +
      "</w:body></w:document>";

    var documentRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels.map(function (r) {
        return '<Relationship Id="' + r.id + '" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
          'Target="' + esc(r.url) + '" TargetMode="External"/>';
      }).join("") +
      "</Relationships>";

    var contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      "</Types>";

    var pkgRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
      'Target="word/document.xml"/>' +
      "</Relationships>";

    return {
      parts: {
        "[Content_Types].xml": contentTypes,
        "_rels/.rels": pkgRels,
        "word/document.xml": documentXml,
        "word/_rels/document.xml.rels": documentRels,
      },
      relCount: rels.length,
    };
  }

  root.Docx = { buildDocx: buildDocx, _internals: { esc: esc } };
  if (typeof module !== "undefined" && module.exports) { module.exports = root.Docx; }
})(typeof self !== "undefined" ? self : this);
