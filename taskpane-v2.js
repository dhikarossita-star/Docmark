/* global Office, Word */

const HEADER_CC_TITLES = ["DOCMARK_HEADER_V1", "DOCMARK_HEADER_V2", "DOCMARK_HEADER_V3"];
const WATERMARK_CC_TITLES = ["DOCMARK_WATERMARK_V1", "DOCMARK_WATERMARK_V2", "DOCMARK_WATERMARK_V3"];
const HEADER_CC_TITLE = "DOCMARK_HEADER_V3";
const WATERMARK_CC_TITLE = "DOCMARK_WATERMARK_V3";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("Add-in ini khusus Microsoft Word.", true);
    return;
  }

  document.getElementById("apply").addEventListener("click", () => applyTools({ header: true, watermark: true }));
  document.getElementById("watermarkOnly").addEventListener("click", () => applyTools({ header: false, watermark: true }));
  document.getElementById("headerOnly").addEventListener("click", () => applyTools({ header: true, watermark: false }));
  document.getElementById("removeWatermark").addEventListener("click", removeWatermarkEverywhere);

  setStatus("Siap — watermark diagonal V3 (image-based untuk iPad).");
});

function setStatus(message, isError = false) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = "status " + (isError ? "err" : "ok");
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/*
 * Word iPad currently renders classic VML text-watermark rotation inconsistently.
 * V3 therefore draws the diagonal text into a transparent PNG first, then inserts
 * that image as a floating DrawingML object behind the document text.
 */
function makeWatermarkPngBase64(text) {
  const canvas = document.createElement("canvas");
  canvas.width = 1600;
  canvas.height = 1600;

  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.save();
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate(-45 * Math.PI / 180);
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = "rgba(128,128,128,0.23)";

  // Fit arbitrary watermark text while keeping TERBATAS close to the reference.
  let fontPx = 270;
  do {
    ctx.font = `700 ${fontPx}px Arial, Helvetica, sans-serif`;
    if (ctx.measureText(text).width <= 1380) break;
    fontPx -= 10;
  } while (fontPx > 100);

  ctx.fillText(text, 0, 0);
  ctx.restore();

  return canvas.toDataURL("image/png").split(",")[1];
}

function buildImageWatermarkOoxml(base64Png) {
  const stamp = String(Date.now());
  const docPrId = Number(stamp.slice(-8)) || 736291;

  // 520 pt square, centered on page. 1 pt = 12700 EMU.
  const cx = 6604000;
  const cy = 6604000;

  return `
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels"
            pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
          Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>

  <pkg:part pkg:name="/word/_rels/document.xml.rels"
            pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdImg1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
          Target="media/docmark-watermark.png"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>

  <pkg:part pkg:name="/word/media/docmark-watermark.png"
            pkg:contentType="image/png"
            pkg:compression="store">
    <pkg:binaryData>${base64Png}</pkg:binaryData>
  </pkg:part>

  <pkg:part pkg:name="/word/document.xml"
            pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr><w:noProof/></w:rPr>
              <w:drawing>
                <wp:anchor distT="0" distB="0" distL="0" distR="0"
                           simplePos="0" relativeHeight="0"
                           behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
                  <wp:simplePos x="0" y="0"/>
                  <wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>
                  <wp:positionV relativeFrom="page"><wp:align>center</wp:align></wp:positionV>
                  <wp:extent cx="${cx}" cy="${cy}"/>
                  <wp:effectExtent l="0" t="0" r="0" b="0"/>
                  <wp:wrapNone/>
                  <wp:docPr id="${docPrId}" name="DocMark Watermark" descr="DocMark diagonal watermark"/>
                  <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks noChangeAspect="1"/>
                  </wp:cNvGraphicFramePr>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <pic:pic>
                        <pic:nvPicPr>
                          <pic:cNvPr id="0" name="docmark-watermark.png"/>
                          <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
                        </pic:nvPicPr>
                        <pic:blipFill>
                          <a:blip r:embed="rIdImg1" cstate="print"/>
                          <a:stretch><a:fillRect/></a:stretch>
                        </pic:blipFill>
                        <pic:spPr>
                          <a:xfrm>
                            <a:off x="0" y="0"/>
                            <a:ext cx="${cx}" cy="${cy}"/>
                          </a:xfrm>
                          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                          <a:noFill/>
                          <a:ln><a:noFill/></a:ln>
                        </pic:spPr>
                      </pic:pic>
                    </a:graphicData>
                  </a:graphic>
                </wp:anchor>
              </w:drawing>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

async function getHeaders(context) {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  const headers = [];
  for (const section of sections.items) {
    headers.push(section.getHeader(Word.HeaderFooterType.primary));
    headers.push(section.getHeader(Word.HeaderFooterType.firstPage));
    headers.push(section.getHeader(Word.HeaderFooterType.evenPages));
  }
  return headers;
}

async function deleteControlsByTitles(context, body, titles) {
  for (const title of titles) {
    const controls = body.contentControls.getByTitle(title);
    controls.load("items");
    await context.sync();
    for (const cc of controls.items) cc.delete(false);
    if (controls.items.length) await context.sync();
  }
}

async function setHeaderText(context, headerBody, text) {
  await deleteControlsByTitles(context, headerBody, HEADER_CC_TITLES);
  if (!text.trim()) return;

  const p = headerBody.insertParagraph(text.trim(), Word.InsertLocation.start);
  p.alignment = Word.Alignment.right;

  const range = p.getRange();
  range.font.name = "Arial";
  range.font.size = 9;
  range.font.italic = true;
  range.font.bold = false;
  range.font.color = "#B7B7B7";

  const cc = range.insertContentControl();
  cc.title = HEADER_CC_TITLE;
  cc.tag = HEADER_CC_TITLE;
  cc.appearance = Word.ContentControlAppearance.hidden;
  await context.sync();
}

async function setWatermark(context, headerBody, text) {
  await deleteControlsByTitles(context, headerBody, WATERMARK_CC_TITLES);
  if (!text.trim()) return;

  const pngBase64 = makeWatermarkPngBase64(text.trim());
  const ooxml = buildImageWatermarkOoxml(pngBase64);
  const range = headerBody.insertOoxml(ooxml, Word.InsertLocation.end);

  const cc = range.insertContentControl();
  cc.title = WATERMARK_CC_TITLE;
  cc.tag = WATERMARK_CC_TITLE;
  cc.appearance = Word.ContentControlAppearance.hidden;
  await context.sync();
}

async function applyTools(options) {
  const watermark = document.getElementById("watermark").value;
  const header = document.getElementById("header").value;

  if (options.watermark && !watermark.trim()) {
    setStatus("Isi teks watermark dulu.", true);
    return;
  }

  setStatus("Menerapkan watermark diagonal...");

  try {
    await Word.run(async (context) => {
      const headers = await getHeaders(context);
      for (const headerBody of headers) {
        if (options.header) await setHeaderText(context, headerBody, header);
        if (options.watermark) await setWatermark(context, headerBody, watermark);
      }
    });
    setStatus("Selesai. Watermark diagonal V3 sudah diterapkan.");
  } catch (error) {
    console.error(error);
    setStatus("Gagal membuat watermark V3: " + (error && error.message ? error.message : String(error)), true);
  }
}

async function removeWatermarkEverywhere() {
  setStatus("Menghapus watermark...");
  try {
    await Word.run(async (context) => {
      const headers = await getHeaders(context);
      for (const headerBody of headers) {
        await deleteControlsByTitles(context, headerBody, WATERMARK_CC_TITLES);
      }
    });
    setStatus("Watermark berhasil dihapus.");
  } catch (error) {
    console.error(error);
    setStatus("Gagal menghapus: " + (error && error.message ? error.message : String(error)), true);
  }
}
