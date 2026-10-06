/* global Office, Word */

const HEADER_CC_TITLE = "DOCMARK_HEADER_V1";
const WATERMARK_CC_TITLE = "DOCMARK_WATERMARK_V1";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("Add-in ini khusus Microsoft Word.", true);
    return;
  }

  document.getElementById("apply").addEventListener("click", () => applyTools({ header: true, watermark: true }));
  document.getElementById("watermarkOnly").addEventListener("click", () => applyTools({ header: false, watermark: true }));
  document.getElementById("headerOnly").addEventListener("click", () => applyTools({ header: true, watermark: false }));
  document.getElementById("removeWatermark").addEventListener("click", removeWatermarkEverywhere);

  setStatus("Siap.");
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

function buildWatermarkOoxml(text) {
  const safeText = xmlEscape(text);
  const uniqueId = "DOCMARK_WATERMARK_" + Date.now();

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

  <pkg:part pkg:name="/word/document.xml"
            pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:w10="urn:schemas-microsoft-com:office:word">
        <w:body>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
              <w:rPr><w:noProof/></w:rPr>
              <w:pict>
                <v:rect id="${uniqueId}"
                  style="position:absolute;margin-left:0;margin-top:0;width:420pt;height:95pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin;"
                  filled="f" stroked="f">
                  <v:textbox inset="0,0,0,0">
                    <w:txbxContent>
                      <w:p>
                        <w:pPr><w:jc w:val="center"/></w:pPr>
                        <w:r>
                          <w:rPr>
                            <w:b/>
                            <w:color w:val="B7B7B7"/>
                            <w:sz w:val="112"/>
                            <w:szCs w:val="112"/>
                            <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                          </w:rPr>
                          <w:t xml:space="preserve">${safeText}</w:t>
                        </w:r>
                      </w:p>
                    </w:txbxContent>
                  </v:textbox>
                  <w10:wrap type="none" anchorx="margin" anchory="margin"/>
                </v:rect>
              </w:pict>
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

async function deleteControlsByTitle(context, body, title) {
  const controls = body.contentControls.getByTitle(title);
  controls.load("items");
  await context.sync();

  for (const cc of controls.items) {
    cc.delete(false);
  }
  if (controls.items.length) await context.sync();
}

async function setHeaderText(context, headerBody, text) {
  await deleteControlsByTitle(context, headerBody, HEADER_CC_TITLE);
  if (!text.trim()) return;

  const p = headerBody.insertParagraph(text.trim(), Word.InsertLocation.start);
  p.alignment = Word.Alignment.left;

  const range = p.getRange();
  range.font.name = "Arial";
  range.font.size = 10;
  range.font.bold = false;

  const cc = range.insertContentControl();
  cc.title = HEADER_CC_TITLE;
  cc.tag = HEADER_CC_TITLE;
  cc.appearance = Word.ContentControlAppearance.hidden;
  await context.sync();
}

async function setWatermark(context, headerBody, text) {
  await deleteControlsByTitle(context, headerBody, WATERMARK_CC_TITLE);
  if (!text.trim()) return;

  const ooxml = buildWatermarkOoxml(text.trim());

  try {
    const range = headerBody.insertOoxml(ooxml, Word.InsertLocation.end);
    const cc = range.insertContentControl();
    cc.title = WATERMARK_CC_TITLE;
    cc.tag = WATERMARK_CC_TITLE;
    cc.appearance = Word.ContentControlAppearance.hidden;
    await context.sync();
  } catch (error) {
    const fallback = headerBody.insertParagraph(text.trim(), Word.InsertLocation.end);
    fallback.alignment = Word.Alignment.centered;
    const range = fallback.getRange();
    range.font.name = "Arial";
    range.font.size = 28;
    range.font.bold = true;
    range.font.color = "#B7B7B7";

    const cc = range.insertContentControl();
    cc.title = WATERMARK_CC_TITLE;
    cc.tag = WATERMARK_CC_TITLE;
    cc.appearance = Word.ContentControlAppearance.hidden;
    await context.sync();
  }
}

async function applyTools(options) {
  const watermark = document.getElementById("watermark").value;
  const header = document.getElementById("header").value;

  if (options.watermark && !watermark.trim()) {
    setStatus("Isi teks watermark dulu.", true);
    return;
  }

  setStatus("Menerapkan ke dokumen...");

  try {
    await Word.run(async (context) => {
      const headers = await getHeaders(context);
      for (const headerBody of headers) {
        if (options.header) await setHeaderText(context, headerBody, header);
        if (options.watermark) await setWatermark(context, headerBody, watermark);
      }
    });
    setStatus("Selesai. Watermark/header sudah diterapkan.");
  } catch (error) {
    console.error(error);
    setStatus("Gagal: " + (error && error.message ? error.message : String(error)), true);
  }
}

async function removeWatermarkEverywhere() {
  setStatus("Menghapus watermark...");
  try {
    await Word.run(async (context) => {
      const headers = await getHeaders(context);
      for (const headerBody of headers) {
        await deleteControlsByTitle(context, headerBody, WATERMARK_CC_TITLE);
      }
    });
    setStatus("Watermark berhasil dihapus.");
  } catch (error) {
    console.error(error);
    setStatus("Gagal menghapus: " + (error && error.message ? error.message : String(error)), true);
  }
}
