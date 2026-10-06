/* global Office, Word */

const HEADER_CC_TITLES = ["DOCMARK_HEADER_V1", "DOCMARK_HEADER_V2"];
const WATERMARK_CC_TITLES = ["DOCMARK_WATERMARK_V1", "DOCMARK_WATERMARK_V2"];
const HEADER_CC_TITLE = "DOCMARK_HEADER_V2";
const WATERMARK_CC_TITLE = "DOCMARK_WATERMARK_V2";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("Add-in ini khusus Microsoft Word.", true);
    return;
  }

  document.getElementById("apply").addEventListener("click", () => applyTools({ header: true, watermark: true }));
  document.getElementById("watermarkOnly").addEventListener("click", () => applyTools({ header: false, watermark: true }));
  document.getElementById("headerOnly").addEventListener("click", () => applyTools({ header: true, watermark: false }));
  document.getElementById("removeWatermark").addEventListener("click", removeWatermarkEverywhere);

  setStatus("Siap — versi watermark diagonal V2.");
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
 * True Word-style text watermark.
 * Uses the same VML WordArt pattern Word stores for diagonal watermarks:
 * PowerPlusWaterMarkObject + shapetype #_x0000_t136.
 * Sized to closely match the user's reference PDF.
 */
function buildWatermarkOoxml(text) {
  const safeText = xmlEscape(text);
  const unique = String(Date.now()).slice(-9);

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
            <w:pPr><w:pStyle w:val="Header"/></w:pPr>
            <w:r>
              <w:rPr><w:noProof/></w:rPr>
              <w:pict>
                <v:shapetype id="_x0000_t136"
                  coordsize="1600,21600"
                  o:spt="136"
                  adj="10800"
                  path="m@7,0l@8,0m@5,21600l@6,21600e">
                  <v:formulas>
                    <v:f eqn="sum #0 0 10800"/>
                    <v:f eqn="prod #0 2 1"/>
                    <v:f eqn="sum 21600 0 @1"/>
                    <v:f eqn="sum 0 0 @2"/>
                    <v:f eqn="sum 21600 0 @3"/>
                    <v:f eqn="if @0 @3 0"/>
                    <v:f eqn="if @0 21600 @1"/>
                    <v:f eqn="if @0 0 @2"/>
                    <v:f eqn="if @0 @4 21600"/>
                    <v:f eqn="mid @5 @6"/>
                    <v:f eqn="mid @8 @5"/>
                    <v:f eqn="mid @7 @8"/>
                    <v:f eqn="mid @6 @7"/>
                    <v:f eqn="sum @6 0 @5"/>
                  </v:formulas>
                  <v:path textpathok="t"
                          o:connecttype="custom"
                          o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"
                          o:connectangles="270,180,90,0"/>
                  <v:textpath on="t" fitshape="t"/>
                  <v:handles>
                    <v:h position="#0,bottomRight" xrange="6629,14971"/>
                  </v:handles>
                  <o:lock v:ext="edit" text="t" shapetype="t"/>
                </v:shapetype>

                <v:shape
                  id="PowerPlusWaterMarkObject${unique}"
                  o:spid="_x0000_s2049"
                  type="#_x0000_t136"
                  style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
                  o:allowincell="f"
                  fillcolor="silver"
                  stroked="f">
                  <v:fill opacity=".5"/>
                  <v:textpath style="font-family:&quot;Arial&quot;;font-size:1pt" string="${safeText}"/>
                  <w10:wrap anchorx="margin" anchory="margin"/>
                </v:shape>
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

  const ooxml = buildWatermarkOoxml(text.trim());
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

  setStatus("Menerapkan ke dokumen...");

  try {
    await Word.run(async (context) => {
      const headers = await getHeaders(context);
      for (const headerBody of headers) {
        if (options.header) await setHeaderText(context, headerBody, header);
        if (options.watermark) await setWatermark(context, headerBody, watermark);
      }
    });
    setStatus("Selesai. Watermark diagonal/header sudah diterapkan.");
  } catch (error) {
    console.error(error);
    setStatus("Gagal membuat watermark diagonal: " + (error && error.message ? error.message : String(error)), true);
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
