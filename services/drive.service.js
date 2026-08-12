const axios = require("axios");
const fs = require("fs");
const PizZip = require("pizzip");
const { DOMParser, XMLSerializer } = require("xmldom");
const path = require("path");
const { PDFDocument, StandardFonts, rgb, degrees } = require("pdf-lib");

// DOCX Modification Helpers
function createCell(text, bold = false) {
  return `
        <w:tc>
          <w:tcPr>
                <w:tcW w:w="5000" w:type="dxa"/>
          </w:tcPr>
          <w:p>
                <w:r>
                  ${bold ? "<w:rPr><w:b/></w:rPr>" : ""}
                  <w:t xml:space="preserve">${text}</w:t>
                </w:r>
          </w:p>
        </w:tc>`;
}

// function createRow(values, options = {}) {
  // if (options.header) {
    // return `
      // <w:tr>
        // <w:tc>
          // <w:tcPr>
            // <w:tcW w:w="10000" w:type="dxa"/>
            // <w:gridSpan w:val="2"/>
            // <w:shd w:fill="D9D9D9"/>
          // </w:tcPr>
          // <w:p>
            // <w:r>
              // <w:rPr><w:b/></w:rPr>
              // <w:t>${values[0]}</w:t>
            // </w:r>
          // </w:p>
        // </w:tc>
      // </w:tr>`;
  // }

  // return `
    // <w:tr>
      // ${createCell(values[0])}
      // ${createCell(values[1])}
    // </w:tr>`;
// }

function createRow(label, value) {
  return `
    <w:tr>
      ${createCell(label, true)}
      ${createCell(value, false)}
    </w:tr>`;
}

// Modify DOCX: add first page + table
function modifyDocx(inputPath, outputPath) {
  const content = fs.readFileSync(inputPath);
  const zip = new PizZip(content);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  let documentXml = zip.file("word/document.xml").asText();
  let documentDom = parser.parseFromString(documentXml, "text/xml");
  const body = documentDom.getElementsByTagName("w:body")[0];

  const firstPageXml = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:br w:type="page"/></w:r>
        </w:p>

        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
                <w:tblLayout w:type="fixed"/>
                <w:tblW w:w="10000" w:type="dxa"/>
                <w:tblBorders>
                  <w:top w:val="single" w:sz="4"/>
                  <w:left w:val="single" w:sz="4"/>
                  <w:bottom w:val="single" w:sz="4"/>
                  <w:right w:val="single" w:sz="4"/>
                  <w:insideH w:val="single" w:sz="4"/>
                  <w:insideV w:val="single" w:sz="4"/>
                </w:tblBorders>
          </w:tblPr>

          <w:tblGrid>
                <w:gridCol w:w="5000"/>
                <w:gridCol w:w="5000"/>
          </w:tblGrid>

          ${createRow("Effective Date:", "22/DEC/2025")}
          ${createRow("Next Review Due Date:", "21/DEC/2027")}
        </w:tbl>
        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
                <w:tblLayout w:type="fixed"/>
                <w:tblW w:w="10000" w:type="dxa"/>
                <w:tblBorders>
                  <w:top w:val="single" w:sz="4"/>
                  <w:left w:val="single" w:sz="4"/>
                  <w:bottom w:val="single" w:sz="4"/>
                  <w:right w:val="single" w:sz="4"/>
                  <w:insideH w:val="single" w:sz="4"/>
                  <w:insideV w:val="single" w:sz="4"/>
                </w:tblBorders>
          </w:tblPr>
          <w:tblGrid>
                <w:gridCol w:w="2500"/>
                <w:gridCol w:w="2500"/>
                <w:gridCol w:w="2500"/>
                <w:gridCol w:w="2500"/>
          </w:tblGrid>

          <!-- Header row (merged) -->
          <w:tr>
                <w:tc>
                  <w:tcPr>
                        <w:tcW w:w="10000" w:type="dxa"/>
                        <w:gridSpan w:val="4"/>
                  </w:tcPr>
                  <w:p>
                        <w:r>
                          <w:rPr><w:b/></w:rPr>
                          <w:t>Author, Reviewer and Approver of SOP</w:t>
                        </w:r>
                  </w:p>
                </w:tc>
          </w:tr>

          <!-- Author row 1 -->
          <w:tr>
                ${createCell("Author:", true)}
                ${createCell("Dr. Nafisa Kathiwala")}
                ${createCell("Signature:", true)}
                ${createCell("")}
          </w:tr>
          <!-- Author row 2 -->
          <w:tr>
                ${createCell("Title:", true)}
                ${createCell("Director - Quality Assurance")}
                ${createCell("Date:", true)}
                ${createCell("")}
          </w:tr>

          <!-- Reviewer row 1 -->
          <w:tr>
                ${createCell("Reviewer:", true)}
                ${createCell("Arun Janardhanan")}
                ${createCell("Signature:", true)}
                ${createCell("")}
          </w:tr>
          <!-- Reviewer row 2 -->
          <w:tr>
                ${createCell("Title:", true)}
                ${createCell("Sr. Manager - Project Delivery")}
                ${createCell("Date:", true)}
                ${createCell("")}
          </w:tr>

          <!-- Approver row 1 -->
          <w:tr>
                ${createCell("Approver:", true)}
                ${createCell("Hiren Thakkar")}
                ${createCell("Signature:", true)}
                ${createCell("")}
          </w:tr>
          <!-- Approver row 2 -->
          <w:tr>
                ${createCell("Title:", true)}
                ${createCell("Managing Director")}
                ${createCell("Date:", true)}
                ${createCell("")}
          </w:tr>
        </w:tbl>

        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
                <w:spacing w:before="360" w:after="360"/> <!-- gap between tables -->
          </w:pPr>
        </w:p>

        <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:tblPr>
                <w:tblLayout w:type="fixed"/>
                <w:tblW w:w="10000" w:type="dxa"/>
                <w:tblBorders>
                  <w:top w:val="single" w:sz="4"/>
                  <w:left w:val="single" w:sz="4"/>
                  <w:bottom w:val="single" w:sz="4"/>
                  <w:right w:val="single" w:sz="4"/>
                  <w:insideH w:val="single" w:sz="4"/>
                  <w:insideV w:val="single" w:sz="4"/>
                </w:tblBorders>
          </w:tblPr>
          <w:tblGrid>
                <w:gridCol w:w="2500"/>
                <w:gridCol w:w="2500"/>
                <w:gridCol w:w="2500"/>
                <w:gridCol w:w="2500"/>
          </w:tblGrid>

          <!-- Effective and review dates -->
          <w:tr>
                ${createCell("Effective Date:", true)}
                ${createCell("22/DEC/2025")}
                ${createCell("Next Review Due Date:", true)}
                ${createCell("21/DEC/2027")}
          </w:tr>

          <!-- Superseded SOP row (label merged across 2 rows, right side merged) -->
          <w:tr>
                <w:tc>
                  <w:tcPr>
                        <w:tcW w:w="2500" w:type="dxa"/>
                        <w:vMerge w:val="restart"/>
                  </w:tcPr>
                  <w:p>
                        <w:r>
                          <w:rPr><w:b/></w:rPr>
                          <w:t>Superseded SOP with SOP's effective Date</w:t>
                        </w:r>
                  </w:p>
                </w:tc>
                <w:tc>
                  <w:tcPr>
                        <w:tcW w:w="7500" w:type="dxa"/>
                        <w:gridSpan w:val="3"/>
                  </w:tcPr>
                  <w:p>
                        <w:r>
                          <w:t>D-011 Version No: 4.0 with effective date 22 NOV 2023</w:t>
                        </w:r>
                  </w:p>
                </w:tc>
          </w:tr>
          <w:tr>
                <w:tc>
                  <w:tcPr>
                        <w:vMerge/>
                  </w:tcPr>
                  <w:p/>
                </w:tc>
                <w:tc>
                  <w:tcPr>
                        <w:gridSpan w:val="3"/>
                  </w:tcPr>
                  <w:p/>
                </w:tc>
          </w:tr>
        </w:tbl>
        `;

  const tempDom = parser.parseFromString(
    `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${firstPageXml}</root>`,
    "text/xml"
  );

  while (tempDom.documentElement.firstChild) {
    body.insertBefore(
                tempDom.documentElement.firstChild,
                body.firstChild
        );
  }

  // =========================
  // ✅ ADD LOGO IMAGE TO DOCX
  // =========================
  const logoPath = path.join(__dirname, "logo.png");

  console.log("Logo path:", logoPath);
  console.log("Logo exists:", fs.existsSync(logoPath));

  if (!fs.existsSync(logoPath)) {
    throw new Error("Logo file not found at: " + logoPath);
  }

  const logoBuffer = fs.readFileSync(logoPath);

  // Add image inside docx
  zip.file("word/media/logo.png", logoBuffer);

        // =========================
  // Create Header & Footer Files
  // =========================
        zip.file(
        "word/header1.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">

        <!-- Top spacing -->
        <w:p>
        <w:pPr>
          <w:spacing w:before="400"/>
        </w:pPr>
        </w:p>

        <!-- Main Table -->
        <w:tbl>
        <w:tblPr>
          <w:tblW w:w="10000" w:type="dxa"/>
          <w:tblLayout w:type="fixed"/>
          <w:tblBorders>
                <w:top w:val="single" w:sz="6" w:color="000000"/>
                <w:left w:val="single" w:sz="6" w:color="000000"/>
                <w:bottom w:val="single" w:sz="6" w:color="000000"/>
                <w:right w:val="single" w:sz="6" w:color="000000"/>
                <w:insideH w:val="single" w:sz="4" w:color="000000"/>
                <w:insideV w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>

        <w:tblGrid>
          <w:gridCol w:w="3000"/>
          <w:gridCol w:w="7000"/>
        </w:tblGrid>

        <w:tr>

          <!-- LOGO -->
      <w:tc>
        <w:p>
          <w:r>
            <w:drawing>
              <wp:inline>
                <wp:extent cx="990000" cy="600000"/>
                <wp:docPr id="1" name="Logo"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic>
                      <pic:blipFill>
                        <a:blip r:embed="rIdImage1"/>
                        <a:stretch><a:fillRect/></a:stretch>
                      </pic:blipFill>
                      <pic:spPr>
                        <a:prstGeom prst="rect"/>
                      </pic:spPr>
                    </pic:pic>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
          </w:r>
        </w:p>
      </w:tc>

          <!-- Right Content Cell -->
          <w:tc>
                <w:tcPr>
                  <w:tcW w:w="7000" w:type="dxa"/>
                </w:tcPr>

                <!-- Inner Table -->
                <w:tbl>
                  <w:tblPr>
                        <w:tblW w:w="7000" w:type="dxa"/>
                        <w:tblLayout w:type="fixed"/>
                        <w:tblBorders>
                          <w:top w:val="single" w:sz="4"/>
                          <w:left w:val="single" w:sz="4"/>
                          <w:bottom w:val="single" w:sz="4"/>
                          <w:right w:val="single" w:sz="4"/>
                          <w:insideH w:val="single" w:sz="4"/>
                          <w:insideV w:val="single" w:sz="4"/>
                        </w:tblBorders>
                  </w:tblPr>

                  <w:tblGrid>
                        <w:gridCol w:w="3500"/>
                        <w:gridCol w:w="3500"/>
                  </w:tblGrid>

                  <!-- SOP Title -->
                  <w:tr>
                        <w:tc>
                          <w:tcPr>
                                <w:gridSpan w:val="2"/>
                          </w:tcPr>
                          <w:p>
                                <w:r>
                                  <w:rPr><w:b/></w:rPr>
                                  <w:t>SOP Title: SOP for Project Management</w:t>
                                </w:r>
                          </w:p>
                        </w:tc>
                  </w:tr>

                  <!-- SOP Number / Version -->
                  <w:tr>
                        <w:tc>
                          <w:p>
                                <w:r><w:rPr><w:b/></w:rPr><w:t>SOP Number: </w:t></w:r>
                                <w:r><w:t>D-011</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc>
                          <w:p>
                                <w:r><w:rPr><w:b/></w:rPr><w:t>Version Number: </w:t></w:r>
                                <w:r><w:t>5.0</w:t></w:r>
                          </w:p>
                        </w:tc>
                  </w:tr>

                  <!-- Effective Date / Page -->
                  <w:tr>
                        <w:tc>
                          <w:p>
                                <w:r><w:rPr><w:b/></w:rPr><w:t>Effective Date: </w:t></w:r>
                                <w:r><w:t>22/DEC/2025</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc>
                          <w:p>
                                <w:r><w:rPr><w:b/></w:rPr><w:t>Page </w:t></w:r>

                                <!-- PAGE FIELD -->
                                <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                                <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
                                <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                                <w:r><w:t>1</w:t></w:r>
                                <w:r><w:fldChar w:fldCharType="end"/></w:r>

                                <w:r><w:t> of 7</w:t></w:r>
                          </w:p>
                        </w:tc>
                  </w:tr>

                </w:tbl>

          </w:tc>
        </w:tr>
        </w:tbl>

        </w:hdr>`
        );

        zip.file(
                "word/_rels/header1.xml.rels",
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1"
                        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                        Target="media/logo.png"/>
                </Relationships>`
        );

        zip.file(
                "word/footer1.xml",
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>Custom Footer - All Pages</w:t></w:r></w:p>
                </w:ftr>`
        );

  // =========================
  // Update Relationships SAFELY
  // =========================
  let relsXml = zip.file("word/_rels/document.xml.rels").asText();
  let relsDom = parser.parseFromString(relsXml, "text/xml");

  const relationships = relsDom.getElementsByTagName("Relationship");

  let maxId = 1;
  for (let i = 0; i < relationships.length; i++) {
    const id = relationships[i].getAttribute("Id");
    const num = parseInt(id.replace("rId", ""));
    if (!isNaN(num) && num > maxId) maxId = num;
  }

  const headerRelId = "rId" + (maxId + 1);
  const footerRelId = "rId" + (maxId + 2);

  const headerRel = relsDom.createElement("Relationship");
  headerRel.setAttribute("Id", headerRelId);
  headerRel.setAttribute(
    "Type",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
  );
  headerRel.setAttribute("Target", "header1.xml");

  const footerRel = relsDom.createElement("Relationship");
  footerRel.setAttribute("Id", footerRelId);
  footerRel.setAttribute(
    "Type",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  );
  footerRel.setAttribute("Target", "footer1.xml");

  relsDom.documentElement.appendChild(headerRel);
  relsDom.documentElement.appendChild(footerRel);

  zip.file(
    "word/_rels/document.xml.rels",
    serializer.serializeToString(relsDom)
  );

  // =========================
  // Attach header/footer to section
  // =========================
  let sectPr = body.getElementsByTagName("w:sectPr")[0];

  if (!sectPr) {
    sectPr = documentDom.createElement("w:sectPr");
    body.appendChild(sectPr);
  }

  let pgMar = sectPr.getElementsByTagName("w:pgMar")[0];
  if (!pgMar) {
    pgMar = documentDom.createElement("w:pgMar");
    sectPr.appendChild(pgMar);
  }

  // Set your desired margins here (top margin 3,000 twips = 2.08 inches approx)
  pgMar.setAttribute("w:top", "1440");
  // pgMar.setAttribute("w:right", "1440");
  // pgMar.setAttribute("w:bottom", "1440");
  // pgMar.setAttribute("w:left", "1440");
  // pgMar.setAttribute("w:header", "720");
  // pgMar.setAttribute("w:footer", "720");
  // pgMar.setAttribute("w:gutter", "0");

  const headerRef = documentDom.createElement("w:headerReference");
  headerRef.setAttribute("w:type", "default");
  headerRef.setAttribute("r:id", headerRelId);

  const footerRef = documentDom.createElement("w:footerReference");
  footerRef.setAttribute("w:type", "default");
  footerRef.setAttribute("r:id", footerRelId);

  sectPr.appendChild(headerRef);
  sectPr.appendChild(footerRef);

  zip.file("word/document.xml", serializer.serializeToString(documentDom));

  // =========================
  // Update Content Types
  // =========================
        let contentTypes = zip.file("[Content_Types].xml").asText();

  if (!contentTypes.includes("header1.xml")) {
    contentTypes = contentTypes.replace(
      "</Types>",
      `
      <Override PartName="/word/header1.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
      <Override PartName="/word/footer1.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
      </Types>`
    );

    zip.file("[Content_Types].xml", contentTypes);
  }

  // =========================
  // Save
  // =========================
  const newDoc = zip.generate({ type: "nodebuffer" });
  fs.writeFileSync(outputPath, newDoc);

  console.log("DOCX modified successfully:", outputPath);
}


// Auto-create folder if missing
async function ensureFolder(appToken, folderPath, targetUser) {
  const pathSegments = folderPath.split("/").filter(Boolean); // ["Documents", "Octalsoft-Sparc-Development-Docs", "Final-SOP"]
  let parentId = "root";

  for (const segment of pathSegments) {
    try {
      // Get all children of the current parent folder
      const res = await axios.get(
        `https://graph.microsoft.com/v1.0/users/${targetUser}/drive/items/${parentId}/children`,
        {
          headers: { Authorization: `Bearer ${appToken}` },
        }
      );

      // Check if folder already exists
      let folder = res.data.value.find(
        (item) => item.folder && item.name.toLowerCase() === segment.toLowerCase()
      );

      if (folder) {
        parentId = folder.id; // folder exists, use its id
      } else {
        // Folder does not exist, create it
        const createRes = await axios.post(
          `https://graph.microsoft.com/v1.0/users/${targetUser}/drive/items/${parentId}/children`,
          {
            name: segment,
            folder: {},
            "@microsoft.graph.conflictBehavior": "rename",
          },
          {
            headers: { Authorization: `Bearer ${appToken}` },
          }
        );

        parentId = createRes.data.id; // newly created folder id
      }
    } catch (err) {
      console.error(`Error processing folder segment "${segment}":`, err.response?.data || err.message);
      throw err;
    }
  }

  console.log("Final folder ID:", parentId);
  return parentId;
}

/* async function ensureFolder(appToken, folderPath, targetUser) {
        try{
                const res = await axios.get(`https://graph.microsoft.com/v1.0/users/${targetUser}/drive/root:/${folderPath}`,
                        { headers: {
                                        Authorization: `Bearer ${appToken}`
                                }
                        }
                );
                                console.log("folderPat----->", folderPath);
                                console.log("res.data----->", res.data);
                                console.log("FOLDER ID:", res.data.id);
                return res.data.id;
        }catch{
                // Create folder
                const createRes = await axios.post(`https://graph.microsoft.com/v1.0/users/${targetUser}/drive/root/children`,
                        {
                                name: folderPath.split("/").pop(),
                                folder: {},
                                "@microsoft.graph.conflictBehavior": "rename"
                        },
                        { headers: {
                                        Authorization: `Bearer ${appToken}`
                                }
                        }
                );
                                console.log("folderPat----->", folderPath);
                                console.log("res.data----->", res.data);
                return createRes.data.id;
        }
} */

// Upload file to folder
async function uploadFile(appToken, folderId, filePath, fileName) {
	
	let response= "";
	try {
		const buffer = fs.readFileSync(filePath);
		const url = `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${folderId}:/${fileName}:/content`;
		response = await axios.put(url, buffer, {
			headers: {
				Authorization: `Bearer ${appToken}`,
				"Content-Type": "application/octet-stream"
			}
		});
		//const fileId = response.data.id;
		
		// Delete file only after successful upload
		fs.unlinkSync(filePath);
		//return response.data;

	} catch (error) {
		console.error("File upload failed:");
		console.error(error.response?.data || error.message);
		try {
			if (fs.existsSync(filePath)) {
							fs.unlinkSync(filePath);
			}
		} catch (cleanupError) {
			console.error("File cleanup failed:", cleanupError.message);
		}
		throw error;
	}
	console.log("before permission change file data",response )
	try {
		const fileId = response.data.id;

		const linkUrl =
		  `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.TARGET_USER)}` +
		  `/drive/items/${encodeURIComponent(fileId)}/createLink`;

		console.log("Creating sharing link:", linkUrl);

		const linkResponse = await axios.post(
		  linkUrl,
		  {
			type: "view",
			scope: "organization"
		  },
		  {
			headers: {
			  Authorization: `Bearer ${appToken}`,
			  "Content-Type": "application/json"
			}
		  }
		);

		console.log("Link created:", linkResponse.data.link?.webUrl);

		return response.data;

	} catch (error) {
		console.error("CREATE LINK failed:");
		console.error("Status:", error.response?.status);
		console.error("URL:", error.config?.url);
		console.error(
		  "Response:",
		  JSON.stringify(error.response?.data, null, 2)
		);

		// Don't delete anything here—the upload already succeeded.
		throw error;
	}
}

//grant user permission for particular document on MS365
async function grantMixedPermissions(driveId, itemId, accessToken, permissionsConfig) {
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/invite`;

    const requestPromises = permissionsConfig.map(config => {
        const recipientObjects = config.emails.map(email => ({ email }));

        const payload = {
            recipients: recipientObjects,
            roles: [config.role], // ["write"] or ["read"]
            requireSignIn: true,
            sendInvitation: false
        };

        return fetch(url, {
            method: "POST",
            headers: {
                Authorization: `Bearer ${accessToken}`,
				
                "Content-Type": "application/json"
            },
            body: JSON.stringify(payload)
        }).then(async response => {
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(
                    `Failed for role ${config.role}: ${response.status} - ${
                        errorData.error?.message || response.statusText
                    }`
                );
            }
            return response.json();
        });
    });

    const results = await Promise.all(requestPromises);
    console.log("All permissions processed successfully:", results);
    return results;
}


// Create blank Word document
async function createBlankDoc(appToken, sourceFileId, targetFolderId, fileName) {
    //Get original file name
    const metadata = await getFileMetadata(appToken, sourceFileId);
	const permissionsToGrant = [
		{
			role: "read",
			emails: ["dummy.test@sunpharma.com"]
		}
	];
    // const originalName = metadata.name;
    const originalName = fileName+".docx";
        console.log("metadata.parentReference---->", metadata.parentReference)
        //console.log('createBlankDoc--->', appToken, targetFolderId, originalName);
        // const url = `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${folderId}:/${fileName}.docx:/content`;
        const url = `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${sourceFileId}/copy`;
        try {
			/* const response = await axios.put(url, "", {
					headers: {
							Authorization: `Bearer ${appToken}`,
							"Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
			}
			}); */
			const response= await axios.post(
				url,
				{
					parentReference: { id: targetFolderId },
					name: originalName // ✅ reuse original name
				},
				{
					headers: {
						Authorization: `Bearer ${appToken}`,
						"Content-Type": "application/json"
					}
				}
			);
			console.log("successful copied file");
			/* console.log("response data", response);
			return response.data; */
			// ✅ Get the monitor URL from the Location header
			const monitorUrl = response.headers['location'];
			if (!monitorUrl) throw new Error("No monitor URL returned from copy operation");
			console.log("monitorUrl---->", monitorUrl)
			// ✅ Poll until the copy operation completes
			await pollCopyOperation(monitorUrl);
			const newFileId = await findFileInFolder(appToken, targetFolderId, originalName);
			console.log("newFile----->", newFileId);
			const targetMetadata = await getFileMetadata(appToken, targetFolderId);
			const driveId= targetMetadata.parentReference.driveId;
			await grantMixedPermissions(driveId, newFileId, appToken, permissionsToGrant);//grant user permission
			return newFileId;
        } catch (error) {
			console.error("Upload failed:");
			console.error(error.response?.data || error.message);
			throw error;
        }
}

async function pollCopyOperation(monitorUrl, intervalMs = 1000, maxAttempts = 30) {
    for (let attempt = 0; attempt < maxAttempts; attempt++) {
        await new Promise(resolve => setTimeout(resolve, intervalMs));
        const statusResponse = await axios.get(monitorUrl);
        const { status, resourceId, errorCode } = statusResponse.data;
        console.log(`Poll attempt ${attempt + 1}: status = ${status}`);
        if (status === 'completed') {
            // ✅ resourceId is the new file's drive item ID
            return resourceId;
        }
        if (status === 'failed') {
            throw new Error(`Copy operation failed: ${errorCode}`);
        }
        // status is 'inProgress' or 'notStarted' — keep polling
    }
    throw new Error("Copy operation timed out after max polling attempts");
}

// ✅ FIND the newly created file in target folder
async function findFileInFolder(appToken, folderId, fileName) {
    const url = `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${folderId}/children`;
    const response = await axios.get(url, {
        headers: { Authorization: `Bearer ${appToken}` }
    });
	console.log("findFileInFolder response----->", response.data.value);
    // Find exact match (case-insensitive)
    const newFile = response.data.value.find(item => 
        item.name.toLowerCase() === fileName.toLowerCase() && 
        !item.folder // Ensure it's a file, not folder
    );

    if (!newFile) {
        throw new Error(`New file "${fileName}" not found in folder`);
    }
    return newFile; // ✅ This is the ACTUAL new file ID!
}

async function getFileMetadata(appToken, fileId) {
    const url = `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${fileId}`;

    const res = await axios.get(url, {
        headers: {
            Authorization: `Bearer ${appToken}`
        }
    });

    return res.data;
}

// Modify local file first, then upload
async function modifyThenUpload(appToken, folderId, localFilePath, uploadFileName){
        console.log(__dirname);
        const tempModifiedFile = path.join(__dirname, `temp_modified_${Date.now()}.docx`);

        // 1️⃣ Modify the local DOCX
        //modifyDocx(localFilePath, tempModifiedFile);

        // 2️⃣ Upload modified DOCX
        //await uploadFile(appToken, folderId, tempModifiedFile, uploadFileName);
        //const uploadFileResData= await uploadFile(appToken, folderId, tempModifiedFile, uploadFileName);
        const uploadFileResData= await uploadFile(appToken, folderId, localFilePath, uploadFileName);
        return uploadFileResData;
}

async function convertDocxToPdf(appToken, fileId, outputPath) {
		console.log(fileId, outputPath);
        try{
                const url = `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${fileId}/content?format=pdf`;
                const response = await axios.get(url, {
                        headers: {
                                Authorization: `Bearer ${appToken}`,
                        },
                        responseType: "arraybuffer",
                });
                fs.writeFileSync(outputPath, response.data);
                console.log("PDF created at:", outputPath);
                return outputPath;
        } catch(error){
                console.error("PDF conversion failed:");
                console.error(error.response?.data || error.message);
                throw error;
        }
}

/* async function appendCustomPagesToPdf(inputPdfPath, outputPdfPath, dynamicData) {
	try {
		const existingPdfBytes = fs.readFileSync(inputPdfPath);
		const pdfDoc = await PDFDocument.load(existingPdfBytes);
		const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
		const dynamicData= ["Approved by QA", "Reviewed on 2026-01-01", "Confidential Document", "Version 5.0 Released"];
		// 👉 Add new page
		const firstPage = pdfDoc.getPages()[0];
		const { width, height } = firstPage.getSize();
		console.log("width--------->", width);
		console.log("height--------->", height);
		const page = pdfDoc.addPage([width, height]);

		let y = height - 50;

		// Title
		page.drawText("Dynamic Data Section", {
			x: 50,
			y,
			size: 16,
			font,
			color: rgb(0, 0, 0),
		});

		y -= 30;

		// Dynamic content
		dynamicData.forEach((item, index) => {
			page.drawText(`${index + 1}. ${item}`, {
				x: 50,
				y,
				size: 12,
				font,
			});
			y -= 20;

			// Auto new page if overflow
			if (y < 50) {
				const newPage = pdfDoc.addPage();
				y = height - 50;
				newPage.drawText(`${index + 1}. ${item}`, {
						x: 50,
						y,
						size: 12,
						font,
				});
			}
		});

		const pdfBytes = await pdfDoc.save();
		fs.writeFileSync(outputPdfPath, pdfBytes);

		console.log("Custom pages appended to PDF");
		return outputPdfPath;

	} catch (error) {
		console.error("Appending PDF pages failed:", error.message);
		throw error;
	}
} */
async function appendCustomPagesToPdf(inputPdfPath, outputPdfPath) {
    try {
        const existingPdfBytes = fs.readFileSync(inputPdfPath);

        const pdfDoc = await PDFDocument.load(existingPdfBytes);

        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

        const firstPage = pdfDoc.getPages()[0];
        const { width, height } = firstPage.getSize();

        // ADD NEW PAGE
        const page = pdfDoc.addPage([width, height]);

        // -----------------------------
        // TABLE CONFIG
        // -----------------------------
        const startX = 40;
        let startY = height - 80;

        const tableWidth = width - 80;

        const colWidths = [
            120, // label column
            140, // name
            140, // designation
            140, // signature
        ];

        const rowHeights = [40];

        const headers = [
            "",
            "Name",
            "Designation,Department/\nFunction",
            "Signature with Date\n(DD-MMM-YYYY)"
        ];

        const rows = [
            ["Author:", "", "", ""],
            ["Reviewed By:", "", "", ""],
            ["Quality Assurance\nReviewed By:", "", "", ""],
            ["Approved By:", "", "", ""],
            ["QA Approved By:", "", "", ""]
        ];
		rows.forEach((row, rowIndex) => {
			rowHeights.push(70);
		})

        // -----------------------------
        // DRAW TABLE
        // -----------------------------
		
        let currentY = startY;
        // Draw Header Row
        let currentX = startX;
        for (let i = 0; i < colWidths.length; i++) {
            // CELL BORDER
            page.drawRectangle({
                x: currentX,
                y: currentY - rowHeights[0],
                width: colWidths[i],
                height: rowHeights[0],
                borderWidth: 1,
                color: rgb(1, 1, 1),
                borderColor: rgb(0, 0, 0),
            });
            // HEADER TEXT
            drawMultilineText(
                page,
                headers[i],
                currentX + 5,
                currentY - 20,
                boldFont,
                10,
                12
            );
            currentX += colWidths[i];
        }
        currentY -= rowHeights[0];

        // -----------------------------
        // BODY ROWS
        // -----------------------------
        rows.forEach((row, rowIndex) => {
            let x = startX;
            for (let col = 0; col < colWidths.length; col++) {
                page.drawRectangle({
                    x,
                    y: currentY - rowHeights[rowIndex + 1],
                    width: colWidths[col],
                    height: rowHeights[rowIndex + 1],
                    borderWidth: 1,
                    color: rgb(1, 1, 1),
                    borderColor: rgb(0, 0, 0),
                });
                drawMultilineText(page, row[col], x + 5, currentY - 20, col === 0 ? boldFont : font, 10, 12);
                x += colWidths[col];
            }
            currentY -= rowHeights[rowIndex + 1];
        });

        // SAVE PDF
        const pdfBytes = await pdfDoc.save();
        fs.writeFileSync(outputPdfPath, pdfBytes);
        console.log("Custom approval page added successfully");
        return outputPdfPath;
    } catch (error) {
        console.error("Appending PDF pages failed:", error.message);
        throw error;
    }
}

function drawMultilineText(page, text, x, y, font, size, lineHeight){
    const lines = text.split("\n");
    lines.forEach((line, index) => {
        page.drawText(line, {
            x,
            y: y - (index * lineHeight),
            size,
            font,
            color: rgb(0, 0, 0),
        });
    });
}

async function addWatermarkToPdf(inputPdfPath, outputPdfPath, watermarkText) {
	try {
		const existingPdfBytes = fs.readFileSync(inputPdfPath);
		const pdfDoc = await PDFDocument.load(existingPdfBytes);
		const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
		const pages = pdfDoc.getPages();
		pages.forEach((page) => {
			const { width, height } = page.getSize();
			page.drawText(watermarkText, {
				x: width / 4,
				y: height / 2,
				size: 40,
				font,
				color: rgb(0.75, 0.75, 0.75),
				rotate: degrees(30),
				opacity: 0.2
			});
		});
		const pdfBytes = await pdfDoc.save();
		fs.writeFileSync(outputPdfPath, pdfBytes);
		console.log("Watermark added successfully");
		return outputPdfPath;
	} catch (error) {
		console.error("Watermark failed:", error.message);
		throw error;
	}
}

async function convertAndAppendPdfPipeline(appToken, sourceFileId, folderId) {
	const fileMetadata = await getFileMetadata(appToken, sourceFileId);
	const orFileName = fileMetadata.name;
	const nameWithoutExt = orFileName.split('.').slice(0, -1).join('.');
	console.log(nameWithoutExt);

	try {
		const tempPdfPath = path.join(__dirname, `${nameWithoutExt}_${Date.now()}.pdf`);
		const finalPdfPath = path.join(__dirname, `${nameWithoutExt}.pdf`);
		console.log("sourceFileId---->", sourceFileId, orFileName)
		const newFileId = await findFileInFolder(appToken, folderId, orFileName);
		console.log("newFileId---->", newFileId, orFileName)
		// 1️⃣ Convert DOCX → PDF
		await convertDocxToPdf(appToken, sourceFileId, tempPdfPath);

		// 2️⃣ Append custom pages
		await appendCustomPagesToPdf(tempPdfPath, finalPdfPath);

		// 3️⃣ Add watermark (FINAL STEP)
		await addWatermarkToPdf(finalPdfPath, finalPdfPath, "OCTALSOFT");

		// 4️⃣ Upload final PDF
		const uploadRes = await uploadFile(appToken, folderId, finalPdfPath, `${nameWithoutExt}.pdf`);

		// Cleanup temp file
		if (fs.existsSync(tempPdfPath)) {
			fs.unlinkSync(tempPdfPath);
		}

		console.log("Pipeline completed successfully");

		return uploadRes;

	}catch(error) {
		console.error("Pipeline failed:", error.message);
		throw error;
	}
}

module.exports = { ensureFolder, uploadFile, createBlankDoc, modifyThenUpload, convertAndAppendPdfPipeline };
