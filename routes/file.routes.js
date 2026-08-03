const express = require("express");
const router = express.Router();
const multer = require("multer");
const upload = multer({ dest: "uploads/" });
const { ensureAuthenticated } = require("../middleware/auth.middleware");
const { getAppToken } = require("../services/token.service");
const { ensureFolder, uploadFile, createBlankDoc, modifyThenUpload, convertAndAppendPdfPipeline } = require("../services/drive.service");
const { grantPermission } = require("../services/permission.service");

router.post("/upload", ensureAuthenticated, upload.single("file"), async (req, res) => {
    try{
                const flag = req.query.flag || req.session.selectedFlag;
                const selDocId = req.query.flag || req.session.selDocId;
                const fileName = req.query.fileName || req.session.fileName;
				console.log("uploaded filename", fileName);
				console.log("req original filename", req.file.originalname);
                console.log("upload flag:", flag);
                console.log("Selected Document Id:", selDocId);
                const appToken = await getAppToken();
                const folderId = await ensureFolder(appToken, process.env.SOP_TEMPLATE_FOLDER, process.env.TARGET_USER);
                //const data = await uploadFile(appToken, folderId, req.file.path, req.file.originalname);
                const data = await modifyThenUpload(appToken, folderId, req.file.path, fileName+".docx");
                console.log("modifyThenUpload response---->", data);
                await grantPermission(data.id, req.session.user.email, "write");
                res.json(data);
    }catch(err) {
                console.error(err.response?.data || err.message);
                res.status(500).send("Upload failed");
    }
});

router.post("/uploadSignedPdf", ensureAuthenticated, upload.single("file"), async (req, res) => {
    try{
                const { pdfFileId, role } = req.body;
                if (!pdfFileId) return res.status(400).send("Missing fileId");
                const appToken = await getAppToken();
                const folderId = await ensureFolder(appToken, process.env.FINAL_SOP_FOLDER, process.env.TARGET_USER);
                //const data = await uploadFile(appToken, folderId, req.file.path, req.file.originalname);
                const sourceFileId = pdfFileId;
                const data = await convertAndAppendPdfPipeline(appToken, sourceFileId, folderId);
                await grantPermission(data.id, req.session.user.email, "write");
                res.json(data);
    }catch(err) {
                console.error(err.response?.data || err.message);
                res.status(500).send("Upload failed");
    }
});

router.post("/create-blank", ensureAuthenticated, async (req, res) => {
    try{
                const { fileId, role } = req.body;
                if (!fileId) return res.status(400).send("Missing fileId");
                const selectedRole = role || "write";
                const appToken = await getAppToken();
                //console.log("appToken----->", appToken);
                const sourceFileId = fileId;
                const targeFolderId = await ensureFolder(appToken, process.env.FINAL_SOP_FOLDER, process.env.TARGET_USER);
                console.log("sourceFileId----->", sourceFileId);
                console.log("targeFolderId----->", targeFolderId);
                const data = await createBlankDoc(appToken, sourceFileId, targeFolderId);
                console.log("data----->", data);
                await grantPermission(data.id, req.session.user.email, selectedRole);
                res.json(data);
    }catch (err) {
                console.error(err.response?.data || err.message);
                res.status(500).send("Create blank document failed");
    }
});

module.exports = router;
