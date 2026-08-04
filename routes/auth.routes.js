const express = require("express");
const router = express.Router();
const { cca } = require("../config/msalConfig");

router.get("/login", async (req, res) => {
		let flag = null, selDocId = null, fileName= null, savedFlag= null, docType= null, sopType= null, labSOP = null;
        flag = req.query.flag;
        selDocId = req.query.curDocId;
		if(req.query.verNo != undefined && req.query.verNo !="null"){
			fileName = req.query.fileName+"_V"+req.query.verNo;
		}else{
			fileName = req.query.fileName;
		}
        
        savedFlag = req.query.saved;
        docType = req.query.docType;
        sopType = req.query.sopType;
        labSOP = req.query.labSOP;
        req.session.selectedFlag = flag;
        req.session.selDocId = selDocId;
        req.session.fileName = fileName;
        req.session.savedFlag = savedFlag;
        req.session.docType = docType;
        req.session.sopType = sopType;
        req.session.labSOP = labSOP;
        console.log("flag:", flag);
        console.log("req.query:15", req.query);
		const state = JSON.stringify({ flag, selDocId, fileName, savedFlag, docType, sopType, labSOP});
        const authUrl = await cca.getAuthCodeUrl({
                scopes: ["User.Read", "offline_access"],
                redirectUri: process.env.REDIRECT_URI,
                state: state
        });
        res.redirect(authUrl);
});

router.get("/auth/callback", async (req, res) => {
        console.log("/auth/callback");
        console.log("req.query.state 28--------->", req.query.state);
        const state = req.query.state ? JSON.parse(req.query.state) : {};
		const flag = state.flag;
		const selDocId = state.selDocId;
		const fileName = state.fileName;
		const verNo = state.verNo;
		const savedFlag = state.savedFlag;
		const docType = state.docType || null;
		const sopType = state.sopType || null;
		const labSOP = state.labSOP || null;
		console.log("selDocId---->", selDocId)
        const token = await cca.acquireTokenByCode({
                code: req.query.code,
                scopes: ["User.Read", "offline_access"],
                redirectUri: process.env.REDIRECT_URI,
        });

        req.session.user = {
                email: token.account.username,
                name: token.account.name,
                refreshToken: token.refreshToken
        };
        if (flag) {
                req.session.selectedFlag = flag;
        }
        res.redirect("/?flag="+flag+"&selDocId="+selDocId+"&fileName="+fileName+"&verNo="+verNo+"&saved="+savedFlag+"&docType="+docType+"&sopType="+sopType+"&labSOP="+labSOP);
});



module.exports = router;

