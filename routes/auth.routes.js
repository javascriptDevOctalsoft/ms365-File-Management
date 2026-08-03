const express = require("express");
const router = express.Router();
const { cca } = require("../config/msalConfig");

router.get("/login", async (req, res) => {
        const flag = req.query.flag;
        const selDocId = req.query.curDocId;
        const fileName = req.query.fileName+"_V"+req.query.verNo;
        const savedFlag = req.query.saved;
        req.session.selectedFlag = flag;
        req.session.selDocId = selDocId;
        req.session.fileName = fileName;
        req.session.savedFlag = savedFlag;
        console.log("flag:", flag);
        console.log("req.query:15", req.query);
		const state = JSON.stringify({ flag, selDocId, fileName, savedFlag});
		const 
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
        res.redirect("/?flag="+flag+"&selDocId="+selDocId+"&fileName="+fileName+"&saved="+savedFlag);
});



module.exports = router;

