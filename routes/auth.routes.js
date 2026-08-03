const express = require("express");
const router = express.Router();
const { cca } = require("../config/msalConfig");

router.get("/login", async (req, res) => {
        const flag = req.query.flag;
        const selDocId = req.query.curDocId;
        req.session.selectedFlag = flag;
        req.session.selDocId = selDocId;
        console.log("flag:", flag);
        const authUrl = await cca.getAuthCodeUrl({
                scopes: ["User.Read", "offline_access"],
                redirectUri: process.env.REDIRECT_URI,
                state: {
					"flag":flag,
					"selDocId":selDocId
				}
        });
        res.redirect(authUrl);
});

router.get("/auth/callback", async (req, res) => {
        console.log("/auth/callback");
		console.log("req.query.---->", req.query)
        const flag = req.query.state.flag;
        const selDocId = req.query.state.selDocId;
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
        res.redirect("/?flag="+flag+"selDocId="+selDocId);
});



module.exports = router;

