const express = require("express");
const router = express.Router();
const { getAppToken } = require("../services/token.service");
const axios = require("axios");

router.post("/freeze/:fileId", async (req, res) => {
  try {
    const fileId = req.params.fileId;
    const token = await getAppToken();

    // 1️⃣ Downgrade all write users to read
    const perms = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${fileId}/permissions`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    for (const perm of perms.data.value) {
      if (perm.roles.includes("write")) {
        await axios.patch(
          `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${fileId}/permissions/${perm.id}`,
          { roles: ["read"] },
          { headers: { Authorization: `Bearer ${token}` } }
        );
      }
    }

    // 2️⃣ Create new editable copy
    await axios.post(
      `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${fileId}/copy`,
      { parentReference: { id: process.env.FOLDER_ID }, name: `NewVersion-${Date.now()}.docx` },
      { headers: { Authorization: `Bearer ${token}` } }
    );

    res.send("Freeze completed");
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).send("Freeze failed");
  }
});

module.exports = router;