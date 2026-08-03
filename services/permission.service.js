const axios = require("axios");
const { getAppToken } = require("./token.service");

async function grantPermission(fileId, email, role) {
  const token = await getAppToken();
  await axios.post(
    `https://graph.microsoft.com/v1.0/users/${process.env.TARGET_USER}/drive/items/${fileId}/invite`,
    {
		recipients: [{ email }],
		roles: [role],
		requireSignIn: true,
		sendInvitation: false
    },
    { headers: { Authorization: `Bearer ${token}` } }
  );
}

module.exports = { grantPermission };