const { cca } = require("../config/msalConfig");

async function getAppToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result.accessToken;
}

module.exports = { getAppToken };