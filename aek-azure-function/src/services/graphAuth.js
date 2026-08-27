const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

let graphClient = null;

async function getAccessTokenDirectly() {
  const tenantId = (process.env.TENANT_ID || "").trim();
  const clientId = (process.env.CLIENT_ID || "").trim();
  const clientSecret = (process.env.CLIENT_SECRET || "").trim();

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing TENANT_ID, CLIENT_ID, or CLIENT_SECRET in local.settings.json");
  }

  const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  
  const params = new URLSearchParams();
  params.append("client_id", clientId);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("client_secret", clientSecret);
  params.append("grant_type", "client_credentials");

  const response = await fetch(tokenEndpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: params.toString()
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(`Azure AD Token Error: ${data.error_description || data.error || response.statusText}`);
  }

  return data.access_token;
}

async function getGraphClient() {
  if (graphClient) return graphClient;

  graphClient = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        return await getAccessTokenDirectly();
      }
    }
  });

  return graphClient;
}

module.exports = { getGraphClient };