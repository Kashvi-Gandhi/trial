/**
 * Fetches user details from AIFS API server-to-server (bypassing browser CORS)
 */
async function getUserDetailsFromAIFS(contactId) {
  const apiKey = process.env.X_CAMPUS_M_API_KEY || "XRnFCtOk1xZqXI6Vru7LpNrifGXbk1c63TtCwTb3lg4JrTvvshwvqb7zuVzgz537";
  const url = `https://secure.aifsabroad.com/apistg/v1/CampusM/Contact/GetAllApplications?contactID=${encodeURIComponent(contactId)}`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      "Accept": "application/json",
      "X-CAMPUS-M-API-KEY": apiKey
    }
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(data.message || data.error || `AIFS API error: ${response.statusText}`);
  }

  return data;
}

module.exports = { getUserDetailsFromAIFS };