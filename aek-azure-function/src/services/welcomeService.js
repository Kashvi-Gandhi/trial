const { getGraphClient } = require("./graphAuth");

/**
 * Searches SharePoint Document Library for a matching PDF file
 */
async function findSharePointDocument(campus, documentType) {
  const client = await getGraphClient();
  const siteId = process.env.SITE_ID;
  const driveId = process.env.DRIVE_ID;

  const campusField = process.env.CAMPUS_FIELD;
  const docTypeField = process.env.DOCUMENT_TYPE_FIELD;

  if (!siteId || !driveId) {
    throw new Error("Missing SITE_ID or DRIVE_ID in local.settings.json.");
  }

  // Fetch drive items directly
  const response = await client
    .api(`/sites/${siteId}/drives/${driveId}/root/children`)
    .get();

  if (!response.value || response.value.length === 0) {
    return null;
  }

  const items = response.value;

  // 1. Search by Filename matching (e.g. "London_Welcome_Pack.pdf" or "Rome_Contact_Info.pdf")
  const targetCampus = campus.toLowerCase();
  const targetDocType = documentType.toLowerCase().replace(/\s+/g, "");

  let matchedItem = items.find((item) => {
    const name = item.name.toLowerCase().replace(/[\_\-\s]/g, "");
    return name.includes(targetCampus) && name.includes(targetDocType);
  });

  // 2. Fallback: match by Campus name alone if document type isn't in file name
  if (!matchedItem) {
    matchedItem = items.find((item) => {
      const name = item.name.toLowerCase();
      return name.includes(targetCampus) && name.endsWith(".pdf");
    });
  }

  return matchedItem || null;
}

/**
 * Streams the PDF binary array buffer using the file's Graph item ID
 */
async function getDocumentStream(itemId) {
  const client = await getGraphClient();
  const siteId = process.env.SITE_ID;
  const driveId = process.env.DRIVE_ID;

  return await client
    .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}/content`)
    .get();
}

module.exports = {
  findSharePointDocument,
  getDocumentStream
};