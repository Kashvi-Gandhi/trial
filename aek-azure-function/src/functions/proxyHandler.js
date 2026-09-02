const { app } = require("@azure/functions");
const {
  findSharePointDocument,
  getDocumentStream,
  getUserDetailsFromAIFS,
} = require("../services/welcomeService");

app.http("proxyHandler", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "welcome-pack/{action}",
  handler: async (request, context) => {
    const action = request.params.action; // "info", "content", or "user-details"
    const campus = request.query.get("campus");
    const documentType = request.query.get("documentType");
    const contactID = request.query.get("contactID") || "CGNYHA11PDO5";

    context.log(
      `[PROXY] Request: action=${action}, campus=${campus}, documentType=${documentType}, contactID=${contactID}`,
    );

    try {
      // 1. Endpoint for User Details (AIFS API)
      if (action === "user-details") {
        const userData = await getUserDetailsFromAIFS(contactID);
        return {
          status: 200,
          headers: {
            "Content-Type": "application/json",
            "Access-Control-Allow-Origin": "*"
          },
          jsonBody: userData
        };
      }

      // 2. SharePoint Document Actions
      const fileItem = await findSharePointDocument(campus, documentType);

      if (!fileItem) {
        return {
          status: 404,
          jsonBody: {
            available: false,
            message: `No ${documentType} document found for ${campus} campus.`,
          },
        };
      }

      // Metadata endpoint
      if (action === "info") {
        return {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*"
          },
          jsonBody: {
            available: true,
            campus: campus,
            documentType: documentType,
            fileName: fileItem.name,
            fileSize: fileItem.size,
          },
        };
      }

      // Binary PDF content stream endpoint
      if (action === "content") {
        const stream = await getDocumentStream(fileItem.id);
        return {
          status: 200,
          headers: {
            "Content-Type": "application/pdf",
            "Access-Control-Allow-Origin": "*"
          },
          body: stream,
        };
      }

      return {
        status: 400,
        jsonBody: { message: "Invalid action specified in route." },
      };
    } catch (error) {
      context.error("[PROXY SYSTEM ERROR]:", error);
      return {
        status: 500,
        headers: {
          "Access-Control-Allow-Origin": "*"
        },
        jsonBody: {
          available: false,
          error: error.name || "Error",
          message: error.message || "Internal Server Error",
        },
      };
    }
  },
});