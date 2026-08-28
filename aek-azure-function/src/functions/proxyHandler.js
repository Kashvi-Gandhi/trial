const { app } = require("@azure/functions");
const {
  findSharePointDocument,
  getDocumentStream,
} = require("../services/welcomeService");

app.http("proxyHandler", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "welcome-pack/{action}",
  handler: async (request, context) => {
    const action = request.params.action; // "info" or "content"
    const campus = request.query.get("campus");
    const documentType = request.query.get("documentType");

    context.log(
      `[PROXY] Request: action=${action}, campus=${campus}, documentType=${documentType}`,
    );

    try {
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
            // "Access-Control-Allow-Origin": "*",
            // "Access-Control-Allow-Methods": "GET, OPTIONS",
            // "Access-Control-Allow-Headers": "Content-Type, Accept",
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
        jsonBody: {
          available: false,
          error: error.name || "Error",
          message: error.message || "Internal Server Error",
        },
      };
    }
  },
});
