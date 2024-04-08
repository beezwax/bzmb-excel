const createWorkbook = require("./createWorkbook.js");

const createWorkbookSchema = {
  body: {
    type: "object",
    required: [
      "sheets"
    ],
    properties: {
      sheets: {
        type: "array",
      },
      styles: {
        type: "object"
      }
    }
  }
};

async function bzmbExcel(fastify, options) {
  fastify.post(
    "/bzmb-excel-createWorkbook",
    { schema: createWorkbookSchema },
    async (req, res) => {
      try {
        const workbookBase64 = await createWorkbook(req.body);
        res
          .code(200)
          .send(workbookBase64);
      } catch (error) {
        res
          .code(500)
          .send(error);
      }
    }
  );
}

module.exports = { microbond: bzmbExcel };