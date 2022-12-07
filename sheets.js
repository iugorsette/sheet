require('dotenv').config()
const { GoogleSpreadsheet } = require("google-spreadsheet");
const credentials = require("./credenciais.json");
const idPlanilha = process.env.IDPLANILHA;

async function abrirPlanilha() {
  console.log(idPlanilha)
  const doc = new GoogleSpreadsheet(idPlanilha);
  await doc.useServiceAccountAuth({
    client_email: credentials.client_email,
    private_key: credentials.private_key,
  });
  await doc.loadInfo();
  return doc;
}

module.exports = { abrirPlanilha };
