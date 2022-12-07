const { abrirPlanilha } = require("./sheets");
const refs = require("./ref.json");

(async () => {
  try {
    await executar();
  } catch (e) {
    console.log(e);
  }
})();

async function executar() {
  const doc = await abrirPlanilha();
  await doc.loadInfo();
  const folhaRefPorMes = doc.sheetsByIndex[4];
  const folhaRefPorMesComTagCliente = doc.sheetsByIndex[5];
  const sheet = doc.sheetsByIndex[0];
  const linhas = await sheet.getRows();
  const meses = ["Julho", "Agosto", "Setembro", "Outubro", "Novembro"];
  await folhaRefPorMes.clearRows();
  await folhaRefPorMesComTagCliente.clearRows();
  await folhaRefPorMes.setHeaderRow(meses);
  await folhaRefPorMesComTagCliente.setHeaderRow(meses);
  await addLinhaNaFolhaRefTagCliente(linhas, meses, refs, folhaRefPorMesComTagCliente)
  await addLinhaNaFolhaRef(linhas, meses, refs, folhaRefPorMes);
}

async function addLinhaNaFolhaRefTagCliente(linhas, meses, refs, folhaRefPorMesComTagCliente) {
  let coluna = [];
  for (let ref of refs) {
    for (let mes of meses) {
      coluna.push(somaPorMesTagCliente(linhas, mes, ref));
    }
    coluna.push(somaTotalTagCliente(linhas, ref));
    coluna.push(ref);
    await folhaRefPorMesComTagCliente.addRow(coluna);
    coluna = [];
  }
}

async function addLinhaNaFolhaRef(linhas, meses, refs, folhaRefPorMes) {
  let coluna = [];
  for (let ref of refs) {
    for (let mes of meses) {
      coluna.push(somaPorMes(linhas, mes, ref));
    }
    coluna.push(somaTotal(linhas, ref));
    coluna.push(ref);
    await folhaRefPorMes.addRow(coluna);
    coluna = [];
  }
}

function somaTotalTagCliente(linhas, ref) {
  let soma = 0;
  linhas.forEach((linha) => {
    if (linha._rawData[2] === ref && linha._rawData[6].includes('CLIENTE')) {
      soma++;
    }
  });
  return soma;
}

function somaTotal(linhas, ref) {
  let soma = 0;
  linhas.forEach((linha) => {
    if (linha._rawData[2] === ref) {
      soma++;
    }
  });
  return soma;
}

function somaPorMesTagCliente(linhas,mes,ref){
  let soma = 0;
  linhas.forEach((linha) => {
    if (linha._rawData[2] === ref && linha._rawData[11] === mes && linha._rawData[6].includes('CLIENTE')) {
      soma++;
    }
  });
  return soma
}

function somaPorMes(linhas, mes, ref) {
  let soma = 0;
  linhas.forEach((linha) => {
    if (linha._rawData[2] === ref && linha._rawData[11] === mes) {
      soma++;
    }
  });
  return soma;
}