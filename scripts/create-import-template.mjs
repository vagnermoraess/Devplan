import ExcelJS from "exceljs";

const workbook = new ExcelJS.Workbook();
workbook.creator = "Roadmap Portal";
workbook.created = new Date();

const sheet = workbook.addWorksheet("Demandas", {
  views: [{ state: "frozen", ySplit: 1 }],
});

sheet.columns = [
  { header: "Demanda", key: "demanda", width: 38 },
  { header: "Produto", key: "produto", width: 20 },
  { header: "Categoria", key: "categoria", width: 18 },
  { header: "Prioridade", key: "prioridade", width: 14 },
  { header: "Status", key: "status", width: 20 },
  { header: "Versão", key: "versao", width: 12 },
  { header: "Capacidade estimada", key: "capacidade", width: 22 },
  { header: "Data de início", key: "inicio", width: 18 },
  { header: "Prazo", key: "prazo", width: 18 },
  { header: "Observações", key: "observacoes", width: 45 },
];

sheet.addRows([
  {
    demanda: "Novo fluxo CIOT",
    produto: "Commerce",
    categoria: "Evolutiva",
    prioridade: "Alta",
    status: "Planejada",
    versao: "5.2",
    capacidade: 160,
    inicio: "01/07/2026",
    prazo: "30/09/2026",
    observacoes: "Linha de exemplo — substitua pelos dados reais",
  },
  {
    demanda: "Melhorias na API",
    produto: "Payments",
    categoria: "Técnica",
    prioridade: "Média",
    status: "Backlog",
    versao: "5.3",
    capacidade: 80,
    inicio: "15/07/2026",
    prazo: "20/09/2026",
    observacoes: "Produto deve estar previamente cadastrado e ativo",
  },
]);

const header = sheet.getRow(1);
header.height = 28;
header.font = { bold: true, color: { argb: "FFFFFFFF" } };
header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF246BFD" } };
header.alignment = { vertical: "middle", horizontal: "center" };

sheet.autoFilter = { from: "A1", to: "J3" };
sheet.getColumn("capacidade").numFmt = "0";

for (let row = 2; row <= 501; row += 1) {
  sheet.getCell(`C${row}`).dataValidation = {
    type: "list",
    allowBlank: true,
    formulae: ['"Estratégica,Evolutiva,Bug,Incidente,Sustentação,Regulatória,Técnica,Emergencial"'],
  };
  sheet.getCell(`D${row}`).dataValidation = {
    type: "list",
    allowBlank: true,
    formulae: ['"Baixa,Média,Alta,Crítica"'],
  };
  sheet.getCell(`E${row}`).dataValidation = {
    type: "list",
    allowBlank: true,
    formulae: ['"Backlog,Planejada,Em análise,Desenvolvimento,Em testes,Homologação,Concluída,Cancelada,Bloqueada"'],
  };
}

const instructions = workbook.addWorksheet("Instruções");
instructions.columns = [
  { header: "Campo", key: "campo", width: 24 },
  { header: "Obrigatório", key: "obrigatorio", width: 15 },
  { header: "Orientação", key: "orientacao", width: 80 },
];
instructions.addRows([
  ["Demanda", "Sim", "Nome da demanda."],
  ["Produto", "Sim", "Deve existir e estar ativo no cadastro de Produtos do portal."],
  ["Categoria", "Não", "Selecione um dos valores disponíveis na lista."],
  ["Prioridade", "Não", "Padrão adotado na ausência do valor: Média."],
  ["Status", "Não", "Padrão adotado na ausência do valor: Backlog."],
  ["Versão", "Não", "Versão relacionada à entrega, por exemplo 5.2."],
  ["Capacidade estimada", "Não", "Informe somente o número de horas, sem a letra h."],
  ["Data de início", "Não", "Utilize DD/MM/AAAA ou AAAA-MM-DD."],
  ["Prazo", "Não", "Utilize DD/MM/AAAA ou AAAA-MM-DD."],
  ["Observações", "Não", "Informações complementares."],
]);
const instructionHeader = instructions.getRow(1);
instructionHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
instructionHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF172D5B" } };
instructions.views = [{ state: "frozen", ySplit: 1 }];

await workbook.xlsx.writeFile("modelo-importacao-demandas.xlsx");
