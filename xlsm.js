// Leitura do informe (.xlsx) e geração das planilhas de horas (COMAU .xlsm e GME .xlsx) sem Excel.
// O arquivo é editado direto no XML: só as células de lançamento mudam, e os valores
// das fórmulas são recalculados aqui para a planilha abrir certa até em pré-visualização.

const PRIMEIRA_LINHA = 12;
const ULTIMA_LINHA = 42;
const MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
  "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
const MESES_GME = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
  "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
// Planilha GME: linhas dos dias 1 a 31 e limites dos períodos (em minutos), conforme as regras escritas nela.
const LINHA_GME = 9;
const ULTIMA_GME = 39;
const FORMULA_HORAS_GME = "((G9-F9)+(E9-D9)+(C9-B9))*24";
const MIN_13H = 780;
const MIN_13H30 = 810;
const MIN_18H = 1080;
const MIN_DIA = 1440;

export class ErroPlanilha extends Error {}

// ---------- utilidades ----------
const BASE_SERIAL = Date.UTC(1899, 11, 30);
export const serialData = (ano, mes, dia) => (Date.UTC(ano, mes - 1, dia) - BASE_SERIAL) / 86400000;
export const minutosDe = (hhmm) => { const [h, m] = hhmm.split(":").map(Number); return h * 60 + m; };
const fracao = (hhmm) => minutosDe(hhmm) / 1440;
const escXml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const desXml = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const pad = (n) => String(n).padStart(2, "0");

/** O modelo da GME é reconhecido pelo nome do arquivo. */
export const ehGME = (nome) => /gme/i.test(String(nome || ""));

/** Minutos de pausa de um lançamento: o almoço informado ou a pausa fixa de 60 min. */
export function pausaMinutos(r) {
  if (r.almocoSaida && r.almocoVolta) {
    const m = minutosDe(r.almocoVolta) - minutosDe(r.almocoSaida);
    return m < 0 ? m + MIN_DIA : m;
  }
  return r.jantar ? 60 : 0;
}

/** Horas trabalhadas (em horas decimais) de um lançamento, com virada de meia-noite. */
export function horasDoDia({ inicio, fim, jantar }) {
  if (!inicio || !fim) return 0;
  let min = minutosDe(fim) - minutosDe(inicio);
  if (min < 0) min += 1440;
  return (min - (jantar ? 60 : 0)) / 60;
}

function textosCompartilhados(xml) {
  if (!xml) return [];
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) =>
    [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => desXml(t[1])).join(""));
}

// ---------- leitura do informe ----------
function valoresDaPlanilha(sheetXml, sst) {
  const linhas = {};
  for (const m of sheetXml.matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const [, col, lin, attrs, corpo = ""] = m;
    const tipo = (attrs.match(/t="(\w+)"/) || [])[1];
    let valor = null;
    if (tipo === "inlineStr") valor = desXml((corpo.match(/<t[^>]*>([\s\S]*?)<\/t>/) || [, ""])[1]);
    else {
      const v = (corpo.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      if (v !== undefined) valor = tipo === "s" ? sst[Number(v)] : tipo === "str" ? desXml(v) : Number(v);
    }
    (linhas[lin] ||= {})[col] = valor;
  }
  return linhas;
}

function lerHora(valor, linha, nome) {
  if (typeof valor === "number") {
    const min = Math.round(valor * 1440) % 1440;
    return `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
  }
  const m = String(valor ?? "").trim().match(/^(\d{1,2}):(\d{2})/);
  if (!m || +m[1] > 23 || +m[2] > 59)
    throw new ErroPlanilha(`Informe, linha ${linha}: ${nome} inválido ('${valor ?? ""}'). Esperado no formato hh:mm.`);
  return `${pad(m[1])}:${m[2]}`;
}

function lerData(valor, linha) {
  let d;
  if (typeof valor === "number") d = new Date(BASE_SERIAL + Math.round(valor) * 86400000);
  else {
    const m = String(valor ?? "").trim().match(/^(\d{2})\/(\d{2})\/(\d{2})$/);
    if (m) {
      d = new Date(Date.UTC(2000 + +m[3], +m[2] - 1, +m[1]));
      if (d.getUTCDate() !== +m[1] || d.getUTCMonth() !== +m[2] - 1) d = null;
    }
  }
  if (!d) throw new ErroPlanilha(`Informe, linha ${linha}: Dia inválido ('${valor ?? ""}'). Esperado no formato dd/mm/aa.`);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}

/** Lê o informe do app de ponto. Retorna { ano, mes, registros: [{data, inicio, fim, jantar, taxa}] }. */
export async function lerInforme(JSZip, arquivo) {
  let zip;
  try { zip = await JSZip.loadAsync(arquivo); } catch {
    throw new ErroPlanilha("O arquivo escolhido não é um informe .xlsx válido.");
  }
  const sheet = zip.file("xl/worksheets/sheet1.xml");
  if (!sheet) throw new ErroPlanilha("O arquivo escolhido não é um informe .xlsx válido.");
  const sst = textosCompartilhados(await zip.file("xl/sharedStrings.xml")?.async("string"));
  const linhas = valoresDaPlanilha(await sheet.async("string"), sst);

  const cab = linhas["1"] || {};
  const norm = (v) => String(v ?? "").trim().toLowerCase();
  if (norm(cab.A) !== "dia" || norm(cab.D) !== "fim" || norm(cab.E) !== "pausa")
    throw new ErroPlanilha("O arquivo não está no formato do informe do app de ponto " +
      "(colunas esperadas: Dia, Projeto, Início, Fim, Pausa).");

  const registros = [], vistos = {};
  for (const n of Object.keys(linhas).map(Number).filter((n) => n > 1).sort((a, b) => a - b)) {
    const l = linhas[n];
    if (l.A === null || l.A === undefined || l.A === "") continue;
    const data = lerData(l.A, n);
    if (vistos[data]) {
      const [a, m, d] = data.split("-");
      throw new ErroPlanilha(`Informe, linhas ${vistos[data]} e ${n}: o dia ${d}/${m}/${a} aparece duas vezes.`);
    }
    vistos[data] = n;
    const pausa = Number(l.E ?? 0);
    if (pausa !== 0 && pausa !== 60)
      throw new ErroPlanilha(`Informe, linha ${n}: Pausa ${l.E} inválida. Use 0 (sem janta) ou 60 (com janta).`);
    const taxa = Number(l.G);
    registros.push({
      data, inicio: lerHora(l.C, n, "Início"), fim: lerHora(l.D, n, "Fim"),
      jantar: pausa === 60, taxa: Number.isFinite(taxa) ? taxa : null,
    });
  }
  if (!registros.length) throw new ErroPlanilha("O informe não tem nenhum dia lançado.");
  const meses = [...new Set(registros.map((r) => r.data.slice(0, 7)))].sort();
  if (meses.length > 1) {
    const lista = meses.map((m) => `${m.slice(5)}/${m.slice(0, 4)}`).join(", ");
    throw new ErroPlanilha(`O informe tem dias de mais de um mês (${lista}). Exporte um informe por mês.`);
  }
  return { ano: +meses[0].slice(0, 4), mes: +meses[0].slice(5), registros };
}

// ---------- modelo ----------
const MODELO_INVALIDO = "O arquivo não é uma planilha de horas conhecida. Escolha uma planilha oficial da COMAU (.xlsm) ou da GME (.xlsx).";
const MODELO_GME_INVALIDO = "O arquivo tem GME no nome, mas não está no formato da planilha de horas da GME.";

/**
 * Confere se o arquivo serve de modelo e extrai o que o app usa dele.
 * Qualquer planilha oficial já preenchida serve: os lançamentos são sobrescritos na geração.
 * O tipo vem do nome do arquivo: com "GME" no nome, é conferido no formato da GME.
 * Retorna { tipo, funcionario, centros, centroMaisUsado }.
 */
export async function lerModelo(JSZip, arquivo, nome = "") {
  let zip;
  try { zip = await JSZip.loadAsync(arquivo); } catch { throw new ErroPlanilha(MODELO_INVALIDO); }
  const sheet = zip.file("xl/worksheets/sheet1.xml");
  const sstArq = zip.file("xl/sharedStrings.xml");
  if (!sheet || !sstArq || !zip.file("xl/workbook.xml")) throw new ErroPlanilha(MODELO_INVALIDO);
  const ws = new Planilha(await sheet.async("string"));
  const sst = textosCompartilhados(await sstArq.async("string"));
  const formatoGME = ws.formula("L3") === `SUM(H${LINHA_GME}:H${ULTIMA_GME})` && ws.formula("H9") === FORMULA_HORAS_GME;
  if (ehGME(nome)) {
    if (!formatoGME) throw new ErroPlanilha(MODELO_GME_INVALIDO);
    return { tipo: "gme", funcionario: String(ws.atual("F2", sst) || "").trim(), centros: [], centroMaisUsado: null };
  }
  if (formatoGME) throw new ErroPlanilha("Esta planilha está no formato da GME, mas não tem GME no nome do arquivo. Renomeie o arquivo incluindo GME e escolha de novo.");
  if (ws.formula("G44") !== "SUM(G12:G42)*24" || ws.formula("B12") !== "D9") throw new ErroPlanilha(MODELO_INVALIDO);

  const centros = [];
  for (let lin = 47; lin <= 63; lin++) {
    const v = ws.atual(`B${lin}`, sst);
    if (typeof v === "string" && v.trim() && v.trim() !== "--" && !centros.includes(v)) centros.push(v);
  }
  if (!centros.length) throw new ErroPlanilha("O modelo não tem centros de custo na \"Legenda de projetos\" (B47 em diante).");
  const usos = {};
  for (let lin = PRIMEIRA_LINHA; lin <= ULTIMA_LINHA; lin++) {
    const v = ws.atual(`I${lin}`, sst);
    if (centros.includes(v)) usos[v] = (usos[v] || 0) + 1;
  }
  const centroMaisUsado = Object.entries(usos).sort((a, b) => b[1] - a[1])[0]?.[0] || centros[0];
  const funcionario = String(ws.atual("H5", sst) || "").trim();
  return { tipo: "comau", funcionario, centros, centroMaisUsado };
}

// ---------- geração da planilha ----------
class Planilha {
  constructor(xml) { this.xml = xml; }

  _re(ref) { return new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`); }

  _ler(ref) {
    const m = this.xml.match(this._re(ref));
    if (!m) return null;
    return { attrs: m[1].replace(/\s*t="\w+"/, ""), corpo: m[2] || "" };
  }

  _gravar(ref, attrs, tipo, corpo) {
    const t = tipo ? ` t="${tipo}"` : "";
    const novo = corpo ? `<c r="${ref}"${attrs}${t}>${corpo}</c>` : `<c r="${ref}"${attrs}${t}/>`;
    this.xml = this.xml.replace(this._re(ref), () => novo);
  }

  /** Cria uma célula ausente no modelo, na ordem certa da linha, com o estilo da linha de cima. */
  _criar(ref) {
    const [, col, lin] = ref.match(/^([A-Z]+)(\d+)$/);
    const acima = this._ler(`${col}${Number(lin) - 1}`);
    const attrs = acima ? acima.attrs : "";
    const linha = this.xml.match(new RegExp(`<row r="${lin}"[^>]*>[\\s\\S]*?</row>`));
    if (!linha) throw new ErroPlanilha(`Modelo da planilha inválido: linha ${lin} não encontrada.`);
    const proxima = [...linha[0].matchAll(/<c r="([A-Z]+)\d+"/g)].find((m) => m[1].length > col.length
      || (m[1].length === col.length && m[1] > col));
    const nova = `<c r="${ref}"${attrs}/>`;
    const novaLinha = proxima
      ? linha[0].slice(0, proxima.index) + nova + linha[0].slice(proxima.index)
      : linha[0].replace("</row>", nova + "</row>");
    this.xml = this.xml.replace(linha[0], () => novaLinha);
    return { attrs, corpo: "" };
  }

  /** Valor de constante (número, índice de texto compartilhado ou vazio). */
  valor(ref, v, { compartilhado = false } = {}) {
    let c = this._ler(ref);
    if (!c) {
      if (v === null || v === undefined || v === "") return;
      c = this._criar(ref);
    }
    if (v === null || v === undefined || v === "") this._gravar(ref, c.attrs, null, "");
    else this._gravar(ref, c.attrs, compartilhado ? "s" : null, `<v>${v}</v>`);
  }

  /** Atualiza o valor calculado de uma célula com fórmula, mantendo a fórmula. */
  calculado(ref, v) {
    const c = this._ler(ref);
    const f = c && (c.corpo.match(/<f[^>]*\/>|<f[^>]*>[\s\S]*?<\/f>/) || [])[0];
    if (!f) throw new ErroPlanilha(`Modelo da planilha inválido: fórmula em ${ref} não encontrada.`);
    const tipo = typeof v !== "string" ? null : v.startsWith("#") ? "e" : "str";
    this._gravar(ref, c.attrs, tipo, `${f}<v>${typeof v === "string" ? escXml(v) : v}</v>`);
  }

  formula(ref) {
    const c = this._ler(ref);
    return c && (c.corpo.match(/<f[^>]*>([\s\S]*?)<\/f>/) || [])[1];
  }

  /** Valor atual (calculado ou constante) já convertido. */
  atual(ref, sst) {
    const m = this.xml.match(this._re(ref));
    if (!m) return null;
    const tipo = (m[1].match(/t="(\w+)"/) || [])[1];
    const v = ((m[2] || "").match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    if (v === undefined) return null;
    if (tipo === "s") return sst[Number(v)];
    if (tipo === "str" || tipo === "e") return desXml(v);
    return Number(v);
  }
}

/** Textos compartilhados do arquivo, acrescentando os novos sem repetir. */
function criarTextos(xmlInicial) {
  let xml = xmlInicial, usosNovos = 0;
  const lista = textosCompartilhados(xml);
  return {
    lista,
    indice(texto) {
      let i = lista.indexOf(texto);
      if (i < 0) {
        i = lista.push(texto) - 1;
        xml = xml.replace("</sst>", `<si><t xml:space="preserve">${escXml(texto)}</t></si></sst>`);
      }
      usosNovos++;
      return i;
    },
    xmlFinal: () => xml
      .replace(/uniqueCount="\d+"/, `uniqueCount="${lista.length}"`)
      .replace(/ count="(\d+)"/, (_, n) => ` count="${Number(n) + usosNovos}"`),
  };
}

/**
 * Gera a planilha oficial.
 * dados = { tipo, ano, mes, centroPadrao, commessaPadrao,
 *           registros: [{ data:'AAAA-MM-DD', inicio:'HH:MM', fim:'HH:MM', jantar, almocoSaida, almocoVolta, centro, commessa, obs }] }
 * No tipo "gme", registros deve trazer todos os dias: o fim de um turno do mês anterior pode cair no dia 1.
 * Retorna { arquivo: Uint8Array, dias, totalHoras }.
 */
export async function gerarPlanilha(JSZip, modelo, dados) {
  const { ano, mes, registros } = dados;
  if (!modelo) throw new ErroPlanilha("Nenhum modelo de planilha configurado. Escolha um em Ajustes → Modelo da planilha.");
  if (dados.tipo === "gme") return gerarGME(JSZip, modelo, dados);
  const doMes = registros.filter((r) => r.data.startsWith(`${ano}-${pad(mes)}-`));
  if (!doMes.length) throw new ErroPlanilha(`Não há dias lançados em ${pad(mes)}/${ano}.`);
  for (const r of doMes) {
    const dia = r.data.slice(8, 10) + "/" + r.data.slice(5, 7);
    if (!r.inicio) throw new ErroPlanilha(`O dia ${dia} está sem horário de entrada.`);
    if (!r.fim) throw new ErroPlanilha(`O dia ${dia} está sem horário de saída.`);
  }

  const zip = await JSZip.loadAsync(modelo);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const ws = new Planilha(await zip.file(sheetPath).async("string"));
  const textos = criarTextos(await zip.file("xl/sharedStrings.xml").async("string"));
  const sst = textos.lista;
  const indiceTexto = (texto) => textos.indice(texto);

  // Cabeçalho do mês
  const primeiro = serialData(ano, mes, 1);
  ws.valor("C8", primeiro);
  ws.valor("D9", primeiro);

  // Lançamentos
  const porDia = Object.fromEntries(doMes.map((r) => [Number(r.data.slice(8, 10)), r]));
  const G = {}, I = {}, K = {}, C = {};
  for (let lin = PRIMEIRA_LINHA; lin <= ULTIMA_LINHA; lin++) {
    const r = porDia[lin - PRIMEIRA_LINHA + 1];
    const c = r ? fracao(r.inicio) : 0, d = r ? fracao(r.fim) : 0, e = r ? pausaMinutos(r) / 1440 : 0;
    const centro = r ? (r.centro || dados.centroPadrao) : null;
    if (r && !centro) throw new ErroPlanilha(`O dia ${pad(lin - PRIMEIRA_LINHA + 1)}/${pad(mes)} está sem centro de custo.`);
    ws.valor(`C${lin}`, r ? c : null);
    ws.valor(`D${lin}`, r ? d : null);
    ws.valor(`E${lin}`, e || null);
    ws.valor(`H${lin}`, r && r.obs ? indiceTexto(r.obs) : null, { compartilhado: true });
    ws.valor(`I${lin}`, centro ? indiceTexto(centro) : null, { compartilhado: true });

    const b = primeiro + (lin - PRIMEIRA_LINHA);
    ws.calculado(`B${lin}`, b);
    G[lin] = (d - c + (d < c ? 1 : 0)) - e;
    ws.calculado(`G${lin}`, G[lin]);
    K[lin] = c > 0 ? b : "--";
    ws.calculado(`K${lin}`, K[lin]);
    C[lin] = c;
    I[lin] = centro;
  }

  // Totais
  let soma = 0;
  for (let lin = PRIMEIRA_LINHA; lin <= ULTIMA_LINHA; lin++) soma += G[lin];
  const total = soma * 24;
  ws.calculado("G44", total);
  ws.calculado("I43", total);

  // Resumo: Menor/Maior data e dias trabalhados
  const datas = Object.values(K).filter((v) => typeof v === "number");
  ws.calculado("N2", datas.length ? Math.min(...datas) : 0);
  ws.calculado("N3", datas.length ? Math.max(...datas) : 0);
  ws.calculado("N4", Object.values(C).filter((v) => v > 0).length);

  // Resumo por projeto: SUMIF($I$12:I<n>, M<r>, $G$12:G<n>)*24*E44
  const ajuste = ws.atual("E44", sst) ?? 1;
  const celula = (col, lin) => {
    if (lin >= PRIMEIRA_LINHA && lin <= ULTIMA_LINHA) return col === "I" ? I[lin] : G[lin];
    return ws.atual(`${col}${lin}`, sst);
  };
  const resumo = {};
  for (let lin = 7; lin <= 24; lin++) {
    const f = ws.formula(`N${lin}`);
    const m = f && f.match(/^SUMIF\(\$I\$12:I(\d+),M(\d+),\$G\$12:G(\d+)\)\*24\*E44$/);
    if (!m) continue;
    const crit = ws.atual(`M${m[2]}`, sst);
    let s = 0;
    for (let l = PRIMEIRA_LINHA; l <= Number(m[1]); l++) {
      const iv = celula("I", l), gv = celula("G", l);
      const casa = typeof crit === "string"
        ? typeof iv === "string" && iv.toLowerCase() === crit.toLowerCase()
        : typeof iv === "number" && iv === crit;
      if (casa && typeof gv === "number") s += gv;
    }
    resumo[lin] = s * 24 * ajuste;
    ws.calculado(`N${lin}`, resumo[lin]);
  }
  // N25 = SUM(N7:N24); legenda D47:D63 = N8:N24 (fórmula compartilhada a partir de D47 = N8)
  ws.calculado("N25", Object.values(resumo).reduce((a, b) => a + b, 0));
  for (let lin = 47; lin <= 63; lin++) {
    if (resumo[lin - 39] !== undefined) ws.calculado(`D${lin}`, resumo[lin - 39]);
  }

  // Nome da aba
  const nomeAba = `${MESES[mes - 1]} ${ano}`;
  let wb = await zip.file("xl/workbook.xml").async("string");
  wb = wb.replace(/(<sheet [^>]*name=")[^"]*(")/, `$1${escXml(nomeAba)}$2`);
  let app = await zip.file("docProps/app.xml").async("string");
  app = app.replace(/(<TitlesOfParts>[\s\S]*?<vt:lpstr>)[^<]*(<\/vt:lpstr>)/, `$1${escXml(nomeAba)}$2`);

  zip.file(sheetPath, ws.xml);
  zip.file("xl/sharedStrings.xml", textos.xmlFinal());
  zip.file("xl/workbook.xml", wb);
  zip.file("docProps/app.xml", app);
  const arquivo = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { arquivo, dias: doMes.length, totalHoras: total };
}

// ---------- planilha GME ----------
const diaSeguinte = (data) => {
  const [a, m, d] = data.split("-").map(Number);
  return new Date(Date.UTC(a, m - 1, d + 1)).toISOString().slice(0, 10);
};
const diaCurto = (data) => `${data.slice(8, 10)}/${data.slice(5, 7)}`;

/** Trechos trabalhados de um turno, em minutos desde as 00:00 da data de entrada (passa de 1440 após a meia-noite). */
function trechosDoTurno(r) {
  const ini = minutosDe(r.inicio);
  const relativo = (h) => (minutosDe(h) < ini ? minutosDe(h) + MIN_DIA : minutosDe(h));
  const fim = relativo(r.fim);
  if (!r.almocoSaida) return [[ini, fim]];
  return [[ini, relativo(r.almocoSaida)], [relativo(r.almocoVolta), fim]];
}

/**
 * Distribui os trechos de um dia nos pares manhã, tarde e noite (índices 0, 1 e 2).
 * Divide nas viradas da planilha (13:00 e 18:00); se assim não couber em um par por período,
 * mantém os trechos inteiros, em ordem, para o total do dia continuar certo.
 */
function distribuirPeriodos(trechos, data) {
  const pares = [null, null, null];
  let coube = true;
  const colocar = (i, a, b) => { if (pares[i]) coube = false; else pares[i] = [a, b]; };
  for (const [a, b] of trechos) {
    if (a < MIN_18H) {
      const fimDia = Math.min(b, MIN_18H);
      if (a >= MIN_13H) colocar(1, a, fimDia);
      else if (fimDia <= MIN_13H30) colocar(0, a, fimDia);
      else { colocar(0, a, MIN_13H); colocar(1, MIN_13H, fimDia); }
    }
    if (b > MIN_18H) colocar(2, Math.max(a, MIN_18H), b);
  }
  if (coube) return pares;

  if (trechos.length > 3) throw new ErroPlanilha(`O dia ${diaCurto(data)} tem mais trechos de trabalho do que os 3 períodos da planilha.`);
  const indices = trechos.map(([a]) => (a < MIN_13H ? 0 : a < MIN_18H ? 1 : 2));
  for (let i = 1; i < indices.length; i++) if (indices[i] <= indices[i - 1]) indices[i] = indices[i - 1] + 1;
  if (indices.at(-1) > 2) {
    indices[indices.length - 1] = 2;
    for (let i = indices.length - 2; i >= 0; i--) if (indices[i] >= indices[i + 1]) indices[i] = indices[i + 1] - 1;
  }
  const inteiros = [null, null, null];
  trechos.forEach((t, i) => { inteiros[indices[i]] = t; });
  return inteiros;
}

async function gerarGME(JSZip, modelo, dados) {
  const { ano, mes, registros } = dados;
  const prefixo = `${ano}-${pad(mes)}-`;
  const doMes = registros.filter((r) => r.data.startsWith(prefixo));
  if (!doMes.length) throw new ErroPlanilha(`Não há dias lançados em ${pad(mes)}/${ano}.`);
  for (const r of doMes) {
    if (!r.inicio) throw new ErroPlanilha(`O dia ${diaCurto(r.data)} está sem horário de entrada.`);
    if (!r.fim) throw new ErroPlanilha(`O dia ${diaCurto(r.data)} está sem horário de saída.`);
  }

  // Cada turno é dividido nas datas em que foi trabalhado; o fim do turno de ontem vai para a linha de hoje.
  const trechosPorData = {}, registroPorData = {};
  for (const r of registros) {
    if (!r.inicio || !r.fim) continue;
    const trechos = trechosDoTurno(r);
    const usado = r.data.startsWith(prefixo) || (diaSeguinte(r.data).startsWith(prefixo) && trechos.at(-1)[1] > MIN_DIA);
    if (!usado) continue;
    if (r.jantar && !r.almocoSaida)
      throw new ErroPlanilha(`O dia ${diaCurto(r.data)} tem pausa de 60 min sem os horários do almoço. Toque nele em Registros e informe a saída e a volta.`);
    if (r.almocoSaida && !r.almocoVolta)
      throw new ErroPlanilha(`O dia ${diaCurto(r.data)} está sem a volta do almoço.`);
    const seguinte = diaSeguinte(r.data);
    const anotar = (data, a, b) => {
      if (b <= a || !data.startsWith(prefixo)) return;
      (trechosPorData[data] ||= []).push([a, b]);
      if (data === r.data || !registroPorData[data]) registroPorData[data] = r;
    };
    for (const [a, b] of trechos) {
      anotar(r.data, a, Math.min(b, MIN_DIA));
      anotar(seguinte, Math.max(a, MIN_DIA) - MIN_DIA, b - MIN_DIA);
    }
  }

  const zip = await JSZip.loadAsync(modelo);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const ws = new Planilha(await zip.file(sheetPath).async("string"));
  const textos = criarTextos(await zip.file("xl/sharedStrings.xml").async("string"));

  // Cabeçalho do mês
  const nomeMes = MESES_GME[mes - 1];
  ws.valor("B2", textos.indice(nomeMes), { compartilhado: true });
  ws.valor("B3", ano);
  ws.calculado("B6", `1/${nomeMes}/${ano}`);

  // Lançamentos e valores das fórmulas
  const primeiro = serialData(ano, mes, 1);
  const diasNoMes = new Date(Date.UTC(ano, mes, 0)).getUTCDate();
  const H = {};
  for (let lin = LINHA_GME; lin <= ULTIMA_GME; lin++) {
    const dia = lin - LINHA_GME + 1;
    const data = `${prefixo}${pad(dia)}`;
    const trechos = dia <= diasNoMes ? (trechosPorData[data] || []).sort((x, y) => x[0] - y[0]) : [];
    const pares = trechos.length ? distribuirPeriodos(trechos, data) : [null, null, null];
    const v = {};
    [["B", "C"], ["D", "E"], ["F", "G"]].forEach(([ent, sai], i) => {
      v[ent] = pares[i] ? pares[i][0] / MIN_DIA : 0;
      v[sai] = pares[i] ? pares[i][1] / MIN_DIA : 0;
      ws.valor(`${ent}${lin}`, pares[i] ? v[ent] : null);
      ws.valor(`${sai}${lin}`, pares[i] ? v[sai] : null);
    });

    const r = registroPorData[data];
    const commessa = r ? String(r.commessa ?? dados.commessaPadrao ?? "").trim() : "";
    if (/^\d+$/.test(commessa)) ws.valor(`I${lin}`, Number(commessa));
    else ws.valor(`I${lin}`, commessa ? textos.indice(commessa) : null, { compartilhado: true });
    const obs = r && r.data === data ? r.obs : "";
    ws.valor(`K${lin}`, obs ? textos.indice(obs) : null, { compartilhado: true });

    const a = lin < ULTIMA_GME || diasNoMes === 31 ? primeiro + (lin - LINHA_GME) : "";
    ws.calculado(`A${lin}`, a);
    ws.calculado(`M${lin}`, a === "" ? "#VALUE!" : new Date(BASE_SERIAL + a * 86400000).getUTCDay() || 7);
    H[lin] = ((v.G - v.F) + (v.E - v.D) + (v.C - v.B)) * 24;
    ws.calculado(`H${lin}`, H[lin]);
  }

  // Totais: geral, canteiro e sede
  let total = 0, canteiro = 0, sede = 0;
  for (let lin = LINHA_GME; lin <= ULTIMA_GME; lin++) {
    total += H[lin];
    const local = String(ws.atual(`J${lin}`, textos.lista) ?? "").toUpperCase();
    if (local === "CANTEIRO") canteiro += H[lin];
    if (local === "SEDE") sede += H[lin];
  }
  ws.calculado("L3", total);
  ws.calculado("L4", canteiro);
  ws.calculado("L5", sede);

  zip.file(sheetPath, ws.xml);
  zip.file("xl/sharedStrings.xml", textos.xmlFinal());
  const arquivo = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { arquivo, dias: doMes.length, totalHoras: total };
}

/** Nome padrão do arquivo; o primeiro nome vem do campo "Funcionário" do modelo. */
export function nomeArquivo(ano, mes, funcionario, tipo = "comau") {
  const primeiro = String(funcionario || "").trim().split(/\s+/)[0] || "Funcionario";
  return tipo === "gme"
    ? `Tabela Horas_${primeiro}_Projeto-GME_${pad(mes)}_${ano}.xlsx`
    : `Tabela Horas_${primeiro}_Assistência-COMAU_${pad(mes)}_${ano}.xlsm`;
}
