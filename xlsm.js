// Leitura do informe (.xlsx) e geração da planilha oficial COMAU (.xlsm) sem Excel.
// O .xlsm é editado direto no XML: só as células de lançamento mudam, e os valores
// das fórmulas são recalculados aqui para a planilha abrir certa até em pré-visualização.

const PRIMEIRA_LINHA = 12;
const ULTIMA_LINHA = 42;
const MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
  "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

export class ErroPlanilha extends Error {}

// ---------- utilidades ----------
const BASE_SERIAL = Date.UTC(1899, 11, 30);
export const serialData = (ano, mes, dia) => (Date.UTC(ano, mes - 1, dia) - BASE_SERIAL) / 86400000;
export const minutosDe = (hhmm) => { const [h, m] = hhmm.split(":").map(Number); return h * 60 + m; };
const fracao = (hhmm) => minutosDe(hhmm) / 1440;
const escXml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const desXml = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const pad = (n) => String(n).padStart(2, "0");

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
const MODELO_INVALIDO = "O arquivo não é uma planilha de horas COMAU (.xlsm). Escolha uma planilha oficial de qualquer mês.";

/**
 * Confere se o arquivo serve de modelo e extrai o que o app usa dele.
 * Qualquer planilha oficial já preenchida serve: os lançamentos são sobrescritos na geração.
 * Retorna { funcionario, centros, centroMaisUsado }.
 */
export async function lerModelo(JSZip, arquivo) {
  let zip;
  try { zip = await JSZip.loadAsync(arquivo); } catch { throw new ErroPlanilha(MODELO_INVALIDO); }
  const sheet = zip.file("xl/worksheets/sheet1.xml");
  const sstArq = zip.file("xl/sharedStrings.xml");
  if (!sheet || !sstArq || !zip.file("xl/workbook.xml")) throw new ErroPlanilha(MODELO_INVALIDO);
  const ws = new Planilha(await sheet.async("string"));
  if (ws.formula("G44") !== "SUM(G12:G42)*24" || ws.formula("B12") !== "D9") throw new ErroPlanilha(MODELO_INVALIDO);

  const sst = textosCompartilhados(await sstArq.async("string"));
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
  return { funcionario, centros, centroMaisUsado };
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
    const tipo = typeof v === "string" ? "str" : null;
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

/**
 * Gera a planilha oficial.
 * dados = { ano, mes, centroPadrao, registros: [{ data:'AAAA-MM-DD', inicio:'HH:MM', fim:'HH:MM', jantar, centro, obs }] }
 * Retorna { arquivo: Uint8Array, dias, totalHoras }.
 */
export async function gerarPlanilha(JSZip, modelo, dados) {
  const { ano, mes, registros } = dados;
  if (!modelo) throw new ErroPlanilha("Nenhum modelo de planilha configurado. Escolha um em Ajustes → Modelo da planilha.");
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
  let sstXml = await zip.file("xl/sharedStrings.xml").async("string");
  const sst = textosCompartilhados(sstXml);
  let usosNovos = 0;
  const indiceTexto = (texto) => {
    let i = sst.indexOf(texto);
    if (i < 0) {
      i = sst.push(texto) - 1;
      sstXml = sstXml.replace("</sst>", `<si><t xml:space="preserve">${escXml(texto)}</t></si></sst>`);
    }
    usosNovos++;
    return i;
  };

  // Cabeçalho do mês
  const primeiro = serialData(ano, mes, 1);
  ws.valor("C8", primeiro);
  ws.valor("D9", primeiro);

  // Lançamentos
  const porDia = Object.fromEntries(doMes.map((r) => [Number(r.data.slice(8, 10)), r]));
  const G = {}, I = {}, K = {}, C = {};
  for (let lin = PRIMEIRA_LINHA; lin <= ULTIMA_LINHA; lin++) {
    const r = porDia[lin - PRIMEIRA_LINHA + 1];
    const c = r ? fracao(r.inicio) : 0, d = r ? fracao(r.fim) : 0, e = r && r.jantar ? 60 / 1440 : 0;
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

  sstXml = sstXml
    .replace(/uniqueCount="\d+"/, `uniqueCount="${sst.length}"`)
    .replace(/ count="(\d+)"/, (_, n) => ` count="${Number(n) + usosNovos}"`);

  zip.file(sheetPath, ws.xml);
  zip.file("xl/sharedStrings.xml", sstXml);
  zip.file("xl/workbook.xml", wb);
  zip.file("docProps/app.xml", app);
  const arquivo = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { arquivo, dias: doMes.length, totalHoras: total };
}

/** Nome padrão do arquivo; o primeiro nome vem do campo "Funcionário" do modelo. */
export function nomeArquivo(ano, mes, funcionario) {
  const primeiro = String(funcionario || "").trim().split(/\s+/)[0] || "Funcionario";
  return `Tabela Horas_${primeiro}_Assistência-COMAU_${pad(mes)}_${ano}.xlsm`;
}
