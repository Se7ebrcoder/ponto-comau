import { ErroPlanilha, PERIODOS_NGC, ehGME, ehNGC, gerarPlanilha, lerModelo, minutosDe, nomeArquivo, pausaMinutos, trechosDosPeriodos, batidasNGC } from "./xlsm.js";
import { desenharComprovante } from "./comprovante.js";

const CHAVE = "ponto-comau.v1";
const CHAVE_MODELO = "ponto-comau.modelo";
const CHAVE_MODELO_GME = "ponto-comau.modelo-gme";
const CHAVE_MODELO_NGC = "ponto-comau.modelo-ngc";
const CLIENTES = ["comau", "gme", "ngc"];
const NOME_CLIENTE = { comau: "COMAU", gme: "GME", ngc: "NGC" };
const CHAVES_MODELO = { comau: CHAVE_MODELO, gme: CHAVE_MODELO_GME, ngc: CHAVE_MODELO_NGC };
const CHAVE_TEMA = "ponto-comau.tema";
const MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
const SEMANA = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
const MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12";
const MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ROTULO_GERAR = "Gerar planilha de horas";

const $ = (sel, raiz = document) => raiz.querySelector(sel);
const pad = (n) => String(n).padStart(2, "0");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);

// ---------- Datas ----------
const chaveData = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const dataDe = (chave) => { const [a, m, d] = chave.split("-").map(Number); return new Date(a, m - 1, d); };
const horaAgora = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const curta = (chave) => `${chave.slice(8, 10)}/${chave.slice(5, 7)}`;
const horaOk = (s) => /^([01]\d|2[0-3]):[0-5]\d$/.test(s || "");

// Máscara do campo de hora digitado: "1515" vira "15:15".
function mascaraHora(el) {
  const d = el.value.replace(/\D/g, "").slice(0, 4);
  el.value = d.length > 2 ? `${d.slice(0, 2)}:${d.slice(2)}` : d;
}
// Ao sair do campo completa formas curtas: "9" → 09:00, "15" → 15:00, "930" → 09:30.
function completarHora(el) {
  const d = el.value.replace(/\D/g, "");
  if (!d) return;
  const n = d.length <= 2 ? d.padStart(2, "0") + "00" : d.padStart(4, "0");
  el.value = `${n.slice(0, 2)}:${n.slice(2, 4)}`;
}

// ---------- Cálculos ----------
// Enquanto o almoço está sendo digitado, só vale quando os dois horários estão completos.
const pausaDe = (r) => pausaMinutos(horaOk(r.almocoSaida) && horaOk(r.almocoVolta) ? r : { jantar: r.jantar });
// NGC: os seis campos (De/Até de manhã, tarde e noite) são batidas em ordem, pulando as vazias,
// que formam pares entrada/saída: De 08:00 na manhã, 11:00–12:00 na tarde e Até 20:00 na noite
// = 08:00–11:00 e 12:00–20:00 (almoço de 11 a 12). Nenhum campo é obrigatório; o vazio vira 00:00
// na planilha e não conta. 00:00 digitado é meia-noite. Horário incompleto (digitando) conta como vazio.
const periodosVazios = () => [["", ""], ["", ""], ["", ""]];
const soValidos = (periodos) => periodos.map((p) => p.map((h) => (horaOk(h) ? h : "")));
const trechosValidos = (periodos) => trechosDosPeriodos(soValidos(periodos));
// Pares entrada/saída em texto ("08:00–11:00"), mais a batida sem par, se houver.
function paresNGC(periodos) {
  const b = batidasNGC(soValidos(periodos)), pares = [];
  for (let k = 0; k < b.length; k += 2) pares.push({ de: b[k].h, ate: b[k + 1]?.h || "" });
  return pares;
}
// Início e fim do ponto da NGC (lista e sobreposição): da primeira batida à última saída.
// Sem nenhum par, a única batida. Na NGC o ponto nunca fica "sem fim".
function resumoPeriodos(periodos) {
  const pares = paresNGC(periodos), completos = pares.filter((p) => p.ate);
  const inicio = pares[0]?.de || "";
  return { inicio, fim: completos.at(-1)?.ate || inicio };
}
const textoPeriodos = (r) => paresNGC(r.periodos || []).map(({ de, ate }) => `${de}–${ate || "--:--"}`).join(" · ");
function minutosTrabalhados(r) {
  if (r.periodos) return trechosValidos(r.periodos).reduce((s, { a, b }) => s + (b > a ? b - a : 0), 0);
  if (!horaOk(r.inicio) || !horaOk(r.fim)) return 0;
  let min = minutosDe(r.fim) - minutosDe(r.inicio);
  if (min < 0) min += 1440;
  return Math.max(0, min - pausaDe(r));
}
const fmtHM = (min) => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
const viraDia = (r) => (r.periodos ? trechosValidos(r.periodos).some(({ b }) => b > 1440)
  : horaOk(r.inicio) && horaOk(r.fim) && minutosDe(r.fim) < minutosDe(r.inicio));


// ---------- Backup ----------
const ARQUIVO_BACKUP = "backup-pontos.json";
const PASTA_BACKUP = "Backup Pontos";
const INTERVALO_BACKUP = 24 * 60 * 60 * 1000;
const ADIAMENTO_BACKUP = 6 * 60 * 60 * 1000; // "Depois": o lembrete volta em 6 horas
// No Android, o navegador pode gravar direto numa pasta; se não puder, o arquivo vai para Downloads.
const ehAndroid = /Android/i.test(navigator.userAgent);
const temPastaNoAparelho = ehAndroid && !!window.showDirectoryPicker;

// ---------- Armazenamento (só neste aparelho) ----------
const PADRAO = { registros: {}, config: { centroPadrao: null, jantarPadrao: true, commessaPadrao: "", centrosExtras: [], commessasExtras: [], ultimoBackup: null, lembreteBackup: null, alteradoEm: null, pastaBackup: null, pastaFalhou: false } };
let dados = carregar();
let modelos = carregarModelos(); // { comau, gme, ngc }: cada um { arquivo, importadoEm, funcionario, centros, base64 } ou null

// Cada cliente tem seus pontos: o mesmo dia pode ter um ponto de cada cliente.
function chaveRegistro(data, cliente) { return cliente && cliente !== "comau" ? `${data}#${cliente}` : data; }

// Pontos antigos não têm cliente: períodos indicam NGC; almoço ou commessa, GME; os demais são da COMAU.
function normalizarRegistros(registros) {
  const saida = {};
  for (const { taxa, ...r } of Object.values(registros)) {
    const cliente = r.cliente || (r.periodos ? "ngc" : r.almocoSaida || r.commessa ? "gme" : "comau");
    // NGC: pontos salvos antes sem o Até deixam de ficar "sem fim".
    saida[chaveRegistro(r.data, cliente)] = { ...r, cliente, ...(r.periodos ? resumoPeriodos(r.periodos) : {}) };
  }
  return saida;
}

// O valor hora foi removido do app (gerava conflito entre os funcionários): nenhum preço fica guardado.
function semPreco(config) {
  const { taxaPadrao, taxaPadraoGME, taxaPadraoNGC, ...resto } = config;
  return resto;
}

function carregar() {
  try {
    const salvo = JSON.parse(localStorage.getItem(CHAVE));
    if (salvo && salvo.registros) return { registros: normalizarRegistros(salvo.registros), config: semPreco({ ...PADRAO.config, ...salvo.config }) };
  } catch { /* sem dados ou armazenamento bloqueado */ }
  return structuredClone(PADRAO);
}

function carregarModelos() {
  const ler = (chave) => { try { return JSON.parse(localStorage.getItem(chave)); } catch { return null; } };
  const m = { comau: ler(CHAVE_MODELO), gme: ler(CHAVE_MODELO_GME), ngc: ler(CHAVE_MODELO_NGC) };
  // Antes havia um modelo só, guardado sempre na mesma chave, que podia ser o da GME.
  if (ehGME(m.comau?.arquivo)) {
    m.gme = m.gme || m.comau;
    m.comau = null;
    try { localStorage.setItem(CHAVE_MODELO_GME, JSON.stringify(m.gme)); localStorage.removeItem(CHAVE_MODELO); } catch { /* tenta de novo na próxima abertura */ }
  }
  return m;
}

const clientesComModelo = () => CLIENTES.filter((c) => modelos[c]);
const variosClientes = () => clientesComModelo().length > 1;
function clienteInicial() {
  const salvo = dados.config.cliente;
  if (salvo && modelos[salvo]) return salvo;
  return clientesComModelo()[0] || "comau";
}

// No cliente GME, o ponto pede almoço e commessa; na NGC, os períodos manhã, tarde e noite.
const modeloGME = () => ui.cliente === "gme";
const modeloNGC = () => ui.cliente === "ngc";

// Opções do ponto: as da legenda de cada modelo mais as adicionadas à mão no app (Ajustes ou no próprio campo).
const NOVO_ITEM = "__novo__";
const centroPadrao = () => dados.config.centroPadrao || modelos.comau?.centros?.[0] || "";
const listaCentros = () => [...(modelos.comau?.centros || []), ...(dados.config.centrosExtras || [])];
const listaCommessas = () => [...(modelos.gme?.commessas || []), ...(dados.config.commessasExtras || [])];
function rotuloCommessa(codigo) {
  const item = listaCommessas().find((c) => c.codigo === String(codigo));
  return item?.linha ? `${item.codigo} – ${item.linha}` : String(codigo);
}
const opcaoHtml = (valor, rotulo, atual) => `<option value="${esc(valor)}" ${valor === atual ? "selected" : ""}>${esc(rotulo)}</option>`;
function campoCentro(id, atual, rotulo) {
  const lista = [...new Set([...listaCentros(), atual].filter(Boolean))];
  return `<select id="${id}" class="opcao" data-atual="${esc(atual || "")}" aria-label="${rotulo}">${lista.map((c) => opcaoHtml(c, c.trim(), atual)).join("")}
      ${opcaoHtml(NOVO_ITEM, "+ Adicionar centro de custo…", "")}</select>`;
}
function campoCommessa(id, atual, rotulo) {
  const lista = listaCommessas();
  // Sem tabela no modelo, a commessa é digitada.
  if (!lista.length) return `<input id="${id}" class="numero" inputmode="numeric" maxlength="12" value="${esc(atual)}" placeholder="—" aria-label="${rotulo}">`;
  const valor = String(atual ?? "");
  const conhecida = !valor || lista.some((c) => c.codigo === valor);
  return `<select id="${id}" class="opcao" data-atual="${esc(valor)}" aria-label="${rotulo}">
      ${opcaoHtml("", "Sem commessa", valor)}
      ${lista.map((c) => opcaoHtml(c.codigo, rotuloCommessa(c.codigo), valor)).join("")}
      ${conhecida ? "" : opcaoHtml(valor, valor, valor)}
      ${opcaoHtml(NOVO_ITEM, "+ Adicionar commessa…", "")}
    </select>`;
}

// alterou: false quando o que mudou não precisa de um novo backup (ex.: a data do próprio backup).
function salvar({ alterou = true } = {}) {
  try {
    if (alterou) dados.config.alteradoEm = new Date().toISOString();
    localStorage.setItem(CHAVE, JSON.stringify(dados));
    return true;
  } catch {
    mostrarErro("Não foi possível salvar", "O armazenamento do aparelho está cheio ou bloqueado. Libere espaço e tente novamente.");
    return false;
  }
}

async function pedirPersistencia() {
  try { if (navigator.storage?.persist && !(await navigator.storage.persisted())) await navigator.storage.persist(); } catch { /* opcional */ }
}

function paraBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let texto = "";
  for (let i = 0; i < bytes.length; i += 0x8000) texto += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(texto);
}
const deBase64 = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer;

const registrosDoMes = (ano, mes) => Object.values(dados.registros)
  .filter((r) => r.cliente === ui.cliente && r.data.startsWith(`${ano}-${pad(mes)}-`))
  .sort((a, b) => a.data.localeCompare(b.data));

// ---------- Tema: automático (segue o aparelho), claro ou escuro ----------
const midiaEscura = matchMedia("(prefers-color-scheme: dark)");
function temaEscolhido() {
  try { return localStorage.getItem(CHAVE_TEMA) || "auto"; } catch { return "auto"; }
}
const temaEfetivo = () => { const t = temaEscolhido(); return t === "auto" ? (midiaEscura.matches ? "dark" : "light") : t; };
function aplicarTema() {
  const t = temaEfetivo();
  document.documentElement.dataset.theme = t;
  $('meta[name="theme-color"]').content = t === "light" ? "#f6f4f4" : "#0a0a0b";
}
function escolherTema(t) {
  try { localStorage.setItem(CHAVE_TEMA, t); } catch { /* vale só nesta sessão */ }
  aplicarTema();
  if (ui.aba === "ajustes") render();
}
midiaEscura.addEventListener("change", () => { if (temaEscolhido() === "auto") aplicarTema(); });

// ---------- Estado da interface ----------
const hoje = new Date();
const ui = { aba: "registrar", ano: hoje.getFullYear(), mes: hoje.getMonth() + 1, cliente: clienteInicial() };

// Formulário da tela Registrar: sempre começa um ponto novo. `original` é a data do
// registro em edição, aberto pela lista de Registros (ou null se novo).
let form;
const formNovo = (data = chaveData(new Date())) => ({
  data, inicio: "", fim: "", jantar: ui.cliente === "comau" && dados.config.jantarPadrao, almocoSaida: "", almocoVolta: "",
  // Só a NGC tem períodos: nos outros clientes o cálculo usa início, fim e pausa.
  periodos: modeloNGC() ? periodosVazios() : undefined,
  commessa: dados.config.commessaPadrao || "", centro: centroPadrao(), obs: "", original: null, alterado: false,
});
const formDe = (r) => ({
  data: r.data, inicio: r.inicio || "", fim: r.fim || "", jantar: !!r.jantar, almocoSaida: r.almocoSaida || "", almocoVolta: r.almocoVolta || "",
  periodos: r.periodos ? r.periodos.map((p) => [...p]) : modeloNGC() ? periodosVazios() : undefined,
  commessa: r.commessa ?? (dados.config.commessaPadrao || ""), centro: r.centro || centroPadrao(), obs: r.obs || "", original: r.data, alterado: false,
});
function navegar(aba, { manterForm = false } = {}) {
  ui.aba = aba;
  if (aba === "registrar" && !manterForm) form = formNovo();
  document.querySelectorAll(".abas button").forEach((b) =>
    b.dataset.aba === aba ? b.setAttribute("aria-current", "page") : b.removeAttribute("aria-current"));
  render();
  $("#conteudo").scrollTop = 0;
}

function render() {
  const telas = {
    registrar: ["Registrar", cap(new Date().toLocaleDateString("pt-BR", { weekday: "long", day: "numeric", month: "long" })), telaRegistrar,
      '<button class="link mudo" data-acao="limpar">Limpar</button><button class="primario" data-acao="salvar">Salvar</button>'],
    registros: ["Registros", "Horas do mês", telaRegistros, `<button class="primario" data-acao="gerar">${ROTULO_GERAR}</button>`],
    ajustes: ["Ajustes", "Preferências do app", telaAjustes, ""],
  };
  const [titulo, sub, tela, acoes] = telas[ui.aba];
  $("#cab-titulo").textContent = titulo;
  $("#cab-sub").textContent = sub;
  $("#conteudo").innerHTML = seletorCliente() + tela();
  $("#acoes").innerHTML = acoes;
  if (ui.aba === "registrar") atualizarCalculo();
  if (ui.aba === "registros") $('[data-acao="gerar"]').disabled = !registrosDoMes(ui.ano, ui.mes).length;
  atualizarFaixaBackup();
}

// Só aparece com mais de um modelo: pontos, totais, planilha e comprovante seguem o cliente escolhido.
function seletorCliente() {
  if (!variosClientes()) return "";
  const opcao = (c) => `<button data-acao="cliente" data-valor="${c}" aria-pressed="${ui.cliente === c}">${NOME_CLIENTE[c]}</button>`;
  return `<div class="linha">
      <span class="lbl-forte">Cliente</span>
      <div class="seg" role="group" aria-label="Cliente">${clientesComModelo().map(opcao).join("")}</div>
    </div>`;
}

function trocarCliente(cliente) {
  ui.cliente = cliente;
  dados.config.cliente = cliente;
  salvar();
  form = formNovo();
  render();
}

// ---------- Tela: Registrar ----------
// Campo de hora digitado no teclado numérico do celular (em vez do seletor de rolagem).
// O exemplo "–:–" (um traço de cada lado, sem ligadura da fonte) e o cursor do campo vazio são desenhados pelo app, com os dois pontos no centro exato.
const campoHora = (id, valor, rotulo) => `<span class="hora-caixa"><input type="text" id="${id}" class="hora" value="${valor}"
  inputmode="numeric" pattern="[0-9]*" maxlength="5" placeholder=" " autocomplete="off" enterkeyhint="next" aria-label="${rotulo}"><span class="hora-exemplo" aria-hidden="true"><span>–</span><span class="dp">:<i class="cursor-falso"></i></span><span>–</span></span></span>`;

const linhaHora = (campo, valor, rotulo) => `<div class="linha">
        <span class="lbl cresce">${rotulo}</span>
        ${campoHora(`f-${campo}`, valor, rotulo)}
        <button class="link" data-acao="agora" data-campo="${campo}">Agora</button>
      </div>`;

// NGC: De/Até de manhã, tarde e noite, como na planilha. "Agora" preenche o próximo horário vazio.
function blocoPeriodos(f) {
  if (!f.periodos) f.periodos = periodosVazios();
  const linha = (i) => `<span class="lbl">${PERIODOS_NGC[i]}${i === 2 ? ' <span class="mais1" id="c-mais1" hidden>+1</span>' : ""}</span>
        ${campoHora(`f-p${i}-de`, f.periodos[i][0], `${PERIODOS_NGC[i]}, de`)}
        ${campoHora(`f-p${i}-ate`, f.periodos[i][1], `${PERIODOS_NGC[i]}, até`)}`;
  return `<div class="linha">
        <span class="lbl cresce">Horários</span>
        <button class="link" data-acao="agora-periodo">Agora</button>
      </div>
      <div class="periodos">
        <span></span><span class="dica">De</span><span class="dica">Até</span>
        ${[0, 1, 2].map(linha).join("")}
      </div>`;
}

function telaRegistrar() {
  const f = form;
  const gme = modeloGME();
  const ngc = modeloNGC();
  const pausaPendente = gme && f.original && f.jantar && !f.almocoSaida;
  const comPausa = f.jantar || (f.almocoSaida && f.almocoVolta);
  return `<div class="total">
      <span class="rotulo">Horas trabalhadas</span>
      <span class="grande" id="c-horas">00:00 h</span>
      <span class="rotulo num" id="c-sub">0 min</span>
    </div>
    <div>
      <label class="linha linha-data">
        <span class="lbl">Data</span>
        <span class="valor" id="c-data"></span>
        <input type="date" id="f-data" value="${f.data}" aria-label="Data">
      </label>
      ${ngc ? blocoPeriodos(f) : `${linhaHora("inicio", f.inicio, "Início")}
      ${gme ? `${linhaHora("almoco-saida", f.almocoSaida, "Saída almoço")}${linhaHora("almoco-volta", f.almocoVolta, "Volta almoço")}` : ""}
      <div class="linha">
        <span class="lbl cresce">Fim <span class="mais1" id="c-mais1" hidden>+1</span></span>
        ${campoHora("f-fim", f.fim, "Fim")}
        <button class="link" data-acao="agora" data-campo="fim">Agora</button>
      </div>`}
      ${gme || ngc ? "" : `<div class="linha">
        <span class="lbl">Pausa (min)</span>
        <div class="seg" role="group" aria-label="Pausa">
          <button data-acao="pausa" data-valor="0" aria-pressed="${!comPausa}">0</button>
          <button data-acao="pausa" data-valor="60" aria-pressed="${!!comPausa}">60</button>
        </div>
      </div>`}
      ${gme ? `<div class="linha empilhada">
        <span class="lbl">Commessa</span>
        ${campoCommessa("f-commessa", f.commessa, "Commessa")}
      </div>` : !ngc && modelos.comau ? `<div class="linha empilhada">
        <span class="lbl">Centro de custo</span>
        ${campoCentro("f-centro", f.centro, "Centro de custo")}
      </div>` : ""}
      <div class="linha">
        <input id="f-obs" class="nota" value="${esc(f.obs)}" maxlength="120" placeholder="Observação (opcional, vai para a planilha)" aria-label="Observação">
      </div>
    </div>
    <p class="estado">${f.original ? `Editando o registro de ${curta(f.original)}` : "Novo registro"}${pausaPendente ? ". Este dia tinha pausa de 60 min: informe a saída e a volta do almoço." : ""}</p>
    ${f.original ? '<button class="link perigo" data-acao="excluir">Excluir este dia</button>' : ""}`;
}

function lerFormulario() {
  form.data = $("#f-data").value;
  form.obs = $("#f-obs").value.trim();
  if (modeloNGC()) {
    form.periodos = [0, 1, 2].map((i) => [$(`#f-p${i}-de`).value, $(`#f-p${i}-ate`).value]);
    Object.assign(form, resumoPeriodos(form.periodos), { jantar: false });
    return;
  }
  form.inicio = $("#f-inicio").value;
  form.fim = $("#f-fim").value;
  if (modeloGME()) {
    form.jantar = false;
    form.almocoSaida = $("#f-almoco-saida").value;
    form.almocoVolta = $("#f-almoco-volta").value;
    form.commessa = $("#f-commessa").value.trim();
  } else {
    form.centro = $("#f-centro")?.value ?? form.centro;
  }
}

function atualizarCalculo() {
  lerFormulario();
  const min = minutosTrabalhados(form);
  $("#c-horas").textContent = `${fmtHM(min)} h`;
  $("#c-sub").textContent = `${min} min`;
  $("#c-data").textContent = form.data ? dataDe(form.data).toLocaleDateString("pt-BR", { day: "numeric", month: "short", year: "numeric" }) : "—";
  $("#c-mais1").hidden = !viraDia(form);
}

function validar(v) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v.data)) return "Informe a data.";
  const erroHorarios = modeloNGC() ? validarPeriodos(v.periodos) : validarTurno(v);
  if (erroHorarios) return erroHorarios;
  if (dados.registros[chaveRegistro(v.data, ui.cliente)] && v.data !== v.original) return `Já existe um registro em ${curta(v.data)}. Para alterá-lo, toque nele em Registros.`;
  const conflito = pontoSobreposto(v);
  if (conflito) return `Este horário se sobrepõe ao ponto da ${NOME_CLIENTE[conflito.cliente]} de ${curta(conflito.data)} (${conflito.inicio}–${conflito.fim}).`;
  return "";
}

function validarTurno(v) {
  if (!v.inicio) return "Informe o horário de início.";
  if (!horaOk(v.inicio)) return `Horário de início inválido ("${v.inicio}"). Digite no formato hh:mm, ex.: 14:35.`;
  if (v.fim && !horaOk(v.fim)) return `Horário de fim inválido ("${v.fim}"). Digite no formato hh:mm, ex.: 23:30.`;
  if (v.fim && v.fim === v.inicio) return "Início e fim não podem ser iguais.";
  if (v.fim && v.jantar && minutosTrabalhados({ ...v, jantar: false }) <= 60) return "O turno tem 1h ou menos: use pausa 0 ou corrija os horários.";
  return validarAlmoco(v);
}

// NGC: nenhum campo é obrigatório (o vazio vira 00:00 e não conta). Basta um horário preenchido;
// as batidas precisam seguir em ordem ao longo do turno (pode passar da meia-noite, até 24 h).
function validarPeriodos(periodos) {
  const campo = (i) => `"${i % 2 ? "Até" : "De"}" da ${PERIODOS_NGC[i >> 1].toLowerCase()}`;
  for (let i = 0; i < 3; i++) {
    const [de, ate] = periodos[i];
    if (de && !horaOk(de)) return `Horário ${campo(2 * i)} inválido ("${de}"). Digite no formato hh:mm, ex.: 08:00.`;
    if (ate && !horaOk(ate)) return `Horário ${campo(2 * i + 1)} inválido ("${ate}"). Digite no formato hh:mm, ex.: 12:00.`;
  }
  const b = batidasNGC(periodos);
  if (!b.length) return "Preencha pelo menos um horário.";
  for (let k = 1; k < b.length; k++) {
    // Saída igual à entrada não vale; entre uma saída e a próxima entrada, pode (intervalo zero).
    if (b[k].m < b[k - 1].m || (k % 2 === 1 && b[k].m === b[k - 1].m))
      return `Os horários precisam seguir em ordem: o ${campo(b[k].campo)} (${b[k].h}) está antes ou igual ao ${campo(b[k - 1].campo)} (${b[k - 1].h}).`;
  }
  return "";
}

// Intervalo do turno em minutos absolutos, para comparar pontos de datas diferentes.
function intervaloDoTurno(r) {
  const [a, m, d] = r.data.split("-").map(Number);
  const base = Date.UTC(a, m - 1, d) / 60000;
  let fim = minutosDe(r.fim);
  if (fim < minutosDe(r.inicio)) fim += 1440;
  return [base + minutosDe(r.inicio), base + fim];
}

function pontoSobreposto(v) {
  if (!v.fim) return null;
  const [ini, fim] = intervaloDoTurno(v);
  return Object.values(dados.registros).find((r) => {
    if (r.cliente === ui.cliente || !horaOk(r.inicio) || !horaOk(r.fim)) return false;
    const [a, b] = intervaloDoTurno(r);
    return ini < b && a < fim;
  });
}

function validarAlmoco({ inicio, fim, almocoSaida: saida, almocoVolta: volta }) {
  if (saida && !horaOk(saida)) return `Saída do almoço inválida ("${saida}"). Digite no formato hh:mm, ex.: 12:00.`;
  if (volta && !horaOk(volta)) return `Volta do almoço inválida ("${volta}"). Digite no formato hh:mm, ex.: 13:00.`;
  if (volta && !saida) return "Informe a saída do almoço.";
  if (saida && !volta && fim) return "Informe a volta do almoço.";
  if (!saida || !volta) return "";
  const rel = (h) => (minutosDe(h) - minutosDe(inicio) + 1440) % 1440;
  if (rel(saida) === 0 || rel(volta) <= rel(saida) || (fim && rel(volta) >= rel(fim)))
    return "O almoço precisa ficar dentro do turno: depois do início e antes do fim.";
  return "";
}

function salvarRegistro() {
  lerFormulario();
  const erro = validar(form);
  if (erro) return mostrarErro("Não foi possível salvar", erro);
  if (form.original && form.original !== form.data) delete dados.registros[chaveRegistro(form.original, ui.cliente)];
  const { data, inicio, fim, jantar, almocoSaida, almocoVolta, commessa, centro, obs } = form;
  const extra = modeloNGC() ? { periodos: form.periodos } : modeloGME() ? { commessa } : { centro };
  const registro = { data, cliente: ui.cliente, inicio, fim, jantar, almocoSaida, almocoVolta, obs, ...extra };
  dados.registros[chaveRegistro(data, ui.cliente)] = registro;
  if (!salvar()) return;
  const editando = form.original;
  avisar(fim ? `${curta(data)} salvo · ${fmtHM(minutosTrabalhados(registro))} h` : `Início salvo às ${inicio}`);
  pedirPersistencia();
  // O ponto salvo fica em Registros; a tela volta zerada para o próximo.
  form = formNovo();
  if (editando) navegar("registros"); else render();
}

function excluirRegistro() {
  if (!form.original || !confirm(`Excluir o dia ${curta(form.original)}?`)) return;
  delete dados.registros[chaveRegistro(form.original, ui.cliente)];
  if (salvar()) avisar(`${curta(form.original)} excluído`);
  form = formNovo();
  render();
}

// ---------- Tela: Registros ----------
function telaRegistros() {
  const { ano, mes } = ui;
  const lista = registrosDoMes(ano, mes);
  const completos = lista.filter((r) => r.inicio && r.fim);
  const minutos = completos.reduce((s, r) => s + minutosTrabalhados(r), 0);
  const seta = (d) => `<svg viewBox="0 0 24 24"><path d="${d}"/></svg>`;

  const itens = lista.map((r) => {
    const d = dataDe(r.data);
    const completo = r.inicio && r.fim;
    const pausa = r.periodos ? textoPeriodos(r) : r.almocoSaida && r.almocoVolta ? `almoço ${r.almocoSaida}–${r.almocoVolta}` : r.jantar ? "pausa 60 min" : "sem pausa";
    const opcao = modeloGME() ? (r.commessa ? `commessa ${rotuloCommessa(r.commessa)}` : "") : modeloNGC() ? "" : (r.centro && r.centro !== centroPadrao() ? r.centro.trim() : "");
    const meta = [pausa, opcao, r.obs].filter(Boolean).join(" · ");
    return `<li><button class="reg" data-acao="editar" data-data="${r.data}" aria-label="Editar ${curta(r.data)}">
        <span class="reg-dia"><b>${d.getDate()}</b><small>${SEMANA[d.getDay()]}</small></span>
        <span class="reg-meio">
          <span class="reg-faixa">${r.inicio || "--:--"} – ${r.fim || "--:--"}${viraDia(r) ? ' <span class="mais1">+1</span>' : ""}</span>
          <span class="reg-meta">${completo ? "" : '<span class="sem-fim">sem fim · </span>'}${esc(meta)}</span>
        </span>
        <span class="reg-valor">${completo ? `${fmtHM(minutosTrabalhados(r))} h` : ""}</span>
      </button></li>`;
  }).join("");

  return `<div class="nav-mes">
      <button class="icone-btn" data-acao="mes-anterior" aria-label="Mês anterior">${seta("M15 18l-6-6 6-6")}</button>
      <span>${MESES[mes - 1]} de ${ano}</span>
      <button class="icone-btn" data-acao="mes-seguinte" aria-label="Próximo mês">${seta("M9 18l6-6-6-6")}</button>
    </div>
    <div class="estatisticas">
      <div><span class="rotulo">Dias</span><span class="estat">${completos.length}</span></div>
      <div><span class="rotulo">Horas</span><span class="estat">${fmtHM(minutos)}</span></div>
    </div>
    ${lista.length ? `<div class="linha"><span class="lbl-forte cresce">Comprovante do mês</span><button class="link" data-acao="comprovante">Gerar imagem</button></div>` : ""}
    <div class="linha"><span class="lbl-forte cresce">Conferência por período</span><button class="link" data-acao="periodo">Escolher datas</button></div>
    ${lista.length ? `<ul class="lista">${itens}</ul>`
      : `<div class="vazio"><span>Nenhum registro em ${MESES[mes - 1].toLowerCase()}.</span><button class="link encostado" data-aba="registrar">Registrar um dia</button></div>`}`;
}

function mudarMes(delta) {
  const d = new Date(ui.ano, ui.mes - 1 + delta, 1);
  [ui.ano, ui.mes] = [d.getFullYear(), d.getMonth() + 1];
  render();
}

// ---------- Gerar planilha ----------
async function escolherModelo() {
  const arq = await escolherArquivo(`.xlsm,.xlsx,${MIME_XLSM},${MIME_XLSX}`);
  if (!arq) return;
  let buffer, lido;
  try {
    buffer = await arq.arrayBuffer();
    lido = await lerModelo(window.JSZip, buffer, arq.name);
  } catch (e) { return mostrarErro("Modelo inválido", mensagemDe(e)); }
  const novo = { arquivo: arq.name, importadoEm: new Date().toISOString(), funcionario: lido.funcionario, centros: lido.centros, commessas: lido.commessas,
    ...(lido.tipo === "ngc" ? { clienteFinal: lido.clienteFinal, local: lido.local } : {}), base64: paraBase64(buffer) };
  try { localStorage.setItem(CHAVES_MODELO[lido.tipo], JSON.stringify(novo)); } catch {
    return mostrarErro("Não foi possível salvar o modelo", "O armazenamento do aparelho está cheio ou bloqueado. Libere espaço e tente novamente.");
  }
  modelos[lido.tipo] = novo;
  ui.cliente = lido.tipo;
  dados.config.cliente = lido.tipo;
  // O centro de custo padrão começa no mais usado do modelo e depois pode ser trocado em Ajustes.
  if (lido.centroMaisUsado && !lido.centros.includes(dados.config.centroPadrao)) dados.config.centroPadrao = lido.centroMaisUsado;
  salvar();
  pedirPersistencia();
  avisar(`Modelo ${NOME_CLIENTE[lido.tipo]} configurado`);
  render();
}

async function gerar() {
  const { ano, mes } = ui;
  const info = modelos[ui.cliente];
  if (!info) {
    return abrirDialogo(`<div class="folha">
        <h2>Escolha o modelo da planilha</h2>
        <p class="sub-folha">Selecione no app Arquivos uma planilha de horas de qualquer mês (COMAU, GME ou NGC). Ela fica guardada só neste aparelho e serve de base para as próximas.</p>
        <div class="pilha">
          <button class="primario" data-acao="modelo">Escolher planilha modelo</button>
          <button class="secundario" data-acao="fechar">Agora não</button>
        </div>
      </div>`);
  }
  const botao = $('[data-acao="gerar"]');
  botao.disabled = true;
  botao.textContent = "Gerando…";
  try {
    const tipo = ui.cliente;
    const doMes = registrosDoMes(ano, mes);
    const r = await gerarPlanilha(window.JSZip, deBase64(info.base64), {
      tipo, ano, mes, centroPadrao: centroPadrao(), commessaPadrao: dados.config.commessaPadrao, commessasExtras: dados.config.commessasExtras,
      // Na GME e na NGC, o fim de um turno do mês anterior pode cair no dia 1, por isso vão todos os registros do cliente.
      registros: tipo === "comau" ? doMes : Object.values(dados.registros).filter((x) => x.cliente === tipo),
    });
    const arquivo = new File([r.arquivo], nomeArquivo(ano, mes, info.funcionario, tipo), { type: tipo === "comau" ? MIME_XLSM : MIME_XLSX });
    // O comprovante já fica pronto aqui: o iPhone só compartilha arquivos prontos no toque do botão.
    const comprovante = await criarComprovante(ano, mes);
    abrirDialogo(`<div class="folha">
        <div class="icone-resultado ok"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
        <h2>Planilha gerada</h2>
        <p class="sub-folha">${esc(arquivo.name)}<br>+ comprovante com os ${doMes.length} dias registrados</p>
        <div class="estatisticas" style="margin-bottom:20px">
          <div><span class="rotulo">Dias</span><span class="estat">${r.dias}</span></div>
          <div><span class="rotulo">Horas</span><span class="estat">${fmtHM(Math.round(r.totalHoras * 60))}</span></div>
            </div>
        <div class="pilha">
          <button class="primario" data-acao="compartilhar-ambos">Salvar planilha + comprovante</button>
          <button class="secundario" data-acao="compartilhar">Só a planilha</button>
          <button class="link mudo" style="width:100%;min-height:44px" data-acao="fechar">Fechar</button>
        </div>
      </div>`);
    $('[data-acao="compartilhar-ambos"]').addEventListener("click", () => entregarArquivos([arquivo, comprovante]));
    $('[data-acao="compartilhar"]').addEventListener("click", () => entregarArquivos([arquivo]));
  } catch (e) {
    mostrarErro("Não foi possível gerar a planilha", mensagemDe(e));
  } finally {
    botao.disabled = false;
    botao.textContent = ROTULO_GERAR;
  }
}

// ---------- Gerar por período (conferência) ----------
// Opcional: a geração normal continua sendo a do mês. Gera só os dias escolhidos (até 31),
// de um cliente só; se o período cruza meses, sai uma planilha para cada mês.
const dataBR = (c) => `${c.slice(8, 10)}/${c.slice(5, 7)}/${c.slice(0, 4)}`;
const listaNomes = (nomes) => (nomes.length > 1 ? `${nomes.slice(0, -1).join(", ")} e ${nomes.at(-1)}` : nomes[0]);
function mesesEntre(de, ate) {
  const meses = [];
  for (let [a, m] = [+de.slice(0, 4), +de.slice(5, 7)]; `${a}-${pad(m)}` <= ate.slice(0, 7); m === 12 ? (a++, m = 1) : m++) meses.push({ ano: a, mes: m });
  return meses;
}

function abrirPeriodo() {
  const { ano, mes } = ui;
  const inicioMes = `${ano}-${pad(mes)}-01`;
  const fimMes = `${ano}-${pad(mes)}-${pad(new Date(ano, mes, 0).getDate())}`;
  const hojeChave = chaveData(new Date());
  abrirDialogo(`<div class="folha">
      <h2>Gerar por período</h2>
      <p class="sub-folha">Para conferência: a planilha sai só com os dias escolhidos (até 31) e os outros ficam em branco. A geração do mês continua igual.</p>
      <div class="linha"><span class="lbl cresce">De</span><input type="date" id="p-de" class="campo-data" value="${inicioMes}" aria-label="Data inicial"></div>
      <div class="linha" style="margin-bottom:18px"><span class="lbl cresce">Até</span><input type="date" id="p-ate" class="campo-data" value="${hojeChave >= inicioMes && hojeChave <= fimMes ? hojeChave : fimMes}" aria-label="Data final"></div>
      <div class="pilha">
        <button class="primario" data-acao="gerar-periodo">Gerar planilha do período</button>
        <button class="secundario" data-acao="fechar">Cancelar</button>
      </div>
    </div>`);
}

async function gerarPorPeriodo() {
  const de = $("#p-de").value, ate = $("#p-ate").value;
  if (!de || !ate) return mostrarErro("Período incompleto", "Informe a data inicial e a data final.");
  if (de > ate) return mostrarErro("Período inválido", "A data inicial precisa ser igual ou anterior à data final.");
  const dias = Math.round((dataDe(ate) - dataDe(de)) / 86400000) + 1;
  if (dias > 31) return mostrarErro("Período longo demais", `O período tem ${dias} dias. Escolha no máximo 31 dias, como em uma planilha mensal.`);
  const lista = Object.values(dados.registros).filter((r) => r.data >= de && r.data <= ate).sort((a, b) => a.data.localeCompare(b.data));
  if (!lista.length) return mostrarErro("Nada para gerar", `Não há pontos entre ${dataBR(de)} e ${dataBR(ate)}.`);
  // Cada cliente tem uma planilha diferente: um período com pontos de mais de um cliente não é gerado.
  const clientes = CLIENTES.filter((c) => lista.some((r) => r.cliente === c));
  if (clientes.length > 1)
    return mostrarErro("Mais de um cliente no período", `Entre ${dataBR(de)} e ${dataBR(ate)} há pontos da ${listaNomes(clientes.map((c) => NOME_CLIENTE[c]))}. Como cada planilha é diferente, gere cada cliente separadamente: escolha um período que tenha pontos de um cliente só.`);
  const cliente = clientes[0];
  const info = modelos[cliente];
  if (!info) return mostrarErro("Modelo não escolhido", `Os pontos do período são da ${NOME_CLIENTE[cliente]}, mas o modelo da planilha dela ainda não foi escolhido. Escolha em Ajustes → Modelo da planilha.`);
  if (ui.cliente !== cliente) { ui.cliente = cliente; dados.config.cliente = cliente; salvar({ alterou: false }); }

  const botao = $('[data-acao="gerar-periodo"]');
  botao.disabled = true;
  botao.textContent = "Gerando…";
  try {
    const periodo = { de, ate };
    // Na GME e na NGC, o fim do turno da véspera pode cair no primeiro dia do período.
    const todos = Object.values(dados.registros).filter((x) => x.cliente === cliente);
    const arquivos = [];
    let horas = 0, diasLancados = 0;
    for (const { ano, mes } of mesesEntre(de, ate)) {
      const prefixo = `${ano}-${pad(mes)}-`;
      if (!lista.some((r) => r.data.startsWith(prefixo))) continue;
      const r = await gerarPlanilha(window.JSZip, deBase64(info.base64), {
        tipo: cliente, ano, mes, periodo, centroPadrao: centroPadrao(), commessaPadrao: dados.config.commessaPadrao, commessasExtras: dados.config.commessasExtras,
        registros: cliente === "comau" ? lista : todos,
      });
      horas += r.totalHoras;
      diasLancados += r.dias;
      const fimMes = `${prefixo}${pad(new Date(ano, mes, 0).getDate())}`;
      const ini = de > `${prefixo}01` ? de : `${prefixo}01`, fim = ate < fimMes ? ate : fimMes;
      const nome = nomeArquivo(ano, mes, info.funcionario, cliente).replace(/(\.\w+)$/, `_conferencia_${ini.slice(8)}a${fim.slice(8)}$1`);
      arquivos.push(new File([r.arquivo], nome, { type: cliente === "comau" ? MIME_XLSM : MIME_XLSX }));
    }
    const comprovante = await criarComprovante(null, null, {
      lista, titulo: `${dataBR(de)} a ${dataBR(ate)}`,
      nome: `${de.slice(8)}-${de.slice(5, 7)}_a_${ate.slice(8)}-${ate.slice(5, 7)}_${ate.slice(0, 4)}`,
    });
    render();
    const varias = arquivos.length > 1;
    abrirDialogo(`<div class="folha">
        <div class="icone-resultado ok"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
        <h2>${varias ? "Planilhas do período geradas" : "Planilha do período gerada"}</h2>
        <p class="sub-folha">${NOME_CLIENTE[cliente]} · ${dataBR(de)} a ${dataBR(ate)}<br>${arquivos.map((a) => esc(a.name)).join("<br>")}<br>+ comprovante do período</p>
        <div class="estatisticas" style="margin-bottom:20px">
          <div><span class="rotulo">Dias</span><span class="estat">${diasLancados}</span></div>
          <div><span class="rotulo">Horas</span><span class="estat">${fmtHM(Math.round(horas * 60))}</span></div>
            </div>
        <div class="pilha">
          <button class="primario" data-acao="compartilhar-ambos">Salvar ${varias ? "planilhas" : "planilha"} + comprovante</button>
          <button class="secundario" data-acao="compartilhar">${varias ? "Só as planilhas" : "Só a planilha"}</button>
          <button class="link mudo" style="width:100%;min-height:44px" data-acao="fechar">Fechar</button>
        </div>
      </div>`);
    $('[data-acao="compartilhar-ambos"]').addEventListener("click", () => entregarArquivos([...arquivos, comprovante]));
    $('[data-acao="compartilhar"]').addEventListener("click", () => entregarArquivos(arquivos));
  } catch (e) {
    mostrarErro("Não foi possível gerar a planilha do período", mensagemDe(e));
  } finally {
    if (botao.isConnected) { botao.disabled = false; botao.textContent = "Gerar planilha do período"; }
  }
}

// Abre o "Compartilhar" do iPhone com os arquivos (Salvar em Arquivos, WhatsApp, e-mail...).
// Sem suporte a compartilhar (ex.: computador), baixa os arquivos.
async function entregarArquivos(arquivos) {
  if (navigator.canShare?.({ files: arquivos })) {
    try { await navigator.share({ files: arquivos }); return true; }
    catch (e) { if (e.name === "AbortError") return false; }
  }
  await baixarArquivos(arquivos);
  return true;
}

async function baixarArquivos(arquivos) {
  for (const arquivo of arquivos) {
    const url = URL.createObjectURL(arquivo);
    const a = Object.assign(document.createElement("a"), { href: url, download: arquivo.name });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    await new Promise((r) => setTimeout(r, 400));
  }
}

// ---------- Comprovante do mês ----------
// extra (geração por período): { lista, titulo, nome } no lugar do mês.
async function criarComprovante(ano, mes, extra = {}) {
  const lista = extra.lista || registrosDoMes(ano, mes);
  const tituloBase = extra.titulo || `${MESES[mes - 1]} de ${ano}`;
  const completos = lista.filter((r) => r.inicio && r.fim);
  const agora = new Date();
  const funcionario = modelos[ui.cliente]?.funcionario || "";
  const dois = variosClientes();
  const geradoEm = `${agora.toLocaleDateString("pt-BR")} às ${horaAgora()}`;
  const totais = { dias: completos.length, horas: fmtHM(completos.reduce((s, r) => s + minutosTrabalhados(r), 0)) };
  const imagem = modeloNGC() ? await desenharComprovante({
    // NGC: colunas como na planilha (manhã, tarde, noite e total), com o cabeçalho do cliente.
    layout: "periodos", funcionario, titulo: tituloBase,
    subtitulo: ["NGC", modelos.ngc?.clienteFinal, modelos.ngc?.local].filter(Boolean).join(" · "),
    geradoEm, totais,
    linhas: lista.map((r) => {
      const d = dataDe(r.data);
      return {
        dia: pad(d.getDate()), semana: SEMANA[d.getDay()], mais1: viraDia(r),
        // Como na planilha: os pares entrada/saída em sequência; batida sem par não conta e não aparece.
        periodos: (() => { const p = paresNGC(r.periodos || []).filter((x) => x.ate); return [0, 1, 2].map((i) => (p[i] ? `${p[i].de}–${p[i].ate}` : "—")); })(),
        horas: fmtHM(minutosTrabalhados(r)), obs: r.obs || "", semFim: false,
      };
    }),
  }) : await desenharComprovante({
    funcionario,
    titulo: `${tituloBase}${dois ? ` · ${NOME_CLIENTE[ui.cliente]}` : ""}`,
    coluna: modeloGME() ? "COMMESSA" : "EQUIPE",
    geradoEm, totais,
    linhas: lista.map((r) => {
      const d = dataDe(r.data);
      return {
        dia: pad(d.getDate()), semana: SEMANA[d.getDay()],
        faixa: `${r.inicio || "--:--"} – ${r.fim || "--:--"}`, mais1: viraDia(r),
        pausa: `${pausaDe(r)} min`, horas: fmtHM(minutosTrabalhados(r)), obs: r.obs || "", semFim: !r.fim,
        opcao: modeloGME() ? (r.commessa ? rotuloCommessa(r.commessa) : "—") : (r.centro || centroPadrao() || "—").trim(),
      };
    }),
  });
  const primeiro = funcionario.trim().split(/\s+/)[0];
  const partes = ["Comprovante_Horas", primeiro, (dois || modeloNGC()) && NOME_CLIENTE[ui.cliente], ...(extra.nome ? [extra.nome] : [pad(mes), ano])].filter(Boolean);
  return new File([imagem], `${partes.join("_")}.png`, { type: "image/png" });
}

async function gerarSoComprovante() {
  const { ano, mes } = ui;
  const botao = $('[data-acao="comprovante"]');
  botao.textContent = "Gerando…";
  try {
    const arquivo = await criarComprovante(ano, mes);
    const url = URL.createObjectURL(arquivo);
    abrirDialogo(`<div class="folha">
        <h2>Comprovante de ${MESES[mes - 1].toLowerCase()}</h2>
        <p class="sub-folha">${esc(arquivo.name)}</p>
        <div style="max-height:46dvh;overflow-y:auto;border:1px solid var(--borda);border-radius:12px;margin-bottom:18px">
          <img src="${url}" alt="Prévia do comprovante" style="display:block;width:100%">
        </div>
        <div class="pilha">
          <button class="primario" data-acao="compartilhar">Salvar ou compartilhar</button>
          <button class="secundario" data-acao="fechar">Fechar</button>
        </div>
      </div>`);
    $('[data-acao="compartilhar"]').addEventListener("click", () => entregarArquivos([arquivo]));
    $("#dialogo").addEventListener("close", () => URL.revokeObjectURL(url), { once: true });
  } catch (e) {
    mostrarErro("Não foi possível gerar o comprovante", mensagemDe(e));
  } finally {
    botao.textContent = "Gerar imagem";
  }
}

// Modelo da GME guardado antes de existir a lista de commessas: lê a tabela do arquivo guardado.
async function completarModelos() {
  const m = modelos.gme;
  if (!m || m.commessas) return;
  try {
    m.commessas = (await lerModelo(window.JSZip, deBase64(m.base64), m.arquivo)).commessas;
    localStorage.setItem(CHAVE_MODELO_GME, JSON.stringify(m));
  } catch { /* sem a lista, a commessa continua sendo digitada */ }
}

function escolherArquivo(aceita) {
  return new Promise((resolve) => {
    const input = $("#arquivo");
    input.value = "";
    input.accept = aceita;
    input.onchange = () => resolve(input.files[0] || null);
    input.click();
  });
}

// ---------- Tela: Ajustes ----------
function telaAjustes() {
  const c = dados.config;
  const tema = temaEscolhido();
  const opcaoTema = (v, rotulo) => `<button data-acao="tema" data-valor="${v}" aria-pressed="${tema === v}">${rotulo}</button>`;
  const ultimo = c.ultimoBackup ? new Date(c.ultimoBackup).toLocaleDateString("pt-BR") : "nunca";
  const linhaModelo = (tipo) => {
    const m = modelos[tipo];
    return `<div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Planilha ${NOME_CLIENTE[tipo]}</span>
        <span class="dica reticencias">${m ? esc(m.arquivo) : "Nenhum modelo escolhido"}</span>
      </div>
      <button class="link" data-acao="modelo">${m ? "Trocar" : "Escolher"}</button>
    </div>`;
  };
  return `<h2 class="secao">Aparência</h2>
    <div class="linha">
      <span class="lbl-forte">Tema</span>
      <div class="seg" role="group" aria-label="Tema">${opcaoTema("auto", "Auto")}${opcaoTema("light", "Claro")}${opcaoTema("dark", "Escuro")}</div>
    </div>

    ${modeloNGC() ? "" : `<h2 class="secao">Novos registros${variosClientes() ? ` · ${NOME_CLIENTE[ui.cliente]}` : ""}</h2>`}
    ${modeloGME() ? `<div class="linha empilhada">
      <span class="lbl-forte">Commessa padrão</span>
      ${campoCommessa("a-commessa", c.commessaPadrao, "Commessa padrão")}
    </div>` : modeloNGC() ? "" : `${modelos.comau ? `<div class="linha empilhada">
      <span class="lbl-forte">Centro de custo padrão</span>
      ${campoCentro("a-centro", centroPadrao(), "Centro de custo padrão")}
    </div>` : ""}<div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Pausa de 60 min por padrão</span>
        <span class="dica">Desmarque se costuma não jantar</span>
      </div>
      <button class="interruptor" data-acao="pausa-padrao" role="switch" aria-checked="${c.jantarPadrao}" aria-label="Pausa de 60 minutos por padrão"><span></span></button>
    </div>`}
    ${modeloGME() || (!modeloNGC() && modelos.comau) ? `<div class="linha">
      <span class="cresce"></span>
      <button class="link" data-acao="salvar-ajustes">Salvar</button>
    </div>` : ""}
    ${secaoLista()}

    <h2 class="secao">Modelo da planilha</h2>
    ${linhaModelo("comau")}
    ${linhaModelo("gme")}
    ${linhaModelo("ngc")}
    <p class="dica" style="margin:10px 0 0">Qualquer planilha já feita serve. Arquivo com GME ou NGC no nome vira o modelo desse cliente. Fica só neste aparelho.</p>

    <h2 class="secao">Backup</h2>
    <div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Último backup</span>
        <span class="dica">${ultimo} · inclui registros, ajustes e modelos</span>
      </div>
      <button class="link" data-acao="backup">Fazer agora</button>
    </div>
    <div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Lembrete a cada 24h</span>
        <span class="dica">Avisa quando houver alterações sem backup</span>
      </div>
      <button class="interruptor" data-acao="lembrete-backup" role="switch" aria-checked="${c.lembreteBackup === true}" aria-label="Lembrete de backup a cada 24 horas"><span></span></button>
    </div>
    ${temPastaNoAparelho ? `<div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Pasta do backup</span>
        <span class="dica reticencias">${esc(descricaoPasta())}</span>
      </div>
      <button class="link" data-acao="backup-trocar-pasta">Escolher</button>
    </div>` : ""}
    <div class="linha">
      <span class="lbl-forte cresce">Restaurar de um backup</span>
      <button class="link" data-acao="restaurar">Restaurar</button>
    </div>
    <p class="dica" style="margin:10px 0 0">${dicaBackup()} Ele contém seus dados e modelos: guarde em local privado.</p>
    <p class="estado" style="text-align:center;margin-top:28px">Ponto Automisa · ${Object.keys(dados.registros).length} dias registrados</p>`;
}

// ---------- Centros de custo (COMAU) e commessas (GME) adicionados à mão ----------
// Ficam nos ajustes (e no backup), valem para todos os meses e entram na legenda da planilha gerada.
function secaoLista() {
  const centro = !modeloGME() && !modeloNGC() && modelos.comau;
  if (!centro && !modeloGME()) return "";
  const tipo = centro ? "centro" : "commessa";
  const extras = centro ? dados.config.centrosExtras || [] : dados.config.commessasExtras || [];
  const doModelo = centro ? modelos.comau.centros.length : modelos.gme?.commessas?.length || 0;
  const linhas = extras.map((item) => {
    const valor = centro ? item : item.codigo;
    const rotulo = centro ? item : rotuloCommessa(item.codigo);
    return `<div class="linha"><span class="lbl cresce">${esc(rotulo)}</span>
      <button class="link perigo" data-acao="remover-item" data-tipo="${tipo}" data-valor="${esc(valor)}">Remover</button></div>`;
  }).join("");
  return `<h2 class="secao">${centro ? "Centros de custo" : "Commessas"}</h2>
    ${linhas}
    <div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Adicionar ${centro ? "centro de custo" : "commessa"}</span>
        <span class="dica">${doModelo} do modelo${extras.length ? ` · ${extras.length} adicionad${centro ? "o" : "a"}${extras.length > 1 ? "s" : ""} aqui` : ""}</span>
      </div>
      <button class="link" data-acao="novo-item" data-tipo="${tipo}">Adicionar</button>
    </div>`;
}

// destino: id do campo de seleção que pediu o item novo (ele passa a mostrar o item criado).
function abrirNovoItem(tipo, destino = "") {
  const centro = tipo === "centro";
  abrirDialogo(`<div class="folha">
      <h2>${centro ? "Novo centro de custo" : "Nova commessa"}</h2>
      <p class="sub-folha">Fica salvo neste aparelho e aparece na lista ${centro ? "da COMAU" : "da GME"}. Na planilha gerada, entra também na ${centro ? "legenda de projetos" : "tabela de commessas"}.</p>
      ${centro
        ? `<input id="n-nome" class="campo-texto" maxlength="40" autocomplete="off" placeholder="Ex.: PROJETO 300" aria-label="Nome do centro de custo">`
        : `<input id="n-codigo" class="campo-texto" maxlength="12" inputmode="numeric" autocomplete="off" placeholder="Número da commessa (ex.: 450)" aria-label="Número da commessa">
           <input id="n-linha" class="campo-texto" maxlength="20" autocomplete="off" placeholder="Linha (opcional, ex.: PAN)" aria-label="Linha da commessa">`}
      <div class="pilha" style="margin-top:18px">
        <button class="primario" data-acao="salvar-item" data-tipo="${tipo}" data-destino="${esc(destino)}">Adicionar</button>
        <button class="secundario" data-acao="fechar">Cancelar</button>
      </div>
    </div>`);
}

function salvarItem(tipo, destino) {
  const c = dados.config;
  let novo;
  if (tipo === "centro") {
    novo = $("#n-nome").value.trim().replace(/\s+/g, " ").toUpperCase();
    if (!novo) return mostrarErro("Centro de custo vazio", "Digite o nome do centro de custo.");
    if (listaCentros().some((x) => x.trim().toUpperCase() === novo)) return mostrarErro("Centro de custo repetido", `"${novo}" já está na lista.`);
    c.centrosExtras = [...(c.centrosExtras || []), novo];
  } else {
    novo = $("#n-codigo").value.trim().toUpperCase();
    const linha = $("#n-linha").value.trim().toUpperCase();
    if (!novo) return mostrarErro("Commessa vazia", "Digite o número da commessa.");
    if (listaCommessas().some((x) => x.codigo === novo)) return mostrarErro("Commessa repetida", `A commessa ${novo} já está na lista.`);
    c.commessasExtras = [...(c.commessasExtras || []), { codigo: novo, linha }];
  }
  if (!salvar()) return;
  fecharDialogo();
  avisar(tipo === "centro" ? "Centro de custo adicionado" : "Commessa adicionada");
  // Já deixa o item novo escolhido no campo que pediu (o resto do que foi digitado continua).
  if (destino.startsWith("f-") && ui.aba === "registrar") {
    lerFormulario();
    form[tipo === "centro" ? "centro" : "commessa"] = novo;
    form.alterado = true;
    render();
  } else {
    render();
    const campo = destino && $(`#${destino}`);
    if (campo) { campo.value = novo; campo.dataset.atual = novo; }
  }
}

function removerItem(tipo, valor) {
  const c = dados.config;
  const rotulo = tipo === "centro" ? `o centro de custo "${valor}"` : `a commessa ${rotuloCommessa(valor)}`;
  if (!confirm(`Remover ${rotulo} da lista? Os pontos que já usam continuam com ele.`)) return;
  if (tipo === "centro") c.centrosExtras = (c.centrosExtras || []).filter((x) => x !== valor);
  else c.commessasExtras = (c.commessasExtras || []).filter((x) => x.codigo !== valor);
  if (salvar()) { avisar("Removido da lista"); render(); }
}

function salvarAjustes() {
  const commessa = $("#a-commessa");
  if (commessa) dados.config.commessaPadrao = commessa.value.trim();
  const centro = $("#a-centro");
  if (centro) dados.config.centroPadrao = centro.value;
  return salvar();
}

async function fazerBackup({ trocarPasta = false } = {}) {
  fecharDialogo();
  // O backup vai para onde você escolher; nada é enviado à internet pelo app.
  // Nome fixo: na mesma pasta, o arquivo novo substitui o anterior.
  const conteudo = JSON.stringify({ app: "ponto-comau", versao: 1, exportadoEm: new Date().toISOString(), ...dados, modelos }, null, 2);
  const arquivo = new File([conteudo], ARQUIVO_BACKUP, { type: "application/json" });
  let ok = true;
  let falhaNaPasta = false;
  if (trocarPasta || usaPasta()) {
    try { ok = await gravarNaPasta(arquivo, trocarPasta); }
    catch (e) {
      if (e.name === "AbortError") return;
      // Sem gravar na pasta, o backup não se perde: baixa o arquivo e deixa de insistir nessa pasta.
      falhaNaPasta = true;
      dados.config.pastaFalhou = true;
      dados.config.pastaBackup = null;
      await guardarPasta(null);
      await baixarArquivos([arquivo]);
    }
  } else if (ehAndroid) {
    await baixarArquivos([arquivo]);
  } else {
    ok = await entregarArquivos([arquivo]);
  }
  if (!ok) return;
  dados.config.ultimoBackup = new Date().toISOString();
  salvar({ alterou: false });
  render();
  if (falhaNaPasta) mostrarErro("Backup salvo em Downloads", "Não foi possível gravar na pasta escolhida. Para tentar outra pasta, vá em Ajustes, Pasta do backup.");
  else avisar("Backup criado");
}

// ---------- Pasta do backup (Android) ----------
const usaPasta = () => temPastaNoAparelho && !dados.config.pastaFalhou;
const lugarDoBackup = () => (ehAndroid && !usaPasta() ? "em Downloads" : `na pasta ${PASTA_BACKUP}`);

function descricaoPasta() {
  const nome = dados.config.pastaBackup;
  if (!nome) return dados.config.pastaFalhou ? "Downloads" : "Nenhuma escolhida";
  return nome === PASTA_BACKUP ? PASTA_BACKUP : `${PASTA_BACKUP}, dentro de ${nome}`;
}

function dicaBackup() {
  if (usaPasta()) return `O arquivo novo substitui o anterior na pasta ${PASTA_BACKUP}.`;
  if (ehAndroid) return "O arquivo vai para Downloads. O aparelho pode criar cópias numeradas: ao restaurar, use a mais recente.";
  return `O arquivo tem sempre o mesmo nome. Salve na pasta ${PASTA_BACKUP} e confirme a substituição do anterior.`;
}

// O acesso à pasta fica guardado no próprio navegador, e não vai no arquivo de backup.
function abrirBancoPasta() {
  return new Promise((ok, erro) => {
    const pedido = indexedDB.open("ponto-comau.backup", 1);
    pedido.onupgradeneeded = () => pedido.result.createObjectStore("pasta");
    pedido.onsuccess = () => ok(pedido.result);
    pedido.onerror = () => erro(pedido.error);
  });
}
async function lerPastaGuardada() {
  try {
    const banco = await abrirBancoPasta();
    return await new Promise((ok, erro) => {
      const pedido = banco.transaction("pasta").objectStore("pasta").get("backup");
      pedido.onsuccess = () => ok(pedido.result || null);
      pedido.onerror = () => erro(pedido.error);
    });
  } catch { return null; }
}
async function guardarPasta(pasta) {
  try {
    const banco = await abrirBancoPasta();
    await new Promise((ok, erro) => {
      const tx = banco.transaction("pasta", "readwrite");
      if (pasta) tx.objectStore("pasta").put(pasta, "backup"); else tx.objectStore("pasta").delete("backup");
      tx.oncomplete = ok;
      tx.onerror = () => erro(tx.error);
    });
  } catch { /* sem guardar, o app pergunta a pasta de novo no próximo backup */ }
}

// Retorna false se o usuário negar o acesso. Erros seguem para quem chamou.
async function gravarNaPasta(arquivo, trocarPasta) {
  const modo = { mode: "readwrite" };
  let pasta = trocarPasta ? null : await lerPastaGuardada();
  if (pasta) {
    if ((await pasta.queryPermission(modo)) !== "granted" && (await pasta.requestPermission(modo)) !== "granted") return false;
  } else {
    pasta = await window.showDirectoryPicker(modo);
  }
  const destino = pasta.name === PASTA_BACKUP ? pasta : await pasta.getDirectoryHandle(PASTA_BACKUP, { create: true });
  const item = await destino.getFileHandle(ARQUIVO_BACKUP, { create: true });
  const escrita = await item.createWritable();
  await escrita.write(arquivo);
  await escrita.close();
  // Confere o que foi gravado, porque este arquivo substitui o único backup anterior.
  if ((await item.getFile()).size !== arquivo.size) throw new Error("Arquivo gravado incompleto");
  await guardarPasta(pasta);
  dados.config.pastaBackup = pasta.name;
  dados.config.pastaFalhou = false;
  return true;
}

// ---------- Lembrete de backup ----------
// O navegador não grava sozinho nos arquivos do aparelho: o app avisa e o usuário confirma com um toque.
let backupAdiadoAte = 0;

function estadoBackup() {
  const c = dados.config;
  if (!Object.keys(dados.registros).length) return null;
  if (c.lembreteBackup === null) return "convite";
  if (!c.lembreteBackup || Date.now() < backupAdiadoAte) return null;
  const ultimo = Date.parse(c.ultimoBackup) || 0;
  if (!ultimo) return "pendente";
  return Date.now() - ultimo >= INTERVALO_BACKUP && (Date.parse(c.alteradoEm) || 0) > ultimo ? "pendente" : null;
}

function atualizarFaixaBackup() {
  const faixa = $("#faixa-backup");
  const estado = estadoBackup();
  faixa.hidden = !estado;
  if (!estado) { faixa.innerHTML = ""; return; }
  const ultimo = dados.config.ultimoBackup;
  const [titulo, texto, recusar, aceitar] = estado === "convite"
    ? ["Ativar backup automático?", `A cada 24h o app lembra de salvar tudo ${lugarDoBackup()}.`, ["backup-recusar", "Agora não"], ["backup-ativar", "Ativar"]]
    : ["Há alterações sem backup", ultimo ? `Último backup em ${new Date(ultimo).toLocaleDateString("pt-BR")}.` : "Nenhum backup feito ainda.", ["backup-adiar", "Depois"], ["backup-lembrete", "Fazer backup"]];
  faixa.innerHTML = `<div class="pilha-v">
      <span class="lbl-forte">${titulo}</span>
      <span class="dica">${texto}</span>
    </div>
    <div class="faixa-botoes">
      <button class="link mudo" data-acao="${recusar[0]}">${recusar[1]}</button>
      <button class="link" data-acao="${aceitar[0]}">${aceitar[1]}</button>
    </div>`;
}

// O primeiro backup explica como criar a pasta, e oferece restaurar antes de substituir um backup antigo.
function guiaBackup() {
  const passos = usaPasta()
    ? `<li>Toque em <b>Fazer o backup</b> e crie ou escolha uma pasta do aparelho.</li>
        <li>Permita o acesso à pasta. Se ela não se chamar <b>${PASTA_BACKUP}</b>, o app cria essa pasta dentro dela.</li>
        <li>Nas próximas vezes, um toque substitui o arquivo.</li>`
    : ehAndroid
      ? `<li>Toque em <b>Fazer o backup</b>. O arquivo vai para <b>Downloads</b>.</li>
        <li>O aparelho pode criar cópias numeradas. Ao restaurar, use a mais recente.</li>`
      : `<li>Toque em <b>Fazer o backup</b> e depois em <b>Salvar em Arquivos</b>.</li>
        <li>Escolha <b>No Meu iPhone</b> e crie a pasta <b>${PASTA_BACKUP}</b>.</li>
        <li>Toque em <b>Salvar</b>. Nas próximas vezes, use a mesma pasta e confirme a substituição.</li>`;
  const aviso = usaPasta() || !ehAndroid ? "Cada backup substitui o anterior. Aparelho novo ou sem dados? Restaure antes." : "Aparelho novo ou sem dados? Restaure antes de fazer um novo backup.";
  abrirDialogo(`<div class="folha">
      <h2>Primeiro backup</h2>
      <ol class="passos">${passos}</ol>
      <p class="sub-folha">${aviso}</p>
      <div class="pilha">
        <button class="primario" data-acao="backup-guia">Fazer o backup</button>
        <button class="secundario" data-acao="restaurar">Restaurar um backup</button>
        <button class="link mudo" style="width:100%;min-height:44px" data-acao="fechar">Agora não</button>
      </div>
    </div>`);
}

const precisaGuia = () => !dados.config.ultimoBackup || (usaPasta() && !dados.config.pastaBackup);

function definirLembreteBackup(ativo) {
  dados.config.lembreteBackup = ativo;
  salvar({ alterou: false });
  render();
  if (ativo && precisaGuia()) guiaBackup();
}

async function restaurarBackup() {
  const arq = await escolherArquivo(".json,application/json");
  if (!arq) return;
  let lido;
  try {
    lido = JSON.parse(await arq.text());
    if (lido.app !== "ponto-comau" || typeof lido.registros !== "object") throw new Error();
  } catch { return mostrarErro("Backup inválido", "O arquivo escolhido não é um backup do Ponto Automisa."); }
  const qtd = Object.keys(lido.registros).length;
  if (!confirm(`Restaurar ${qtd} dias do backup de ${new Date(lido.exportadoEm).toLocaleDateString("pt-BR")}? Os registros atuais deste aparelho serão substituídos.`)) return;
  // O conteúdo restaurado é o do backup: nada novo para salvar até a próxima alteração.
  const dataBackup = lido.exportadoEm || null;
  dados = { registros: normalizarRegistros(lido.registros), config: { ...semPreco({ ...PADRAO.config, ...lido.config }), ultimoBackup: dataBackup, alteradoEm: dataBackup, pastaBackup: null, pastaFalhou: false } };
  // Backups antigos trazem um único modelo, no campo "modelo".
  const doBackup = lido.modelos || (lido.modelo ? { [ehNGC(lido.modelo.arquivo) ? "ngc" : ehGME(lido.modelo.arquivo) ? "gme" : "comau"]: lido.modelo } : {});
  for (const [tipo, m] of Object.entries(doBackup)) {
    if (!m?.base64) continue;
    if (!CHAVES_MODELO[tipo]) continue;
    try { localStorage.setItem(CHAVES_MODELO[tipo], JSON.stringify(m)); modelos[tipo] = m; } catch { /* registros restaurados mesmo assim */ }
  }
  ui.cliente = clienteInicial();
  if (salvar({ alterou: false })) avisar(`${qtd} dias restaurados`);
  await completarModelos();
  form = formNovo();
  render();
}

// ---------- Diálogos e avisos ----------
function abrirDialogo(html) {
  const d = $("#dialogo");
  d.innerHTML = html;
  // Foca o título, e não o primeiro campo, para o iPhone não abrir teclado/calendário sozinho.
  // O autofocus vale já na abertura: sem ele, o showModal foca o primeiro campo antes de nós.
  const titulo = d.querySelector("h2");
  if (titulo) { titulo.tabIndex = -1; titulo.autofocus = true; }
  if (!d.open) d.showModal();
  if (titulo) titulo.focus({ preventScroll: true });
}
function fecharDialogo() { $("#dialogo").close(); }

function mostrarErro(titulo, mensagem) {
  abrirDialogo(`<div class="folha">
      <div class="icone-resultado falha"><svg viewBox="0 0 24 24"><path d="M12 8v5M12 16.5h.01"/></svg></div>
      <h2>${esc(titulo)}</h2>
      <p class="sub-folha" style="color:var(--texto)">${esc(mensagem)}</p>
      <button class="primario" data-acao="fechar">Entendi</button>
    </div>`);
}

const mensagemDe = (e) => e instanceof ErroPlanilha ? e.message : `Erro inesperado: ${e.message || e}`;

let avisoTimer;
function avisar(texto) {
  const el = $("#aviso");
  el.textContent = texto;
  el.classList.add("visivel");
  clearTimeout(avisoTimer);
  avisoTimer = setTimeout(() => el.classList.remove("visivel"), 2000);
}

// ---------- Eventos ----------
document.addEventListener("click", (e) => {
  const alvo = e.target.closest("[data-acao], [data-aba]");
  if (!alvo) return;
  if (alvo.dataset.aba) return navegar(alvo.dataset.aba);
  const acoes = {
    agora: () => { $(`#f-${alvo.dataset.campo}`).value = horaAgora(); form.alterado = true; atualizarCalculo(); },
    "agora-periodo": () => {
      const vazio = [...document.querySelectorAll(".periodos .hora")].find((el) => !el.value);
      if (!vazio) return avisar("Todos os horários já estão preenchidos");
      vazio.value = horaAgora();
      form.alterado = true;
      atualizarCalculo();
    },
    pausa: () => {
      form.jantar = alvo.dataset.valor === "60";
      form.almocoSaida = form.almocoVolta = "";
      form.alterado = true;
      alvo.parentElement.querySelectorAll("button").forEach((b) => b.setAttribute("aria-pressed", String(b === alvo)));
      atualizarCalculo();
    },
    salvar: salvarRegistro, excluir: excluirRegistro, limpar: () => { form = formNovo(); render(); },
    editar: () => { form = formDe(dados.registros[chaveRegistro(alvo.dataset.data, ui.cliente)]); navegar("registrar", { manterForm: true }); },
    cliente: () => trocarCliente(alvo.dataset.valor),
    "mes-anterior": () => mudarMes(-1), "mes-seguinte": () => mudarMes(1), gerar, comprovante: gerarSoComprovante,
    periodo: abrirPeriodo, "gerar-periodo": gerarPorPeriodo,
    "alternar-tema": () => escolherTema(temaEfetivo() === "dark" ? "light" : "dark"),
    tema: () => escolherTema(alvo.dataset.valor),
    "pausa-padrao": () => { dados.config.jantarPadrao = !dados.config.jantarPadrao; if (salvar()) render(); },
    "salvar-ajustes": () => { if (salvarAjustes()) { avisar("Ajustes salvos"); render(); } },
    "novo-item": () => abrirNovoItem(alvo.dataset.tipo),
    "salvar-item": () => salvarItem(alvo.dataset.tipo, alvo.dataset.destino || ""),
    "remover-item": () => removerItem(alvo.dataset.tipo, alvo.dataset.valor),
    modelo: () => { fecharDialogo(); escolherModelo(); }, backup: () => { if (usaPasta() && !dados.config.pastaBackup) guiaBackup(); else fazerBackup(); }, "backup-guia": () => fazerBackup(), "backup-trocar-pasta": () => fazerBackup({ trocarPasta: true }), restaurar: () => { fecharDialogo(); restaurarBackup(); }, fechar: fecharDialogo,
    "backup-ativar": () => definirLembreteBackup(true),
    "backup-recusar": () => definirLembreteBackup(false),
    "backup-adiar": () => { backupAdiadoAte = Date.now() + ADIAMENTO_BACKUP; atualizarFaixaBackup(); },
    "backup-lembrete": () => { if (precisaGuia()) guiaBackup(); else fazerBackup(); },
    "lembrete-backup": () => definirLembreteBackup(!dados.config.lembreteBackup),
  };
  if (acoes[alvo.dataset.acao]) { e.preventDefault(); acoes[alvo.dataset.acao](); }
});

document.addEventListener("input", (e) => {
  if (ui.aba !== "registrar" || !e.target.id?.startsWith("f-") || e.target.id === "f-data") return;
  form.alterado = true;
  if (e.target.classList.contains("hora")) {
    mascaraHora(e.target);
    // Horário completo: pula para o próximo campo de hora vazio, na ordem da tela.
    const campos = [...document.querySelectorAll(".hora")];
    const proximo = campos[campos.indexOf(e.target) + 1];
    if (horaOk(e.target.value) && proximo && !proximo.value) proximo.focus();
  }
  atualizarCalculo();
});
document.addEventListener("change", (e) => {
  // "+ Adicionar…" na lista de centro de custo ou commessa: volta à opção anterior e abre o cadastro.
  if (e.target.tagName === "SELECT" && e.target.value === NOVO_ITEM) {
    e.target.value = e.target.dataset.atual || "";
    return abrirNovoItem(/centro/.test(e.target.id) ? "centro" : "commessa", e.target.id);
  }
  if (e.target.tagName === "SELECT" && "atual" in e.target.dataset) e.target.dataset.atual = e.target.value;
  if (ui.aba !== "registrar") return;
  if (e.target.classList?.contains("hora")) completarHora(e.target);
  if (e.target.id?.startsWith("f-")) { form.alterado = true; atualizarCalculo(); }
});

$("#dialogo").addEventListener("click", (e) => { if (e.target.id === "dialogo") fecharDialogo(); });
// Ao voltar para o app (ex.: no dia seguinte), recarrega o formulário se nada foi digitado.
document.addEventListener("visibilitychange", () => {
  if (document.hidden) return;
  if (aplicarAtualizacao()) return;
  if (ui.aba === "registrar" && !form.alterado) navegar("registrar");
  atualizarFaixaBackup();
});

// ---------- Atualização do app ----------
// O iPhone costuma só retomar o app aberto. Ao voltar para ele, procura versão nova;
// quando ela fica pronta, recarrega, mas nunca no meio de um ponto sendo digitado.
let atualizacaoPronta = false;
function aplicarAtualizacao() {
  if (!atualizacaoPronta || form.alterado || $("#dialogo").open || document.hidden) return false;
  atualizacaoPronta = false;
  location.reload();
  return true;
}

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  // Na primeira instalação a página já é a versão atual: só as trocas seguintes pedem recarregar.
  let ignorarProxima = !navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (ignorarProxima) { ignorarProxima = false; return; }
    atualizacaoPronta = true;
    aplicarAtualizacao();
  });
  navigator.serviceWorker.register("sw.js").then((registro) => {
    document.addEventListener("visibilitychange", () => { if (!document.hidden) registro?.update?.().catch(() => {}); });
  }).catch(() => { /* app funciona sem offline */ });
}

aplicarTema();
navegar("registrar");
completarModelos().then(() => { if (!form.alterado && ui.aba === "registrar") navegar("registrar"); });
