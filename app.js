import { ErroPlanilha, ehGME, gerarPlanilha, lerModelo, minutosDe, nomeArquivo, pausaMinutos } from "./xlsm.js";
import { desenharComprovante } from "./comprovante.js";

const CHAVE = "ponto-comau.v1";
const CHAVE_MODELO = "ponto-comau.modelo";
const CHAVE_MODELO_GME = "ponto-comau.modelo-gme";
const NOME_CLIENTE = { comau: "COMAU", gme: "GME" };
const CHAVE_TEMA = "ponto-comau.tema";
const MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
const SEMANA = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
const MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12";
const MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ROTULO_GERAR = "Gerar planilha de horas";

const $ = (sel, raiz = document) => raiz.querySelector(sel);
const pad = (n) => String(n).padStart(2, "0");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const reais = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
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
function minutosTrabalhados(r) {
  if (!horaOk(r.inicio) || !horaOk(r.fim)) return 0;
  let min = minutosDe(r.fim) - minutosDe(r.inicio);
  if (min < 0) min += 1440;
  return Math.max(0, min - pausaDe(r));
}
// Valor de cada dia arredondado a centavos, como no informe; o mês soma os dias.
const valorDoDia = (r) => Math.round((minutosTrabalhados(r) / 60) * (r.taxa || 0) * 100) / 100;
const fmtHM = (min) => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
const viraDia = (r) => horaOk(r.inicio) && horaOk(r.fim) && minutosDe(r.fim) < minutosDe(r.inicio);

function lerTaxa(texto) {
  // Com vírgula, o ponto é separador de milhar (1.234,50); sem vírgula, é decimal (47.50).
  const t = String(texto ?? "").trim();
  const v = Number(t.includes(",") ? t.replace(/\./g, "").replace(",", ".") : t);
  return Number.isFinite(v) && v > 0 ? v : null;
}
// Preço 0 fica com o campo vazio (mostra "0" de exemplo), para digitar sem precisar apagar.
const taxaTexto = (v) => (!v ? "" : String(v).replace(".", ","));

// ---------- Backup ----------
const ARQUIVO_BACKUP = "backup-pontos.json";
const PASTA_BACKUP = "Backup Pontos";
const INTERVALO_BACKUP = 24 * 60 * 60 * 1000;
const ADIAMENTO_BACKUP = 6 * 60 * 60 * 1000; // "Depois": o lembrete volta em 6 horas
// No Android, o navegador pode gravar direto numa pasta; se não puder, o arquivo vai para Downloads.
const ehAndroid = /Android/i.test(navigator.userAgent);
const temPastaNoAparelho = ehAndroid && !!window.showDirectoryPicker;

// ---------- Armazenamento (só neste aparelho) ----------
const PADRAO = { registros: {}, config: { taxaPadrao: 0, centroPadrao: null, jantarPadrao: true, commessaPadrao: "", ultimoBackup: null, lembreteBackup: null, alteradoEm: null, pastaBackup: null, pastaFalhou: false } };
let dados = carregar();
let modelos = carregarModelos(); // { comau, gme }: cada um { arquivo, importadoEm, funcionario, centros, base64 } ou null

// Cada cliente tem seus pontos: o mesmo dia pode ter um ponto da COMAU e outro da GME.
function chaveRegistro(data, cliente) { return cliente === "gme" ? `${data}#gme` : data; }

// Pontos antigos não têm cliente: almoço ou commessa indicam GME; os demais são da COMAU.
function normalizarRegistros(registros) {
  const saida = {};
  for (const r of Object.values(registros)) {
    const cliente = r.cliente || (r.almocoSaida || r.commessa ? "gme" : "comau");
    saida[chaveRegistro(r.data, cliente)] = { ...r, cliente };
  }
  return saida;
}

function carregar() {
  try {
    const salvo = JSON.parse(localStorage.getItem(CHAVE));
    if (salvo && salvo.registros) return { registros: normalizarRegistros(salvo.registros), config: { ...PADRAO.config, ...salvo.config } };
  } catch { /* sem dados ou armazenamento bloqueado */ }
  return structuredClone(PADRAO);
}

function carregarModelos() {
  const ler = (chave) => { try { return JSON.parse(localStorage.getItem(chave)); } catch { return null; } };
  const m = { comau: ler(CHAVE_MODELO), gme: ler(CHAVE_MODELO_GME) };
  // Antes havia um modelo só, guardado sempre na mesma chave, que podia ser o da GME.
  if (ehGME(m.comau?.arquivo)) {
    m.gme = m.gme || m.comau;
    m.comau = null;
    try { localStorage.setItem(CHAVE_MODELO_GME, JSON.stringify(m.gme)); localStorage.removeItem(CHAVE_MODELO); } catch { /* tenta de novo na próxima abertura */ }
  }
  return m;
}

const doisClientes = () => !!(modelos.comau && modelos.gme);
function clienteInicial() {
  const salvo = dados.config.cliente;
  if (salvo && modelos[salvo]) return salvo;
  return modelos.gme && !modelos.comau ? "gme" : "comau";
}

// No cliente GME, o ponto pede almoço e commessa.
const modeloGME = () => ui.cliente === "gme";
const taxaPadrao = () => (modeloGME() ? dados.config.taxaPadraoGME ?? dados.config.taxaPadrao : dados.config.taxaPadrao);

// Opções do ponto, lidas da legenda de cada modelo.
const centroPadrao = () => dados.config.centroPadrao || modelos.comau?.centros?.[0] || "";
const listaCommessas = () => modelos.gme?.commessas || [];
function rotuloCommessa(codigo) {
  const item = listaCommessas().find((c) => c.codigo === String(codigo));
  return item?.linha ? `${item.codigo} – ${item.linha}` : String(codigo);
}
const opcaoHtml = (valor, rotulo, atual) => `<option value="${esc(valor)}" ${valor === atual ? "selected" : ""}>${esc(rotulo)}</option>`;
function campoCentro(id, atual, rotulo) {
  const lista = [...new Set([...(modelos.comau?.centros || []), atual].filter(Boolean))];
  return `<select id="${id}" class="opcao" aria-label="${rotulo}">${lista.map((c) => opcaoHtml(c, c.trim(), atual)).join("")}</select>`;
}
function campoCommessa(id, atual, rotulo) {
  const lista = listaCommessas();
  // Sem tabela no modelo, a commessa é digitada.
  if (!lista.length) return `<input id="${id}" class="numero" inputmode="numeric" maxlength="12" value="${esc(atual)}" placeholder="—" aria-label="${rotulo}">`;
  const valor = String(atual ?? "");
  const conhecida = !valor || lista.some((c) => c.codigo === valor);
  return `<select id="${id}" class="opcao" aria-label="${rotulo}">
      ${opcaoHtml("", "Sem commessa", valor)}
      ${lista.map((c) => opcaoHtml(c.codigo, rotuloCommessa(c.codigo), valor)).join("")}
      ${conhecida ? "" : opcaoHtml(valor, valor, valor)}
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
  data, inicio: "", fim: "", jantar: !modeloGME() && dados.config.jantarPadrao, almocoSaida: "", almocoVolta: "",
  taxa: taxaPadrao(), commessa: dados.config.commessaPadrao || "", centro: centroPadrao(), obs: "", original: null, alterado: false,
});
const formDe = (r) => ({
  data: r.data, inicio: r.inicio || "", fim: r.fim || "", jantar: !!r.jantar, almocoSaida: r.almocoSaida || "", almocoVolta: r.almocoVolta || "",
  taxa: r.taxa ?? taxaPadrao(), commessa: r.commessa ?? (dados.config.commessaPadrao || ""), centro: r.centro || centroPadrao(), obs: r.obs || "", original: r.data, alterado: false,
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

// Só aparece com os dois modelos: pontos, totais, planilha e comprovante seguem o cliente escolhido.
function seletorCliente() {
  if (!doisClientes()) return "";
  const opcao = (c) => `<button data-acao="cliente" data-valor="${c}" aria-pressed="${ui.cliente === c}">${NOME_CLIENTE[c]}</button>`;
  return `<div class="linha">
      <span class="lbl-forte">Cliente</span>
      <div class="seg" role="group" aria-label="Cliente">${opcao("comau")}${opcao("gme")}</div>
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
// O exemplo "--:--" e o cursor do campo vazio são desenhados pelo app, com os dois pontos no centro exato.
const campoHora = (id, valor, rotulo) => `<span class="hora-caixa"><input type="text" id="${id}" class="hora" value="${valor}"
  inputmode="numeric" pattern="[0-9]*" maxlength="5" placeholder=" " autocomplete="off" enterkeyhint="next" aria-label="${rotulo}"><span class="hora-exemplo" aria-hidden="true"><span>--</span><span class="dp">:<i class="cursor-falso"></i></span><span>--</span></span></span>`;

const linhaHora = (campo, valor, rotulo) => `<div class="linha">
        <span class="lbl cresce">${rotulo}</span>
        ${campoHora(`f-${campo}`, valor, rotulo)}
        <button class="link" data-acao="agora" data-campo="${campo}">Agora</button>
      </div>`;

function telaRegistrar() {
  const f = form;
  const gme = modeloGME();
  const pausaPendente = gme && f.original && f.jantar && !f.almocoSaida;
  const comPausa = f.jantar || (f.almocoSaida && f.almocoVolta);
  return `<div class="total">
      <span class="rotulo">Valor total</span>
      <span class="grande" id="c-valor">R$ 0,00</span>
      <span class="rotulo num" id="c-sub">00:00 h · 0 min</span>
    </div>
    <div>
      <label class="linha linha-data">
        <span class="lbl">Data</span>
        <span class="valor" id="c-data"></span>
        <input type="date" id="f-data" value="${f.data}" aria-label="Data">
      </label>
      ${linhaHora("inicio", f.inicio, "Início")}
      ${gme ? `${linhaHora("almoco-saida", f.almocoSaida, "Saída almoço")}${linhaHora("almoco-volta", f.almocoVolta, "Volta almoço")}` : ""}
      <div class="linha">
        <span class="lbl cresce">Fim <span class="mais1" id="c-mais1" hidden>+1</span></span>
        ${campoHora("f-fim", f.fim, "Fim")}
        <button class="link" data-acao="agora" data-campo="fim">Agora</button>
      </div>
      ${gme ? "" : `<div class="linha">
        <span class="lbl">Pausa (min)</span>
        <div class="seg" role="group" aria-label="Pausa">
          <button data-acao="pausa" data-valor="0" aria-pressed="${!comPausa}">0</button>
          <button data-acao="pausa" data-valor="60" aria-pressed="${!!comPausa}">60</button>
        </div>
      </div>`}
      <div class="linha">
        <span class="lbl cresce">Preço por hora</span>
        <span class="unidade">R$</span>
        <input id="f-taxa" class="numero" inputmode="decimal" value="${taxaTexto(f.taxa)}" placeholder="0" aria-label="Preço por hora">
      </div>
      ${gme ? `<div class="linha empilhada">
        <span class="lbl">Commessa</span>
        ${campoCommessa("f-commessa", f.commessa, "Commessa")}
      </div>` : modelos.comau ? `<div class="linha empilhada">
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
  form.inicio = $("#f-inicio").value;
  form.fim = $("#f-fim").value;
  form.obs = $("#f-obs").value.trim();
  form.taxa = lerTaxa($("#f-taxa").value);
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
  $("#c-valor").textContent = reais.format(valorDoDia(form));
  $("#c-sub").textContent = `${fmtHM(min)} h · ${min} min`;
  $("#c-data").textContent = form.data ? dataDe(form.data).toLocaleDateString("pt-BR", { day: "numeric", month: "short", year: "numeric" }) : "—";
  $("#c-mais1").hidden = !viraDia(form);
}

function validar(v) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v.data)) return "Informe a data.";
  if (!v.inicio) return "Informe o horário de início.";
  if (!horaOk(v.inicio)) return `Horário de início inválido ("${v.inicio}"). Digite no formato hh:mm, ex.: 14:35.`;
  if (v.fim && !horaOk(v.fim)) return `Horário de fim inválido ("${v.fim}"). Digite no formato hh:mm, ex.: 23:30.`;
  if (v.fim && v.fim === v.inicio) return "Início e fim não podem ser iguais.";
  if (v.fim && v.jantar && minutosTrabalhados({ ...v, jantar: false }) <= 60) return "O turno tem 1h ou menos: use pausa 0 ou corrija os horários.";
  const erroAlmoco = validarAlmoco(v);
  if (erroAlmoco) return erroAlmoco;
  if (!v.taxa) return "Informe o preço por hora (ex.: 50 ou 47,50). Para já vir preenchido, defina o padrão em Ajustes.";
  if (dados.registros[chaveRegistro(v.data, ui.cliente)] && v.data !== v.original) return `Já existe um registro em ${curta(v.data)}. Para alterá-lo, toque nele em Registros.`;
  const conflito = pontoSobreposto(v);
  if (conflito) return `Este horário se sobrepõe ao ponto da ${NOME_CLIENTE[conflito.cliente]} de ${curta(conflito.data)} (${conflito.inicio}–${conflito.fim}).`;
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
  const { data, inicio, fim, jantar, almocoSaida, almocoVolta, taxa, commessa, centro, obs } = form;
  const registro = { data, cliente: ui.cliente, inicio, fim, jantar, almocoSaida, almocoVolta, taxa, obs, ...(modeloGME() ? { commessa } : { centro }) };
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
  const valor = completos.reduce((s, r) => s + valorDoDia(r), 0);
  const seta = (d) => `<svg viewBox="0 0 24 24"><path d="${d}"/></svg>`;

  const itens = lista.map((r) => {
    const d = dataDe(r.data);
    const completo = r.inicio && r.fim;
    const pausa = r.almocoSaida && r.almocoVolta ? `almoço ${r.almocoSaida}–${r.almocoVolta}` : r.jantar ? "pausa 60 min" : "sem pausa";
    const opcao = modeloGME() ? (r.commessa ? `commessa ${rotuloCommessa(r.commessa)}` : "") : (r.centro && r.centro !== centroPadrao() ? r.centro.trim() : "");
    const meta = [completo ? `${fmtHM(minutosTrabalhados(r))} h` : "", pausa, opcao, r.obs].filter(Boolean).join(" · ");
    return `<li><button class="reg" data-acao="editar" data-data="${r.data}" aria-label="Editar ${curta(r.data)}">
        <span class="reg-dia"><b>${d.getDate()}</b><small>${SEMANA[d.getDay()]}</small></span>
        <span class="reg-meio">
          <span class="reg-faixa">${r.inicio || "--:--"} – ${r.fim || "--:--"}${viraDia(r) ? ' <span class="mais1">+1</span>' : ""}</span>
          <span class="reg-meta">${completo ? "" : '<span class="sem-fim">sem fim · </span>'}${esc(meta)}</span>
        </span>
        <span class="reg-valor">${completo ? reais.format(valorDoDia(r)) : ""}</span>
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
      <div><span class="rotulo">Total</span><span class="estat">${reais.format(valor)}</span></div>
    </div>
    ${lista.length ? `<div class="linha"><span class="lbl-forte cresce">Comprovante do mês</span><button class="link" data-acao="comprovante">Gerar imagem</button></div>` : ""}
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
  const novo = { arquivo: arq.name, importadoEm: new Date().toISOString(), funcionario: lido.funcionario, centros: lido.centros, commessas: lido.commessas, base64: paraBase64(buffer) };
  try { localStorage.setItem(lido.tipo === "gme" ? CHAVE_MODELO_GME : CHAVE_MODELO, JSON.stringify(novo)); } catch {
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
        <p class="sub-folha">Selecione no app Arquivos uma planilha de horas de qualquer mês (COMAU ou GME). Ela fica guardada só neste aparelho e serve de base para as próximas.</p>
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
      tipo, ano, mes, centroPadrao: centroPadrao(), commessaPadrao: dados.config.commessaPadrao,
      // Na GME, o fim de um turno do mês anterior pode cair no dia 1, por isso vão todos os registros do cliente.
      registros: tipo === "gme" ? Object.values(dados.registros).filter((x) => x.cliente === "gme") : doMes,
    });
    const valor = doMes.filter((x) => x.fim).reduce((s, x) => s + valorDoDia(x), 0);
    const arquivo = new File([r.arquivo], nomeArquivo(ano, mes, info.funcionario, tipo), { type: tipo === "gme" ? MIME_XLSX : MIME_XLSM });
    // O comprovante já fica pronto aqui: o iPhone só compartilha arquivos prontos no toque do botão.
    const comprovante = await criarComprovante(ano, mes);
    abrirDialogo(`<div class="folha">
        <div class="icone-resultado ok"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
        <h2>Planilha gerada</h2>
        <p class="sub-folha">${esc(arquivo.name)}<br>+ comprovante com os ${doMes.length} dias registrados</p>
        <div class="estatisticas" style="margin-bottom:20px">
          <div><span class="rotulo">Dias</span><span class="estat">${r.dias}</span></div>
          <div><span class="rotulo">Horas</span><span class="estat">${fmtHM(Math.round(r.totalHoras * 60))}</span></div>
          <div><span class="rotulo">Total</span><span class="estat">${reais.format(valor)}</span></div>
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
async function criarComprovante(ano, mes) {
  const lista = registrosDoMes(ano, mes);
  const completos = lista.filter((r) => r.inicio && r.fim);
  const agora = new Date();
  const funcionario = modelos[ui.cliente]?.funcionario || "";
  const dois = doisClientes();
  const imagem = await desenharComprovante({
    funcionario,
    titulo: `${MESES[mes - 1]} de ${ano}${dois ? ` · ${NOME_CLIENTE[ui.cliente]}` : ""}`,
    coluna: modeloGME() ? "COMMESSA" : "EQUIPE",
    geradoEm: `${agora.toLocaleDateString("pt-BR")} às ${horaAgora()}`,
    totais: {
      dias: completos.length,
      horas: fmtHM(completos.reduce((s, r) => s + minutosTrabalhados(r), 0)),
    },
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
  const partes = ["Comprovante_Horas", primeiro, dois && NOME_CLIENTE[ui.cliente], pad(mes), ano].filter(Boolean);
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

    <h2 class="secao">Novos registros${doisClientes() ? ` · ${NOME_CLIENTE[ui.cliente]}` : ""}</h2>
    <div class="linha">
      <span class="lbl-forte cresce">Preço por hora padrão</span>
      <span class="unidade">R$</span>
      <input id="a-taxa" class="numero" inputmode="decimal" value="${taxaTexto(taxaPadrao())}" placeholder="0" aria-label="Preço por hora padrão">
    </div>
    ${modeloGME() ? `<div class="linha empilhada">
      <span class="lbl-forte">Commessa padrão</span>
      ${campoCommessa("a-commessa", c.commessaPadrao, "Commessa padrão")}
    </div>` : `${modelos.comau ? `<div class="linha empilhada">
      <span class="lbl-forte">Centro de custo padrão</span>
      ${campoCentro("a-centro", centroPadrao(), "Centro de custo padrão")}
    </div>` : ""}<div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Pausa de 60 min por padrão</span>
        <span class="dica">Desmarque se costuma não jantar</span>
      </div>
      <button class="interruptor" data-acao="pausa-padrao" role="switch" aria-checked="${c.jantarPadrao}" aria-label="Pausa de 60 minutos por padrão"><span></span></button>
    </div>`}
    <div class="linha">
      <button class="link encostado" data-acao="aplicar-taxa">Aplicar preço a ${MESES[ui.mes - 1].toLowerCase()}</button>
      <button class="link" data-acao="salvar-ajustes">Salvar</button>
    </div>

    <h2 class="secao">Modelo da planilha</h2>
    ${linhaModelo("comau")}
    ${linhaModelo("gme")}
    <p class="dica" style="margin:10px 0 0">Qualquer planilha já feita serve. Arquivo com GME no nome vira o modelo da GME. Fica só neste aparelho.</p>

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

function salvarTaxaPadrao() {
  // O padrão pode ficar em 0 (sem preço pré-preenchido); retorna null só se o texto for inválido.
  const campo = $("#a-taxa").value.trim();
  const taxa = /^0*([.,]0*)?$/.test(campo) ? 0 : lerTaxa(campo);
  if (taxa === null) { mostrarErro("Ajustes não salvos", "Informe o preço por hora (ex.: 50 ou 47,50)."); return null; }
  dados.config[modeloGME() ? "taxaPadraoGME" : "taxaPadrao"] = taxa;
  const commessa = $("#a-commessa");
  if (commessa) dados.config.commessaPadrao = commessa.value.trim();
  const centro = $("#a-centro");
  if (centro) dados.config.centroPadrao = centro.value;
  return salvar() ? taxa : null;
}

function aplicarTaxaAoMes() {
  const taxa = salvarTaxaPadrao();
  if (taxa === null) return;
  if (!taxa) return mostrarErro("Nada para aplicar", "Informe um preço por hora maior que zero (ex.: 50 ou 47,50).");
  const lista = registrosDoMes(ui.ano, ui.mes);
  if (!lista.length) return mostrarErro("Nada para alterar", `Não há registros em ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}.`);
  if (!confirm(`Aplicar ${reais.format(taxa)}/h aos ${lista.length} dias de ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}?`)) return;
  lista.forEach((r) => { r.taxa = taxa; });
  if (salvar()) avisar(`Preço aplicado a ${lista.length} dias`);
  render();
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
  dados = { registros: normalizarRegistros(lido.registros), config: { ...PADRAO.config, ...lido.config, ultimoBackup: dataBackup, alteradoEm: dataBackup, pastaBackup: null, pastaFalhou: false } };
  // Backups antigos trazem um único modelo, no campo "modelo".
  const doBackup = lido.modelos || (lido.modelo ? { [ehGME(lido.modelo.arquivo) ? "gme" : "comau"]: lido.modelo } : {});
  for (const [tipo, m] of Object.entries(doBackup)) {
    if (!m?.base64) continue;
    try { localStorage.setItem(tipo === "gme" ? CHAVE_MODELO_GME : CHAVE_MODELO, JSON.stringify(m)); modelos[tipo] = m; } catch { /* registros restaurados mesmo assim */ }
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
  if (!d.open) d.showModal();
  // Foca o título, e não o primeiro campo, para o iPhone não abrir teclado/seletor sozinho.
  const titulo = d.querySelector("h2");
  if (titulo) { titulo.tabIndex = -1; titulo.focus({ preventScroll: true }); }
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
    "alternar-tema": () => escolherTema(temaEfetivo() === "dark" ? "light" : "dark"),
    tema: () => escolherTema(alvo.dataset.valor),
    "pausa-padrao": () => { dados.config.jantarPadrao = !dados.config.jantarPadrao; if (salvar()) render(); },
    "salvar-ajustes": () => { if (salvarTaxaPadrao() !== null) { avisar("Ajustes salvos"); render(); } },
    "aplicar-taxa": aplicarTaxaAoMes,
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
