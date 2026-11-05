import { ErroPlanilha, gerarPlanilha, lerModelo, minutosDe, nomeArquivo } from "./xlsm.js";
import { desenharComprovante } from "./comprovante.js";

const CHAVE = "ponto-comau.v1";
const CHAVE_MODELO = "ponto-comau.modelo";
const CHAVE_TEMA = "ponto-comau.tema";
const MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
const SEMANA = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
const MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12";

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
function minutosTrabalhados(r) {
  if (!horaOk(r.inicio) || !horaOk(r.fim)) return 0;
  let min = minutosDe(r.fim) - minutosDe(r.inicio);
  if (min < 0) min += 1440;
  return Math.max(0, min - (r.jantar ? 60 : 0));
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

// ---------- Armazenamento (só neste aparelho) ----------
const PADRAO = { registros: {}, config: { taxaPadrao: 0, centroPadrao: null, jantarPadrao: true, ultimoBackup: null } };
let dados = carregar();
let infoModelo = carregarModelo(); // { arquivo, importadoEm, funcionario, centros, base64 } ou null

function carregar() {
  try {
    const salvo = JSON.parse(localStorage.getItem(CHAVE));
    if (salvo && salvo.registros) return { registros: salvo.registros, config: { ...PADRAO.config, ...salvo.config } };
  } catch { /* sem dados ou armazenamento bloqueado */ }
  return structuredClone(PADRAO);
}

function carregarModelo() {
  try { return JSON.parse(localStorage.getItem(CHAVE_MODELO)); } catch { return null; }
}

function salvar() {
  try {
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
  .filter((r) => r.data.startsWith(`${ano}-${pad(mes)}-`))
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
const ui = { aba: "registrar", ano: hoje.getFullYear(), mes: hoje.getMonth() + 1 };

// Formulário da tela Registrar: sempre começa um ponto novo. `original` é a data do
// registro em edição, aberto pela lista de Registros (ou null se novo).
let form;
const formNovo = (data = chaveData(new Date())) => ({
  data, inicio: "", fim: "", jantar: dados.config.jantarPadrao, taxa: dados.config.taxaPadrao, obs: "", original: null, alterado: false,
});
const formDe = (r) => ({
  data: r.data, inicio: r.inicio || "", fim: r.fim || "", jantar: !!r.jantar, taxa: r.taxa ?? dados.config.taxaPadrao,
  obs: r.obs || "", original: r.data, alterado: false,
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
    registros: ["Registros", "Horas do mês", telaRegistros, '<button class="primario" data-acao="gerar">Gerar planilha COMAU</button>'],
    ajustes: ["Ajustes", "Preferências do app", telaAjustes, ""],
  };
  const [titulo, sub, tela, acoes] = telas[ui.aba];
  $("#cab-titulo").textContent = titulo;
  $("#cab-sub").textContent = sub;
  $("#conteudo").innerHTML = tela();
  $("#acoes").innerHTML = acoes;
  if (ui.aba === "registrar") atualizarCalculo();
  if (ui.aba === "registros") $('[data-acao="gerar"]').disabled = !registrosDoMes(ui.ano, ui.mes).length;
}

// ---------- Tela: Registrar ----------
// Campo de hora digitado no teclado numérico do celular (em vez do seletor de rolagem).
const campoHora = (id, valor, rotulo) => `<input type="text" id="${id}" class="hora" value="${valor}"
  inputmode="numeric" pattern="[0-9]*" maxlength="5" placeholder="--:--" autocomplete="off" enterkeyhint="next" aria-label="${rotulo}">`;

function telaRegistrar() {
  const f = form;
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
      <div class="linha">
        <span class="lbl cresce">Início</span>
        ${campoHora("f-inicio", f.inicio, "Início")}
        <button class="link" data-acao="agora" data-campo="inicio">Agora</button>
      </div>
      <div class="linha">
        <span class="lbl cresce">Fim <span class="mais1" id="c-mais1" hidden>+1</span></span>
        ${campoHora("f-fim", f.fim, "Fim")}
        <button class="link" data-acao="agora" data-campo="fim">Agora</button>
      </div>
      <div class="linha">
        <span class="lbl">Pausa (min)</span>
        <div class="seg" role="group" aria-label="Pausa">
          <button data-acao="pausa" data-valor="0" aria-pressed="${!f.jantar}">0</button>
          <button data-acao="pausa" data-valor="60" aria-pressed="${f.jantar}">60</button>
        </div>
      </div>
      <div class="linha">
        <span class="lbl cresce">Preço por hora</span>
        <span class="unidade">R$</span>
        <input id="f-taxa" class="numero" inputmode="decimal" value="${taxaTexto(f.taxa)}" placeholder="0" aria-label="Preço por hora">
      </div>
      <div class="linha">
        <input id="f-obs" class="nota" value="${esc(f.obs)}" maxlength="120" placeholder="Observação (opcional, vai para a planilha)" aria-label="Observação">
      </div>
    </div>
    <p class="estado">${f.original ? `Editando o registro de ${curta(f.original)}` : "Novo registro"}</p>
    ${f.original ? '<button class="link perigo" data-acao="excluir">Excluir este dia</button>' : ""}`;
}

function lerFormulario() {
  form.data = $("#f-data").value;
  form.inicio = $("#f-inicio").value;
  form.fim = $("#f-fim").value;
  form.obs = $("#f-obs").value.trim();
  form.taxa = lerTaxa($("#f-taxa").value);
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
  if (!v.taxa) return "Informe o preço por hora (ex.: 50 ou 47,50). Para já vir preenchido, defina o padrão em Ajustes.";
  if (dados.registros[v.data] && v.data !== v.original) return `Já existe um registro em ${curta(v.data)}. Para alterá-lo, toque nele em Registros.`;
  return "";
}

function salvarRegistro() {
  lerFormulario();
  const erro = validar(form);
  if (erro) return mostrarErro("Não foi possível salvar", erro);
  if (form.original && form.original !== form.data) delete dados.registros[form.original];
  const { data, inicio, fim, jantar, taxa, obs } = form;
  dados.registros[data] = { data, inicio, fim, jantar, taxa, obs };
  if (!salvar()) return;
  const editando = form.original;
  avisar(fim ? `${curta(data)} salvo · ${fmtHM(minutosTrabalhados(dados.registros[data]))} h` : `Início salvo às ${inicio}`);
  pedirPersistencia();
  // O ponto salvo fica em Registros; a tela volta zerada para o próximo.
  form = formNovo();
  if (editando) navegar("registros"); else render();
}

function excluirRegistro() {
  if (!form.original || !confirm(`Excluir o dia ${curta(form.original)}?`)) return;
  delete dados.registros[form.original];
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
    const meta = [completo ? `${fmtHM(minutosTrabalhados(r))} h` : "", r.jantar ? "pausa 60 min" : "sem pausa", r.obs].filter(Boolean).join(" · ");
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
  const arq = await escolherArquivo(".xlsm,application/vnd.ms-excel.sheet.macroEnabled.12");
  if (!arq) return;
  let buffer, lido;
  try {
    buffer = await arq.arrayBuffer();
    lido = await lerModelo(window.JSZip, buffer);
  } catch (e) { return mostrarErro("Modelo inválido", mensagemDe(e)); }
  const novo = { arquivo: arq.name, importadoEm: new Date().toISOString(), funcionario: lido.funcionario, centros: lido.centros, base64: paraBase64(buffer) };
  try { localStorage.setItem(CHAVE_MODELO, JSON.stringify(novo)); } catch {
    return mostrarErro("Não foi possível salvar o modelo", "O armazenamento do aparelho está cheio ou bloqueado. Libere espaço e tente novamente.");
  }
  infoModelo = novo;
  // O centro de custo da planilha é sempre o mais usado no modelo.
  dados.config.centroPadrao = lido.centroMaisUsado;
  salvar();
  pedirPersistencia();
  avisar("Modelo configurado");
  render();
}

async function gerar() {
  const { ano, mes } = ui;
  if (!infoModelo) {
    return abrirDialogo(`<div class="folha">
        <h2>Escolha o modelo da planilha</h2>
        <p class="sub-folha">Selecione no app Arquivos uma planilha COMAU de qualquer mês (.xlsm). Ela fica guardada só neste aparelho e serve de base para as próximas.</p>
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
    const registros = registrosDoMes(ano, mes).map((r) => ({ ...r, centro: null }));
    const r = await gerarPlanilha(window.JSZip, deBase64(infoModelo.base64), {
      ano, mes, centroPadrao: dados.config.centroPadrao || infoModelo.centros[0], registros,
    });
    const valor = registros.filter((x) => x.fim).reduce((s, x) => s + valorDoDia(x), 0);
    const arquivo = new File([r.arquivo], nomeArquivo(ano, mes, infoModelo.funcionario), { type: MIME_XLSM });
    // O comprovante já fica pronto aqui: o iPhone só compartilha arquivos prontos no toque do botão.
    const comprovante = await criarComprovante(ano, mes);
    abrirDialogo(`<div class="folha">
        <div class="icone-resultado ok"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
        <h2>Planilha gerada</h2>
        <p class="sub-folha">${esc(arquivo.name)}<br>+ comprovante com os ${registros.length} dias registrados</p>
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
    botao.textContent = "Gerar planilha COMAU";
  }
}

// Abre o "Compartilhar" do iPhone com os arquivos (Salvar em Arquivos, WhatsApp, e-mail...).
// Sem suporte a compartilhar (ex.: computador), baixa os arquivos.
async function entregarArquivos(arquivos) {
  if (navigator.canShare?.({ files: arquivos })) {
    try { await navigator.share({ files: arquivos }); return true; }
    catch (e) { if (e.name === "AbortError") return false; }
  }
  for (const arquivo of arquivos) {
    const url = URL.createObjectURL(arquivo);
    const a = Object.assign(document.createElement("a"), { href: url, download: arquivo.name });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    await new Promise((r) => setTimeout(r, 400));
  }
  return true;
}

// ---------- Comprovante do mês ----------
async function criarComprovante(ano, mes) {
  const lista = registrosDoMes(ano, mes);
  const completos = lista.filter((r) => r.inicio && r.fim);
  const agora = new Date();
  const imagem = await desenharComprovante({
    funcionario: infoModelo?.funcionario || "",
    titulo: `${MESES[mes - 1]} de ${ano}`,
    geradoEm: `${agora.toLocaleDateString("pt-BR")} às ${horaAgora()}`,
    totais: {
      dias: completos.length,
      horas: fmtHM(completos.reduce((s, r) => s + minutosTrabalhados(r), 0)),
      valor: reais.format(completos.reduce((s, r) => s + valorDoDia(r), 0)),
    },
    linhas: lista.map((r) => {
      const d = dataDe(r.data);
      return {
        dia: pad(d.getDate()), semana: SEMANA[d.getDay()],
        faixa: `${r.inicio || "--:--"} – ${r.fim || "--:--"}`, mais1: viraDia(r),
        pausa: r.jantar ? "60 min" : "0 min", horas: fmtHM(minutosTrabalhados(r)),
        valor: reais.format(valorDoDia(r)), obs: r.obs || "", semFim: !r.fim,
      };
    }),
  });
  const primeiro = (infoModelo?.funcionario || "").trim().split(/\s+/)[0];
  return new File([imagem], `Comprovante_Horas_${primeiro ? primeiro + "_" : ""}${pad(mes)}_${ano}.png`, { type: "image/png" });
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
  return `<h2 class="secao">Aparência</h2>
    <div class="linha">
      <span class="lbl-forte">Tema</span>
      <div class="seg" role="group" aria-label="Tema">${opcaoTema("auto", "Auto")}${opcaoTema("light", "Claro")}${opcaoTema("dark", "Escuro")}</div>
    </div>

    <h2 class="secao">Novos registros</h2>
    <div class="linha">
      <span class="lbl-forte cresce">Preço por hora padrão</span>
      <span class="unidade">R$</span>
      <input id="a-taxa" class="numero" inputmode="decimal" value="${taxaTexto(c.taxaPadrao)}" placeholder="0" aria-label="Preço por hora padrão">
    </div>
    <div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Pausa de 60 min por padrão</span>
        <span class="dica">Desmarque se costuma não jantar</span>
      </div>
      <button class="interruptor" data-acao="pausa-padrao" role="switch" aria-checked="${c.jantarPadrao}" aria-label="Pausa de 60 minutos por padrão"><span></span></button>
    </div>
    <div class="linha">
      <button class="link encostado" data-acao="aplicar-taxa">Aplicar preço a ${MESES[ui.mes - 1].toLowerCase()}</button>
      <button class="link" data-acao="salvar-ajustes">Salvar</button>
    </div>

    <h2 class="secao">Modelo da planilha</h2>
    <div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte reticencias">${infoModelo ? esc(infoModelo.arquivo) : "Nenhum modelo escolhido"}</span>
        <span class="dica">${infoModelo ? `Escolhido em ${new Date(infoModelo.importadoEm).toLocaleDateString("pt-BR")} · fica só neste aparelho` : "Qualquer planilha COMAU já feita serve"}</span>
      </div>
      <button class="link" data-acao="modelo">${infoModelo ? "Trocar" : "Escolher"}</button>
    </div>

    <h2 class="secao">Backup</h2>
    <div class="linha alta">
      <div class="cresce pilha-v">
        <span class="lbl-forte">Último backup</span>
        <span class="dica">${ultimo} · inclui registros, ajustes e modelo</span>
      </div>
      <button class="link" data-acao="backup">Fazer agora</button>
    </div>
    <div class="linha">
      <span class="lbl-forte cresce">Restaurar de um backup</span>
      <button class="link" data-acao="restaurar">Restaurar</button>
    </div>
    <p class="estado" style="text-align:center;margin-top:28px">Ponto Automisa · ${Object.keys(dados.registros).length} dias registrados</p>`;
}

function salvarTaxaPadrao() {
  // O padrão pode ficar em 0 (sem preço pré-preenchido); retorna null só se o texto for inválido.
  const campo = $("#a-taxa").value.trim();
  const taxa = /^0*([.,]0*)?$/.test(campo) ? 0 : lerTaxa(campo);
  if (taxa === null) { mostrarErro("Ajustes não salvos", "Informe o preço por hora (ex.: 50 ou 47,50)."); return null; }
  dados.config.taxaPadrao = taxa;
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

async function fazerBackup() {
  // O backup vai para onde você escolher (ex.: iCloud Drive); nada é enviado à internet pelo app.
  const conteudo = JSON.stringify({ app: "ponto-comau", versao: 1, exportadoEm: new Date().toISOString(), ...dados, modelo: infoModelo }, null, 2);
  const ok = await entregarArquivos([new File([conteudo], `backup-ponto-${chaveData(new Date())}.json`, { type: "application/json" })]);
  if (ok) {
    dados.config.ultimoBackup = new Date().toISOString();
    salvar();
    avisar("Backup criado");
    render();
  }
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
  dados = { registros: lido.registros, config: { ...PADRAO.config, ...lido.config } };
  if (lido.modelo?.base64) {
    try { localStorage.setItem(CHAVE_MODELO, JSON.stringify(lido.modelo)); infoModelo = lido.modelo; } catch { /* registros restaurados mesmo assim */ }
  }
  if (salvar()) avisar(`${qtd} dias restaurados`);
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
      form.alterado = true;
      alvo.parentElement.querySelectorAll("button").forEach((b) => b.setAttribute("aria-pressed", String(b === alvo)));
      atualizarCalculo();
    },
    salvar: salvarRegistro, excluir: excluirRegistro, limpar: () => { form = formNovo(); render(); },
    editar: () => { form = formDe(dados.registros[alvo.dataset.data]); navegar("registrar", { manterForm: true }); },
    "mes-anterior": () => mudarMes(-1), "mes-seguinte": () => mudarMes(1), gerar, comprovante: gerarSoComprovante,
    "alternar-tema": () => escolherTema(temaEfetivo() === "dark" ? "light" : "dark"),
    tema: () => escolherTema(alvo.dataset.valor),
    "pausa-padrao": () => { dados.config.jantarPadrao = !dados.config.jantarPadrao; if (salvar()) render(); },
    "salvar-ajustes": () => { if (salvarTaxaPadrao() !== null) { avisar("Ajustes salvos"); render(); } },
    "aplicar-taxa": aplicarTaxaAoMes,
    modelo: () => { fecharDialogo(); escolherModelo(); }, backup: fazerBackup, restaurar: restaurarBackup, fechar: fecharDialogo,
  };
  if (acoes[alvo.dataset.acao]) { e.preventDefault(); acoes[alvo.dataset.acao](); }
});

document.addEventListener("input", (e) => {
  if (ui.aba !== "registrar" || !e.target.id?.startsWith("f-") || e.target.id === "f-data") return;
  form.alterado = true;
  if (e.target.classList.contains("hora")) {
    mascaraHora(e.target);
    // Início completo: pula para o fim.
    if (e.target.id === "f-inicio" && horaOk(e.target.value) && !$("#f-fim").value) $("#f-fim").focus();
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
  if (!document.hidden && ui.aba === "registrar" && !form.alterado) navegar("registrar");
});

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  navigator.serviceWorker.register("sw.js").catch(() => { /* app funciona sem offline */ });
}

aplicarTema();
navegar("registrar");
