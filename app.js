import { ErroPlanilha, gerarPlanilha, lerModelo, minutosDe, nomeArquivo } from "./xlsm.js";

const CHAVE = "ponto-comau.v1";
const CHAVE_MODELO = "ponto-comau.modelo";
const MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
const SEMANA = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
const MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12";

const $ = (sel, raiz = document) => raiz.querySelector(sel);
const pad = (n) => String(n).padStart(2, "0");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const reais = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const ICONE_RELOGIO = '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>';

// ---------- Datas ----------
const chaveData = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const dataDe = (chave) => { const [a, m, d] = chave.split("-").map(Number); return new Date(a, m - 1, d); };
const horaAgora = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const curta = (chave) => `${chave.slice(8, 10)}/${chave.slice(5, 7)}`;
const ontem = () => { const d = new Date(); d.setDate(d.getDate() - 1); return chaveData(d); };
const mesCurto = (d) => d.toLocaleDateString("pt-BR", { month: "short" }).replace(".", "");
const tituloDia = (chave) => { const d = dataDe(chave); return `${SEMANA[d.getDay()]}. ${d.getDate()} ${mesCurto(d)}. ${String(d.getFullYear()).slice(2)}`; };

// ---------- Cálculos ----------
function minutosTrabalhados(r) {
  if (!r.inicio || !r.fim) return 0;
  let min = minutosDe(r.fim) - minutosDe(r.inicio);
  if (min < 0) min += 1440;
  return Math.max(0, min - (r.jantar ? 60 : 0));
}
const valorDoDia = (r) => Math.round((minutosTrabalhados(r) / 60) * (r.taxa || 0) * 100) / 100;
const fmtHM = (min) => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
const viraDia = (r) => r.inicio && r.fim && minutosDe(r.fim) < minutosDe(r.inicio);

function lerTaxa(texto) {
  const v = Number(String(texto).trim().replace(/\./g, "").replace(",", "."));
  return Number.isFinite(v) && v > 0 ? v : null;
}
const taxaTexto = (v) => (v == null ? "" : String(v).replace(".", ","));

// ---------- Armazenamento ----------
const PADRAO = { registros: {}, config: { taxaPadrao: 50, centroPadrao: null, jantarPadrao: true, ultimoBackup: null } };
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

function turnoAberto() {
  const limite = ontem();
  return Object.values(dados.registros)
    .filter((r) => r.inicio && !r.fim && r.data >= limite)
    .sort((a, b) => b.data.localeCompare(a.data))[0];
}

// ---------- Estado da interface ----------
const hoje = new Date();
const ui = { aba: "registrar", ano: hoje.getFullYear(), mes: hoje.getMonth() + 1, crescente: true };

// Formulário da tela Registrar. `original` é a data do registro em edição (ou null se novo).
let form;
const formNovo = (data = chaveData(new Date())) => ({
  data, inicio: "", fim: "", jantar: dados.config.jantarPadrao, taxa: dados.config.taxaPadrao, obs: "", original: null, alterado: false,
});
const formDe = (r) => ({
  data: r.data, inicio: r.inicio || "", fim: r.fim || "", jantar: !!r.jantar, taxa: r.taxa ?? dados.config.taxaPadrao,
  obs: r.obs || "", original: r.data, alterado: false,
});
function formInicial() {
  const aberto = turnoAberto();
  if (aberto) return formDe(aberto);
  const doDia = dados.registros[chaveData(new Date())];
  return doDia ? formDe(doDia) : formNovo();
}

function navegar(aba, { manterForm = false } = {}) {
  ui.aba = aba;
  if (aba === "registrar" && !manterForm) form = formInicial();
  document.querySelectorAll(".aba").forEach((b) =>
    b.dataset.aba === aba ? b.setAttribute("aria-current", "page") : b.removeAttribute("aria-current"));
  render();
  window.scrollTo(0, 0);
}

function render() {
  const telas = { registrar: telaRegistrar, registros: telaRegistros, ajustes: telaAjustes };
  $("#conteudo").innerHTML = telas[ui.aba]();
  if (ui.aba === "registrar") atualizarCalculo();
}

// ---------- Tela: Registrar ----------
function telaRegistrar() {
  const f = form;
  const estado = f.original ? `Editando o registro de ${curta(f.original)}` : "Novo registro";
  return `<div class="grade-2 tempo-valor">
      <div class="cartao">
        <span class="rotulo">Tempo</span><div class="grande num" id="c-tempo">00:00</div>
        <span class="rotulo">Minutos</span><div class="grande num" id="c-minutos">0</div>
      </div>
      <div class="cartao valor-total"><span class="rotulo">Valor total</span><div class="grande num" id="c-valor">R$ 0,00</div></div>
    </div>

    <label class="cartao linha-icone">
      <span class="icone icone-data"><svg viewBox="0 0 24 24"><rect x="3.5" y="5" width="17" height="15" rx="2"/><path d="M3.5 10h17M8 3v4M16 3v4"/></svg></span>
      <span style="flex:1"><span class="rotulo">Data</span><input type="date" id="f-data" value="${f.data}" required></span>
    </label>

    <div class="grade-2" style="margin-top:12px">
      <div class="cartao horario">
        <div class="horario-cabeca"><strong>Início</strong></div>
        <input type="time" id="f-inicio" value="${f.inicio}" aria-label="Início">
        <button class="agora" data-acao="agora" data-campo="inicio">Agora</button>
      </div>
      <div class="cartao horario">
        <div class="horario-cabeca"><strong>Fim</strong><span class="chip-mais1" id="c-mais1" hidden>+1</span></div>
        <input type="time" id="f-fim" value="${f.fim}" aria-label="Fim">
        <button class="agora" data-acao="agora" data-campo="fim">Agora</button>
      </div>
    </div>

    <div class="cartao pausa-taxa" style="margin-top:12px">
      <div>
        <span class="rotulo">Pausa (min)</span>
        <div class="segmento" role="group" aria-label="Pausa">
          <button data-acao="pausa" data-valor="0" aria-pressed="${!f.jantar}">0</button>
          <button data-acao="pausa" data-valor="60" aria-pressed="${f.jantar}">60</button>
        </div>
      </div>
      <div>
        <span class="rotulo">Preço por hora</span>
        <input id="f-taxa" inputmode="decimal" value="${taxaTexto(f.taxa)}" aria-label="Preço por hora">
      </div>
    </div>

    <label class="cartao linha-icone" style="margin-top:12px">
      <span class="icone icone-obs"><svg viewBox="0 0 24 24"><path d="M4 5h16v11H9l-5 4V5Z"/></svg></span>
      <span style="flex:1"><span class="rotulo">Observação</span><input id="f-obs" value="${esc(f.obs)}" maxlength="120" placeholder="Opcional — vai para a planilha"></span>
    </label>

    <div class="acoes">
      <button class="btn btn-icone" data-acao="limpar" aria-label="Limpar e começar um novo registro">
        <svg viewBox="0 0 24 24"><path d="M4 12a8 8 0 1 0 2.3-5.6M4 4v5h5"/></svg>
      </button>
      <button class="btn btn-primario" data-acao="salvar">Salvar</button>
    </div>
    <p class="estado">${estado}</p>
    ${f.original ? '<button class="btn btn-perigo" data-acao="excluir">Excluir este dia</button>' : ""}`;
}

function lerFormulario() {
  form.data = $("#f-data").value;
  form.inicio = $("#f-inicio").value;
  form.fim = $("#f-fim").value;
  form.obs = $("#f-obs").value.trim();
  form.taxaTexto = $("#f-taxa").value;
  form.taxa = lerTaxa(form.taxaTexto);
}

function atualizarCalculo() {
  lerFormulario();
  const min = minutosTrabalhados(form);
  $("#c-tempo").textContent = fmtHM(min);
  $("#c-minutos").textContent = String(min);
  $("#c-valor").textContent = reais.format(valorDoDia(form));
  $("#c-mais1").hidden = !viraDia(form);
}

function validar(v) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v.data)) return "Informe a data.";
  if (!v.inicio) return "Informe o horário de início.";
  if (v.fim && v.fim === v.inicio) return "Início e fim não podem ser iguais.";
  if (v.fim && v.jantar && minutosTrabalhados({ ...v, jantar: false }) <= 60) return "O turno tem 1h ou menos: use pausa 0 ou corrija os horários.";
  if (!v.taxa) return "Informe o preço por hora (ex.: 50 ou 47,50).";
  if (dados.registros[v.data] && v.data !== v.original) return `Já existe um registro em ${curta(v.data)}.`;
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
  form = formDe(dados.registros[data]);
  avisar(fim ? `Salvo · ${fmtHM(minutosTrabalhados(form))}` : `Início salvo às ${inicio}`);
  pedirPersistencia();
  render();
}

function excluirRegistro() {
  if (!form.original || !confirm(`Excluir o dia ${curta(form.original)}?`)) return;
  delete dados.registros[form.original];
  if (salvar()) avisar(`${curta(form.original)} excluído`);
  form = formNovo();
  render();
}

function trocarData() {
  const nova = $("#f-data").value;
  if (!nova) return;
  const existente = dados.registros[nova];
  if (existente && nova !== form.original) { form = formDe(existente); render(); return; }
  atualizarCalculo();
}

// ---------- Tela: Registros ----------
function telaRegistros() {
  const { ano, mes } = ui;
  const lista = registrosDoMes(ano, mes);
  if (!ui.crescente) lista.reverse();
  const completos = lista.filter((r) => r.inicio && r.fim);
  const minutos = completos.reduce((s, r) => s + minutosTrabalhados(r), 0);
  const valor = completos.reduce((s, r) => s + valorDoDia(r), 0);
  const titulo = new Date(ano, mes - 1, 1).toLocaleDateString("pt-BR", { month: "short", year: "numeric" });
  const seta = (d) => `<svg viewBox="0 0 24 24"><path d="${d}"/></svg>`;

  const dias = lista.map((r) => {
    const completo = r.inicio && r.fim;
    const min = minutosTrabalhados(r);
    const extras = [
      viraDia(r) ? '<span class="chip-mais1">+1</span>' : "",
      completo && r.jantar ? `<span class="chip chip-neutro">${ICONE_RELOGIO.replace("<svg", '<svg style="width:14px;height:14px"')}60 min</span>` : "",
      !r.fim ? '<span class="chip chip-alerta">sem fim</span>' : "",
    ].join("");
    return `<li class="dia">
        <div class="dia-cabeca">
          <strong>${tituloDia(r.data)}</strong>
          <span class="chips num">
            <span class="chip chip-tempo">${ICONE_RELOGIO}${fmtHM(min)}</span>
            <span class="chip chip-dinheiro">${reais.format(valorDoDia(r))}</span>
          </span>
        </div>
        <button class="turno" data-acao="editar" data-data="${r.data}" aria-label="Editar ${curta(r.data)}">
          <span>
            <span class="extra"><span class="horas num">${r.inicio || "--:--"} – ${r.fim || "--:--"}</span>${extras}</span>
            ${r.obs ? `<span class="obs">${esc(r.obs)}</span>` : ""}
          </span>
          <span class="dinheiro num">${completo ? reais.format(valorDoDia(r)) : ""}</span>
        </button>
      </li>`;
  }).join("");

  return `<div class="filtros">
      <button class="btn-redondo" data-acao="mes-anterior" aria-label="Mês anterior">${seta("M15 6l-6 6 6 6")}</button>
      <h1>${titulo}</h1>
      <button class="btn-redondo" data-acao="mes-seguinte" aria-label="Próximo mês">${seta("M9 6l6 6-6 6")}</button>
      <button class="btn-redondo" data-acao="ordenar" aria-label="${ui.crescente ? "Mostrar mais recentes primeiro" : "Mostrar mais antigos primeiro"}">${seta(ui.crescente ? "M12 19V5M6 11l6-6 6 6" : "M12 5v14M6 13l6 6 6-6")}</button>
    </div>
    <div class="cartao totais">
      <span class="chips num">
        <span class="chip chip-neutro">${completos.length} dias</span>
        <span class="chip chip-tempo">${ICONE_RELOGIO}${fmtHM(minutos)}</span>
        <span class="chip chip-dinheiro">${reais.format(valor)}</span>
      </span>
    </div>
    <button class="btn btn-primario" style="margin-top:12px" data-acao="gerar" ${lista.length ? "" : "disabled"}>Gerar planilha COMAU</button>
    ${lista.length ? `<ul class="dias">${dias}</ul>` : `<div class="vazio">Nenhum registro em ${MESES[mes - 1].toLowerCase()}.</div>`}`;
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
        <p class="sub">Selecione no app Arquivos uma planilha COMAU de qualquer mês (.xlsm). Ela fica guardada só neste aparelho e serve de base para as próximas.</p>
        <div class="pilha">
          <button class="btn btn-primario" data-acao="modelo">Escolher planilha modelo</button>
          <button class="btn" data-acao="fechar">Agora não</button>
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
    abrirDialogo(`<div class="folha">
        <div class="icone-resultado ok"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
        <h2>Planilha gerada</h2>
        <p class="sub">${esc(arquivo.name)}</p>
        <div class="cartao totais" style="margin-bottom:16px">
          <span class="chips num">
            <span class="chip chip-neutro">${r.dias} dias</span>
            <span class="chip chip-tempo">${ICONE_RELOGIO}${fmtHM(Math.round(r.totalHoras * 60))}</span>
            <span class="chip chip-dinheiro">${reais.format(valor)}</span>
          </span>
        </div>
        <div class="pilha">
          <button class="btn btn-primario" data-acao="compartilhar">Salvar ou compartilhar</button>
          <button class="btn" data-acao="fechar">Fechar</button>
        </div>
      </div>`);
    $('[data-acao="compartilhar"]').addEventListener("click", () => entregarArquivo(arquivo));
  } catch (e) {
    mostrarErro("Não foi possível gerar a planilha", mensagemDe(e));
  } finally {
    botao.disabled = false;
    botao.textContent = "Gerar planilha COMAU";
  }
}

async function entregarArquivo(arquivo) {
  if (navigator.canShare?.({ files: [arquivo] })) {
    try { await navigator.share({ files: [arquivo], title: arquivo.name }); return true; }
    catch (e) { if (e.name === "AbortError") return false; }
  }
  const url = URL.createObjectURL(arquivo);
  const a = Object.assign(document.createElement("a"), { href: url, download: arquivo.name });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  return true;
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
  const ultimo = c.ultimoBackup ? new Date(c.ultimoBackup).toLocaleDateString("pt-BR") : "nunca";
  const modelo = infoModelo
    ? `<div class="cartao">
        <strong style="word-break:break-word">${esc(infoModelo.arquivo)}</strong>
        <p class="mudo" style="margin:4px 0 0;font-size:14px">Escolhido em ${new Date(infoModelo.importadoEm).toLocaleDateString("pt-BR")}</p>
      </div>
      <button class="btn" style="margin-top:10px" data-acao="modelo">Trocar modelo</button>`
    : `<div class="cartao"><p style="margin:0">Nenhum modelo escolhido.</p></div>
      <button class="btn" style="margin-top:10px" data-acao="modelo">Escolher modelo</button>`;
  return `<section class="secao">
      <h2>Modelo da planilha</h2>
      ${modelo}
      <p class="dica">Qualquer planilha COMAU já feita serve. Ela fica guardada só neste aparelho.</p>
    </section>
    <form id="ajustes" novalidate>
      <section class="secao">
        <h2>Novos registros</h2>
        <label class="campo"><span>Preço por hora padrão (R$)</span><input type="text" name="taxa" inputmode="decimal" value="${taxaTexto(c.taxaPadrao)}"></label>
        <div class="cartao linha-switch">
          <div>Pausa de 60 min por padrão<small>Desmarque se costuma não jantar</small></div>
          <label class="switch"><input type="checkbox" name="jantar" ${c.jantarPadrao ? "checked" : ""} aria-label="Pausa de 60 minutos por padrão"><span></span></label>
        </div>
        <div class="pilha" style="margin-top:14px">
          <button class="btn btn-primario" type="submit">Salvar ajustes</button>
          <button class="btn" type="button" data-acao="aplicar-taxa">Aplicar preço aos dias de ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}</button>
        </div>
      </section>
    </form>
    <section class="secao">
      <h2>Backup</h2>
      <div class="cartao"><p style="margin:0">Os registros ficam só neste aparelho. Último backup: <strong>${ultimo}</strong>.</p></div>
      <div class="grade-2" style="margin-top:10px">
        <button class="btn" data-acao="backup">Fazer backup</button>
        <button class="btn" data-acao="restaurar">Restaurar</button>
      </div>
    </section>
    <p class="mudo" style="text-align:center;font-size:13px;margin-top:32px">Ponto COMAU · ${Object.keys(dados.registros).length} dias registrados</p>`;
}

function salvarAjustes(f) {
  const taxa = lerTaxa(f.taxa.value);
  if (!taxa) { mostrarErro("Ajustes não salvos", "Informe o preço por hora (ex.: 50 ou 47,50)."); return null; }
  Object.assign(dados.config, { taxaPadrao: taxa, jantarPadrao: f.jantar.checked });
  return salvar() ? taxa : null;
}

function aplicarTaxaAoMes() {
  const taxa = salvarAjustes($("#ajustes"));
  if (!taxa) return;
  const lista = registrosDoMes(ui.ano, ui.mes);
  if (!lista.length) return mostrarErro("Nada para alterar", `Não há registros em ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}.`);
  if (!confirm(`Aplicar ${reais.format(taxa)}/h aos ${lista.length} dias de ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}?`)) return;
  lista.forEach((r) => { r.taxa = taxa; });
  if (salvar()) avisar(`Preço aplicado a ${lista.length} dias`);
  render();
}

async function fazerBackup() {
  // O backup leva também o modelo, para restaurar tudo num aparelho novo.
  const conteudo = JSON.stringify({ app: "ponto-comau", versao: 1, exportadoEm: new Date().toISOString(), ...dados, modelo: infoModelo }, null, 2);
  const ok = await entregarArquivo(new File([conteudo], `backup-ponto-${chaveData(new Date())}.json`, { type: "application/json" }));
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
  } catch { return mostrarErro("Backup inválido", "O arquivo escolhido não é um backup do Ponto COMAU."); }
  const qtd = Object.keys(lido.registros).length;
  if (!confirm(`Restaurar ${qtd} dias do backup de ${new Date(lido.exportadoEm).toLocaleDateString("pt-BR")}? Os registros atuais deste aparelho serão substituídos.`)) return;
  dados = { registros: lido.registros, config: { ...PADRAO.config, ...lido.config } };
  if (lido.modelo?.base64) {
    try { localStorage.setItem(CHAVE_MODELO, JSON.stringify(lido.modelo)); infoModelo = lido.modelo; } catch { /* registros restaurados mesmo assim */ }
  }
  if (salvar()) avisar(`${qtd} dias restaurados`);
  render();
}

// ---------- Diálogos e confirmações ----------
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
      <p class="sub" style="color:var(--ink-2)">${esc(mensagem)}</p>
      <button class="btn btn-primario" data-acao="fechar">Entendi</button>
    </div>`);
}

const mensagemDe = (e) => e instanceof ErroPlanilha ? e.message : `Erro inesperado: ${e.message || e}`;

let avisoTimer;
function avisar(texto) {
  const el = $("#aviso");
  el.textContent = texto;
  el.classList.add("visivel");
  clearTimeout(avisoTimer);
  avisoTimer = setTimeout(() => el.classList.remove("visivel"), 2200);
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
    "mes-anterior": () => mudarMes(-1), "mes-seguinte": () => mudarMes(1),
    ordenar: () => { ui.crescente = !ui.crescente; render(); }, gerar,
    modelo: () => { fecharDialogo(); escolherModelo(); }, backup: fazerBackup, restaurar: restaurarBackup,
    "aplicar-taxa": aplicarTaxaAoMes, fechar: fecharDialogo,
  };
  if (acoes[alvo.dataset.acao]) { e.preventDefault(); acoes[alvo.dataset.acao](); }
});

document.addEventListener("input", (e) => {
  if (ui.aba === "registrar" && e.target.id?.startsWith("f-") && e.target.id !== "f-data") { form.alterado = true; atualizarCalculo(); }
});
document.addEventListener("change", (e) => {
  if (e.target.id === "f-data") { form.alterado = true; trocarData(); }
});
document.addEventListener("submit", (e) => {
  if (e.target.id !== "ajustes") return;
  e.preventDefault();
  if (salvarAjustes(e.target)) { avisar("Ajustes salvos"); render(); }
});

$("#dialogo").addEventListener("click", (e) => { if (e.target.id === "dialogo") fecharDialogo(); });
// Ao voltar para o app (ex.: no dia seguinte), recarrega o formulário se nada foi digitado.
document.addEventListener("visibilitychange", () => {
  if (!document.hidden && ui.aba === "registrar" && !form.alterado) navegar("registrar");
});

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  navigator.serviceWorker.register("sw.js").catch(() => { /* app funciona sem offline */ });
}

navegar("registrar");
