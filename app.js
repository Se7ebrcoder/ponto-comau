import { ErroPlanilha, gerarPlanilha, lerInforme, lerModelo, minutosDe, nomeArquivo } from "./xlsm.js";

const CHAVE = "ponto-comau.v1";
const CHAVE_MODELO = "ponto-comau.modelo";
const MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
const SEMANA = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
const SEMANA_LONGA = ["Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado"];
const DIAS_LEMBRETE_BACKUP = 30;
const MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12";

const $ = (sel, raiz = document) => raiz.querySelector(sel);
const pad = (n) => String(n).padStart(2, "0");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const reais = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });

// ---------- Datas ----------
const chaveData = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const dataDe = (chave) => { const [a, m, d] = chave.split("-").map(Number); return new Date(a, m - 1, d); };
const horaAgora = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const curta = (chave) => `${chave.slice(8, 10)}/${chave.slice(5, 7)}`;
const ontem = () => { const d = new Date(); d.setDate(d.getDate() - 1); return chaveData(d); };

// ---------- Cálculos ----------
function minutosTrabalhados(r) {
  if (!r.inicio || !r.fim) return 0;
  let min = minutosDe(r.fim) - minutosDe(r.inicio);
  if (min < 0) min += 1440;
  return Math.max(0, min - (r.jantar ? 60 : 0));
}
const valorDoDia = (r) => Math.round((minutosTrabalhados(r) / 60) * (r.taxa || 0) * 100) / 100;
const fmtHoras = (min) => `${Math.floor(min / 60)}h${pad(min % 60)}`;

function totaisDoMes(ano, mes) {
  const lista = registrosDoMes(ano, mes);
  const completos = lista.filter((r) => r.inicio && r.fim);
  return {
    dias: completos.length,
    minutos: completos.reduce((s, r) => s + minutosTrabalhados(r), 0),
    valor: completos.reduce((s, r) => s + valorDoDia(r), 0),
    pendentes: lista.length - completos.length,
  };
}

// ---------- Armazenamento ----------
const PADRAO = { registros: {}, config: { taxaPadrao: 50, centroPadrao: null, jantarPadrao: true, ultimoBackup: null } };
let dados = carregar();
let infoModelo = carregarModelo(); // { arquivo, importadoEm, funcionario, centros, base64 } ou null

function carregarModelo() {
  try { return JSON.parse(localStorage.getItem(CHAVE_MODELO)); } catch { return null; }
}

const centros = () => infoModelo?.centros || [];

function paraBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let texto = "";
  for (let i = 0; i < bytes.length; i += 0x8000) texto += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(texto);
}
const deBase64 = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer;

function carregar() {
  try {
    const salvo = JSON.parse(localStorage.getItem(CHAVE));
    if (salvo && salvo.registros) return { registros: salvo.registros, config: { ...PADRAO.config, ...salvo.config } };
  } catch { /* sem dados ou armazenamento bloqueado */ }
  return structuredClone(PADRAO);
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

const registrosDoMes = (ano, mes) => Object.values(dados.registros)
  .filter((r) => r.data.startsWith(`${ano}-${pad(mes)}-`))
  .sort((a, b) => b.data.localeCompare(a.data));

function turnoAberto() {
  const limite = ontem();
  return Object.values(dados.registros)
    .filter((r) => r.inicio && !r.fim && r.data >= limite)
    .sort((a, b) => b.data.localeCompare(a.data))[0];
}

// ---------- Estado da interface ----------
const hoje = new Date();
const ui = { aba: "hoje", ano: hoje.getFullYear(), mes: hoje.getMonth() + 1 };
let relogio;

function navegar(aba) {
  ui.aba = aba;
  document.querySelectorAll(".aba").forEach((b) =>
    b.dataset.aba === aba ? b.setAttribute("aria-current", "page") : b.removeAttribute("aria-current"));
  render();
  window.scrollTo(0, 0);
}

function render() {
  clearInterval(relogio);
  const conteudo = $("#conteudo");
  if (ui.aba === "hoje") { conteudo.innerHTML = telaHoje(); relogio = setInterval(atualizarRelogio, 15000); }
  else if (ui.aba === "mes") conteudo.innerHTML = telaMes();
  else conteudo.innerHTML = telaAjustes();
}

function atualizarRelogio() {
  if (ui.aba !== "hoje") return;
  const r = $("#relogio");
  if (r) r.textContent = horaAgora();
  const aberto = turnoAberto();
  const dec = $("#decorrido");
  if (aberto && dec) dec.textContent = fmtHoras(minutosTrabalhados({ ...aberto, fim: horaAgora(), jantar: false }));
}

// ---------- Tela: Hoje ----------
function faixaBackup() {
  const datas = Object.keys(dados.registros).sort();
  if (!datas.length) return "";
  // Conta a partir do último backup ou, se nunca houve, do registro mais antigo.
  const ultimo = dados.config.ultimoBackup;
  const referencia = ultimo ? new Date(ultimo) : dataDe(datas[0]);
  const dias = Math.floor((Date.now() - referencia) / 86400000);
  if (dias < DIAS_LEMBRETE_BACKUP) return "";
  const texto = ultimo ? `Último backup há ${dias} dias.` : "Você ainda não fez backup dos registros.";
  return faixa(texto, "backup", "Fazer agora");
}

function faixa(texto, acao, rotulo) {
  return `<div class="faixa faixa-alerta"><svg viewBox="0 0 24 24"><path d="M12 9v4M12 17h.01M10.3 3.9 2.4 18a2 2 0 0 0 1.7 3h15.8a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg>
    <span>${texto}</span><button data-acao="${acao}">${rotulo}</button></div>`;
}

const faixaModelo = () => infoModelo ? "" : faixa("Escolha o modelo da planilha para poder gerá-la.", "modelo", "Escolher");

function telaHoje() {
  const agora = new Date();
  const chaveHoje = chaveData(agora);
  const aberto = turnoAberto();
  const doDia = dados.registros[chaveHoje];
  const t = totaisDoMes(agora.getFullYear(), agora.getMonth() + 1);

  let cartao, botao = "";
  if (aberto) {
    const quando = aberto.data === chaveHoje ? "" : " (ontem)";
    cartao = `<div class="status"><span class="ponto ativo"></span>Em andamento</div>
      <p class="status-detalhe">Entrada às ${aberto.inicio}${quando}</p>
      <div class="decorrido num" id="decorrido">${fmtHoras(minutosTrabalhados({ ...aberto, fim: horaAgora(), jantar: false }))}</div>`;
    botao = `<div class="cartao linha-switch" style="margin-top:12px">
        <div>Jantei<small>Desconta 1h do dia</small></div>
        <label class="switch"><input type="checkbox" data-acao="jantar" ${aberto.jantar ? "checked" : ""} aria-label="Jantei"><span></span></label>
      </div>
      <button class="btn btn-primario btn-grande" data-acao="saida">Registrar saída</button>`;
  } else if (doDia && doDia.fim) {
    cartao = `<div class="status"><span class="ponto"></span>Dia registrado</div>
      <p class="status-detalhe num">${doDia.inicio} – ${doDia.fim} · ${fmtHoras(minutosTrabalhados(doDia))} · ${reais.format(valorDoDia(doDia))}</p>`;
    botao = `<button class="btn btn-grande" data-acao="editar" data-data="${chaveHoje}">Editar registro de hoje</button>`;
  } else {
    cartao = `<div class="status"><span class="ponto"></span>Nenhum registro hoje</div>
      <p class="status-detalhe">Toque em entrada quando começar o turno.</p>`;
    botao = `<button class="btn btn-primario btn-grande" data-acao="entrada">Registrar entrada</button>`;
  }

  return `${faixaModelo()}${faixaBackup()}
    <header class="cabecalho"><p>${SEMANA_LONGA[agora.getDay()]}, ${agora.getDate()} de ${MESES[agora.getMonth()].toLowerCase()}</p></header>
    <div class="relogio num" id="relogio">${horaAgora()}</div>
    <div class="cartao">${cartao}</div>
    ${botao}
    <section class="secao">
      <h2>${MESES[agora.getMonth()]}</h2>
      <a href="#" class="resumo-link" data-acao="ir-mes">
        <div class="resumo num">
          <div><small>Dias</small><strong>${t.dias}</strong></div>
          <div><small>Horas</small><strong>${fmtHoras(t.minutos)}</strong></div>
          <div><small>Valor</small><strong>${reais.format(t.valor)}</strong></div>
        </div>
      </a>
    </section>`;
}

function registrarEntrada() {
  const chave = chaveData(new Date());
  if (dados.registros[chave]) return abrirEditor(chave);
  const c = dados.config;
  dados.registros[chave] = { data: chave, inicio: horaAgora(), fim: "", jantar: c.jantarPadrao, taxa: c.taxaPadrao, centro: c.centroPadrao, obs: "" };
  if (salvar()) { avisar(`Entrada registrada às ${dados.registros[chave].inicio}`); pedirPersistencia(); }
  render();
}

function registrarSaida() {
  const aberto = turnoAberto();
  if (!aberto) return;
  aberto.fim = horaAgora();
  if (aberto.fim === aberto.inicio) {
    aberto.fim = "";
    return mostrarErro("Saída não registrada", "A saída não pode ser no mesmo minuto da entrada.");
  }
  if (salvar()) avisar(`Saída registrada · ${fmtHoras(minutosTrabalhados(aberto))}`);
  render();
}

// ---------- Tela: Mês ----------
function telaMes() {
  const { ano, mes } = ui;
  const lista = registrosDoMes(ano, mes);
  const t = totaisDoMes(ano, mes);
  const seta = (d) => `<svg viewBox="0 0 24 24"><path d="${d}"/></svg>`;

  const itens = lista.map((r) => {
    const d = dataDe(r.data);
    const completo = r.inicio && r.fim;
    const tags = [
      !r.fim ? '<span class="tag tag-alerta">sem saída</span>' : "",
      completo && !r.jantar ? '<span class="tag tag-neutra">sem janta</span>' : "",
    ].join("");
    const detalhe = [`${reais.format(r.taxa || 0)}/h`, r.centro && r.centro !== dados.config.centroPadrao ? esc(r.centro) : "", esc(r.obs)]
      .filter(Boolean).join(" · ");
    return `<li><button data-acao="editar" data-data="${r.data}">
        <span class="dia"><strong class="num">${pad(d.getDate())}</strong><small>${SEMANA[d.getDay()]}</small></span>
        <span><span class="horario num">${r.inicio || "--:--"} – ${r.fim || "--:--"}</span>${tags}<br><span class="detalhe">${detalhe}</span></span>
        <span class="valores num"><strong>${completo ? fmtHoras(minutosTrabalhados(r)) : "—"}</strong><small>${completo ? reais.format(valorDoDia(r)) : ""}</small></span>
      </button></li>`;
  }).join("");

  return `${faixaModelo()}<div class="nav-mes">
      <button data-acao="mes-anterior" aria-label="Mês anterior">${seta("M15 6l-6 6 6 6")}</button>
      <h1>${MESES[mes - 1]} ${ano}</h1>
      <button data-acao="mes-seguinte" aria-label="Próximo mês">${seta("M9 6l6 6-6 6")}</button>
    </div>
    <div class="resumo num">
      <div><small>Dias</small><strong>${t.dias}</strong></div>
      <div><small>Horas</small><strong>${fmtHoras(t.minutos)}</strong></div>
      <div><small>Valor</small><strong>${reais.format(t.valor)}</strong></div>
    </div>
    <section class="secao pilha">
      <button class="btn btn-primario" data-acao="gerar" ${lista.length ? "" : "disabled"}>Gerar planilha COMAU</button>
      <div class="grade-2">
        <button class="btn" data-acao="novo">Adicionar dia</button>
        <button class="btn" data-acao="importar">Importar informe</button>
      </div>
    </section>
    <section class="secao">
      <h2>Dias lançados</h2>
      ${lista.length ? `<ul class="lista">${itens}</ul>` : `<div class="vazio">Nenhum dia lançado em ${MESES[mes - 1].toLowerCase()}.</div>`}
    </section>`;
}

function mudarMes(delta) {
  const d = new Date(ui.ano, ui.mes - 1 + delta, 1);
  ui.ano = d.getFullYear();
  ui.mes = d.getMonth() + 1;
  render();
}

// ---------- Editor de dia ----------
function abrirEditor(chave) {
  const existente = chave ? dados.registros[chave] : null;
  const c = dados.config;
  const padraoData = (() => {
    const hojeChave = chaveData(new Date());
    return hojeChave.startsWith(`${ui.ano}-${pad(ui.mes)}`) ? hojeChave : `${ui.ano}-${pad(ui.mes)}-01`;
  })();
  const r = existente || { data: padraoData, inicio: "", fim: "", jantar: c.jantarPadrao, taxa: c.taxaPadrao, centro: c.centroPadrao, obs: "" };
  const lista = [...new Set([...centros(), r.centro].filter(Boolean))];
  const opcoes = lista.length
    ? lista.map((x) => `<option ${x === r.centro ? "selected" : ""}>${esc(x)}</option>`).join("")
    : '<option value="">Escolha o modelo em Ajustes para ver a lista</option>';

  abrirDialogo(`<form class="folha" id="editor" novalidate>
      <h2>${existente ? `Editar ${curta(r.data)}` : "Adicionar dia"}</h2>
      <p class="sub">${existente ? SEMANA_LONGA[dataDe(r.data).getDay()] : "Lance ou corrija um dia de trabalho."}</p>
      <div class="erro" id="erro-editor" role="alert"></div>
      <label class="campo"><span>Data</span><input type="date" name="data" value="${r.data}" required></label>
      <div class="grade-2">
        <label class="campo"><span>Entrada</span><input type="time" name="inicio" value="${r.inicio}" required></label>
        <label class="campo"><span>Saída</span><input type="time" name="fim" value="${r.fim}"></label>
      </div>
      <div class="cartao linha-switch" style="margin-bottom:14px">
        <div>Jantei<small>Desconta 1h do dia</small></div>
        <label class="switch"><input type="checkbox" name="jantar" ${r.jantar ? "checked" : ""} aria-label="Jantei"><span></span></label>
      </div>
      <label class="campo"><span>Taxa por hora (R$)</span><input type="text" name="taxa" inputmode="decimal" value="${String(r.taxa ?? "").replace(".", ",")}"></label>
      <label class="campo"><span>Centro de custo</span><select name="centro">${opcoes}</select></label>
      <label class="campo"><span>Observação</span><input type="text" name="obs" value="${esc(r.obs)}" maxlength="120" placeholder="Opcional — vai para a planilha"></label>
      <div class="previa num" id="previa"></div>
      <div class="pilha">
        <button class="btn btn-primario" type="submit">Salvar</button>
        ${existente ? '<button class="btn btn-perigo" type="button" data-acao="excluir">Excluir dia</button>' : ""}
        <button class="btn" type="button" data-acao="fechar">Cancelar</button>
      </div>
    </form>`);

  const form = $("#editor");
  const lerForm = () => ({
    data: form.data.value, inicio: form.inicio.value, fim: form.fim.value, jantar: form.jantar.checked,
    taxa: Number(form.taxa.value.replace(/\./g, "").replace(",", ".")), centro: form.centro.value, obs: form.obs.value.trim(),
  });
  const atualizarPrevia = () => {
    const v = lerForm();
    $("#previa").innerHTML = v.inicio && v.fim
      ? `<span>${fmtHoras(minutosTrabalhados(v))}</span><span>${Number.isFinite(v.taxa) ? reais.format(valorDoDia(v)) : "—"}</span>`
      : `<span class="mudo">Preencha entrada e saída</span><span></span>`;
  };
  form.addEventListener("input", atualizarPrevia);
  atualizarPrevia();

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const v = lerForm();
    const erro = validar(v, existente);
    if (erro) { const el = $("#erro-editor"); el.textContent = erro; el.classList.add("visivel"); el.scrollIntoView({ block: "nearest" }); return; }
    if (existente && existente.data !== v.data) delete dados.registros[existente.data];
    dados.registros[v.data] = v;
    if (salvar()) { fecharDialogo(); avisar(`${curta(v.data)} salvo`); pedirPersistencia(); }
    [ui.ano, ui.mes] = [Number(v.data.slice(0, 4)), Number(v.data.slice(5, 7))];
    render();
  });
  form.querySelector('[data-acao="excluir"]')?.addEventListener("click", () => {
    if (!confirm(`Excluir o dia ${curta(existente.data)}?`)) return;
    delete dados.registros[existente.data];
    if (salvar()) { fecharDialogo(); avisar(`${curta(existente.data)} excluído`); }
    render();
  });
}

function validar(v, existente) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v.data)) return "Informe a data.";
  if (!v.inicio) return "Informe o horário de entrada.";
  if (v.fim && v.fim === v.inicio) return "Entrada e saída não podem ser iguais.";
  if (v.fim && v.jantar && minutosTrabalhados({ ...v, jantar: false }) <= 60) return "O turno tem 1h ou menos: desmarque \"Jantei\" ou corrija os horários.";
  if (!Number.isFinite(v.taxa) || v.taxa <= 0) return "Informe a taxa por hora (ex.: 50 ou 47,50).";
  if (dados.registros[v.data] && (!existente || existente.data !== v.data)) return `Já existe um registro em ${curta(v.data)}. Edite esse dia em vez de criar outro.`;
  return "";
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
  if (!lido.centros.includes(dados.config.centroPadrao)) { dados.config.centroPadrao = lido.centroMaisUsado; salvar(); }
  // Registros feitos antes do modelo ficam com o centro padrão.
  Object.values(dados.registros).forEach((r) => { if (!r.centro) r.centro = dados.config.centroPadrao; });
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
    const r = await gerarPlanilha(window.JSZip, deBase64(infoModelo.base64), {
      ano, mes, centroPadrao: dados.config.centroPadrao, registros: registrosDoMes(ano, mes),
    });
    const t = totaisDoMes(ano, mes);
    const arquivo = new File([r.arquivo], nomeArquivo(ano, mes, infoModelo.funcionario), { type: MIME_XLSM });
    abrirDialogo(`<div class="folha">
        <div class="icone-resultado ok"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
        <h2>Planilha gerada</h2>
        <p class="sub">${esc(arquivo.name)}</p>
        <div class="resumo num" style="margin-bottom:16px">
          <div><small>Dias</small><strong>${r.dias}</strong></div>
          <div><small>Horas</small><strong>${fmtHoras(Math.round(r.totalHoras * 60))}</strong></div>
          <div><small>Valor</small><strong>${reais.format(t.valor)}</strong></div>
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

// ---------- Importar informe ----------
function escolherArquivo(aceita) {
  return new Promise((resolve) => {
    const input = $("#arquivo");
    input.value = "";
    input.accept = aceita;
    input.onchange = () => resolve(input.files[0] || null);
    input.click();
  });
}

async function importarInforme() {
  const arq = await escolherArquivo(".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  if (!arq) return;
  let lido;
  try { lido = await lerInforme(window.JSZip, await arq.arrayBuffer()); }
  catch (e) { return mostrarErro("Informe inválido", mensagemDe(e)); }

  const { ano, mes, registros } = lido;
  const repetidos = registros.filter((r) => dados.registros[r.data]).length;
  abrirDialogo(`<div class="folha">
      <h2>Importar informe</h2>
      <p class="sub">${esc(arq.name)}</p>
      <div class="cartao" style="margin-bottom:16px">
        <strong>${registros.length} dias de ${MESES[mes - 1].toLowerCase()} de ${ano}</strong>
        <p class="mudo" style="margin:4px 0 0">${repetidos ? `${repetidos} dia(s) já lançados serão substituídos pelos horários do informe.` : "Nenhum dia desse mês será substituído."}</p>
      </div>
      <div class="pilha">
        <button class="btn btn-primario" data-acao="confirmar-importacao">Importar</button>
        <button class="btn" data-acao="fechar">Cancelar</button>
      </div>
    </div>`);
  $('[data-acao="confirmar-importacao"]').addEventListener("click", () => {
    const c = dados.config;
    for (const r of registros) {
      const atual = dados.registros[r.data];
      dados.registros[r.data] = {
        data: r.data, inicio: r.inicio, fim: r.fim, jantar: r.jantar,
        taxa: r.taxa ?? atual?.taxa ?? c.taxaPadrao, centro: atual?.centro ?? c.centroPadrao, obs: atual?.obs ?? "",
      };
    }
    if (salvar()) { fecharDialogo(); avisar(`${registros.length} dias importados`); pedirPersistencia(); }
    [ui.ano, ui.mes] = [ano, mes];
    render();
  });
}

// ---------- Tela: Ajustes ----------
function telaAjustes() {
  const c = dados.config;
  const ultimo = c.ultimoBackup ? new Date(c.ultimoBackup).toLocaleDateString("pt-BR") : "nunca";
  const opcoes = centros().length
    ? centros().map((x) => `<option ${x === c.centroPadrao ? "selected" : ""}>${esc(x)}</option>`).join("")
    : '<option value="">Escolha o modelo abaixo para ver a lista</option>';
  const modelo = infoModelo
    ? `<div class="cartao">
        <strong style="word-break:break-word">${esc(infoModelo.arquivo)}</strong>
        <p class="mudo" style="margin:4px 0 0;font-size:14px">Importado em ${new Date(infoModelo.importadoEm).toLocaleDateString("pt-BR")} · ${infoModelo.centros.length} centros de custo</p>
      </div>
      <button class="btn" style="margin-top:10px" data-acao="modelo">Trocar modelo</button>`
    : `${faixa("Nenhum modelo escolhido. Ele é necessário para gerar a planilha.", "modelo", "Escolher")}`;
  return `<header class="cabecalho"><h1>Ajustes</h1></header>
    <section class="secao" style="margin-top:0">
      <h2>Modelo da planilha</h2>
      ${modelo}
      <p class="campo-dica" style="margin:10px 0 0">Qualquer planilha COMAU já feita serve. Ela fica guardada só neste aparelho.</p>
    </section>
    <form id="ajustes" novalidate>
      <section class="secao">
        <h2>Novos registros</h2>
        <label class="campo"><span>Taxa por hora padrão (R$)</span><input type="text" name="taxa" inputmode="decimal" value="${String(c.taxaPadrao).replace(".", ",")}"></label>
        <p class="campo-dica">Usada nos próximos dias. Cada dia guarda a sua própria taxa e pode ser editado no Mês.</p>
        <label class="campo"><span>Centro de custo padrão</span><select name="centro">${opcoes}</select></label>
        <div class="cartao linha-switch">
          <div>Jantar marcado por padrão<small>Desconta 1h ao registrar a entrada</small></div>
          <label class="switch"><input type="checkbox" name="jantar" ${c.jantarPadrao ? "checked" : ""} aria-label="Jantar marcado por padrão"><span></span></label>
        </div>
        <div class="erro" id="erro-ajustes" role="alert" style="margin-top:14px"></div>
        <div class="pilha" style="margin-top:14px">
          <button class="btn btn-primario" type="submit">Salvar ajustes</button>
          <button class="btn" type="button" data-acao="aplicar-taxa">Aplicar taxa aos dias de ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}</button>
        </div>
      </section>
    </form>
    <section class="secao">
      <h2>Backup</h2>
      <div class="cartao">
        <p style="margin:0">Os registros ficam salvos só neste aparelho. Último backup: <strong>${ultimo}</strong>.</p>
        <p class="mudo" style="margin:6px 0 0;font-size:14px">Guarde o arquivo no app Arquivos ou no iCloud Drive.</p>
      </div>
      <div class="grade-2" style="margin-top:10px">
        <button class="btn" data-acao="backup">Fazer backup</button>
        <button class="btn" data-acao="restaurar">Restaurar</button>
      </div>
    </section>
    <p class="mudo" style="text-align:center;font-size:13px;margin-top:32px">Ponto COMAU · ${Object.keys(dados.registros).length} dias registrados</p>`;
}

function lerTaxa(texto) {
  const v = Number(String(texto).replace(/\./g, "").replace(",", "."));
  return Number.isFinite(v) && v > 0 ? v : null;
}

function salvarAjustes(form) {
  const taxa = lerTaxa(form.taxa.value);
  if (!taxa) { const el = $("#erro-ajustes"); el.textContent = "Informe a taxa por hora (ex.: 50 ou 47,50)."; el.classList.add("visivel"); return null; }
  Object.assign(dados.config, { taxaPadrao: taxa, centroPadrao: form.centro.value || dados.config.centroPadrao, jantarPadrao: form.jantar.checked });
  return salvar() ? taxa : null;
}

function aplicarTaxaAoMes() {
  const form = $("#ajustes");
  const taxa = salvarAjustes(form);
  if (!taxa) return;
  const lista = registrosDoMes(ui.ano, ui.mes);
  if (!lista.length) return mostrarErro("Nada para alterar", `Não há dias lançados em ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}.`);
  if (!confirm(`Aplicar ${reais.format(taxa)}/h aos ${lista.length} dias de ${MESES[ui.mes - 1].toLowerCase()}/${ui.ano}?`)) return;
  lista.forEach((r) => { r.taxa = taxa; });
  if (salvar()) avisar(`Taxa aplicada a ${lista.length} dias`);
  render();
}

async function fazerBackup() {
  // O backup leva também o modelo, para restaurar tudo num aparelho novo.
  const conteudo = JSON.stringify({ app: "ponto-comau", versao: 1, exportadoEm: new Date().toISOString(), ...dados, modelo: infoModelo }, null, 2);
  const nome = `backup-ponto-${chaveData(new Date())}.json`;
  const ok = await entregarArquivo(new File([conteudo], nome, { type: "application/json" }));
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
  avisoTimer = setTimeout(() => el.classList.remove("visivel"), 2400);
}

// ---------- Eventos ----------
document.addEventListener("click", (e) => {
  const alvo = e.target.closest("[data-acao], [data-aba]");
  if (!alvo) return;
  if (alvo.dataset.aba) return navegar(alvo.dataset.aba);
  const acoes = {
    entrada: registrarEntrada, saida: registrarSaida, editar: () => abrirEditor(alvo.dataset.data),
    novo: () => abrirEditor(null), "ir-mes": () => { const d = new Date(); [ui.ano, ui.mes] = [d.getFullYear(), d.getMonth() + 1]; navegar("mes"); },
    "mes-anterior": () => mudarMes(-1), "mes-seguinte": () => mudarMes(1), gerar, importar: importarInforme,
    backup: fazerBackup, restaurar: restaurarBackup, modelo: () => { fecharDialogo(); escolherModelo(); }, "aplicar-taxa": aplicarTaxaAoMes, fechar: fecharDialogo,
  };
  if (acoes[alvo.dataset.acao]) { e.preventDefault(); acoes[alvo.dataset.acao](); }
});

document.addEventListener("change", (e) => {
  if (e.target.matches('[data-acao="jantar"]')) {
    const aberto = turnoAberto();
    if (aberto) { aberto.jantar = e.target.checked; salvar(); }
  }
});

document.addEventListener("submit", (e) => {
  if (e.target.id !== "ajustes") return;
  e.preventDefault();
  if (salvarAjustes(e.target)) { avisar("Ajustes salvos"); render(); }
});

$("#dialogo").addEventListener("click", (e) => { if (e.target.id === "dialogo") fecharDialogo(); });
document.addEventListener("visibilitychange", () => { if (!document.hidden && ui.aba === "hoje") render(); });

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  navigator.serviceWorker.register("sw.js").catch(() => { /* app funciona sem offline */ });
}

navegar("hoje");
