// Desenha o comprovante do mês (todos os dias registrados) numa única imagem PNG.
// Gerado no próprio aparelho, sem enviar nada para a internet.

const L = 1080;          // largura da imagem (px)
const M = 64;            // margem lateral
const COR = {
  fundo: "#ffffff", texto: "#111111", mudo: "#625a5a", linha: "#ece8e8", listra: "#f8f6f6",
  destaque: "#c2410c", azul: "#1d4ed8", perigo: "#b3261e",
};
const FONTE = "Manrope, -apple-system, system-ui, sans-serif";
const ALTURA_LINHA = 76;
const ALTURA_OBS = 28;

const fonte = (peso, tam) => `${peso} ${tam}px ${FONTE}`;

function texto(ctx, t, x, y, { peso = 400, tam = 28, cor = COR.texto, alinhar = "left" } = {}) {
  ctx.font = fonte(peso, tam);
  ctx.fillStyle = cor;
  ctx.textAlign = alinhar;
  ctx.fillText(t, x, y);
  return ctx.measureText(t).width;
}

function cortar(ctx, t, larguraMax) {
  if (ctx.measureText(t).width <= larguraMax) return t;
  while (t.length && ctx.measureText(t + "…").width > larguraMax) t = t.slice(0, -1);
  return t + "…";
}

/**
 * dados = {
 *   funcionario, titulo ("Outubro de 2026"), geradoEm ("01/11/2026 às 10:52"),
 *   coluna ("EQUIPE" ou "COMMESSA"), totais: { dias, horas },
 *   linhas: [{ dia, semana, faixa, mais1, pausa, opcao, horas, obs, semFim }]
 * }
 * Com layout "periodos" (NGC): subtitulo ("NGC · cliente · local") e cada linha com
 * periodos: [manhã, tarde, noite] ("08:00–12:00" ou "—") no lugar de faixa, pausa e opcao.
 * Retorna um Blob PNG.
 */
export async function desenharComprovante(dados) {
  // Garante a fonte do app carregada antes de desenhar (funciona offline: a fonte vem do próprio app).
  try { await Promise.all([400, 600, 700].map((p) => document.fonts.load(fonte(p, 28)))); } catch { /* usa a fonte do sistema */ }

  const alturaLinhas = dados.linhas.reduce((s, l) => s + ALTURA_LINHA + (l.obs ? ALTURA_OBS : 0), 0);
  const desce = dados.subtitulo ? 44 : 0; // espaço para a linha do cliente
  const topoTabela = 470 + desce;
  const altura = topoTabela + 64 + alturaLinhas + 150;
  const canvas = document.createElement("canvas");
  canvas.width = L;
  canvas.height = altura;
  const ctx = canvas.getContext("2d");
  ctx.textBaseline = "alphabetic";
  ctx.fillStyle = COR.fundo;
  ctx.fillRect(0, 0, L, altura);

  // Cabeçalho
  ctx.letterSpacing = "3px";
  texto(ctx, "COMPROVANTE DE HORAS", M, 104, { peso: 700, tam: 24, cor: COR.destaque });
  ctx.letterSpacing = "0px";
  texto(ctx, dados.funcionario || "Registro de horas", M, 162, { peso: 700, tam: 44 });
  texto(ctx, dados.titulo, M, 212, { peso: 500, tam: 30, cor: COR.mudo });
  if (dados.subtitulo) {
    ctx.font = fonte(700, 26);
    texto(ctx, cortar(ctx, dados.subtitulo, L - 2 * M), M, 256, { peso: 700, tam: 26, cor: COR.azul });
  }

  // Totais
  const colunas = [["Dias", String(dados.totais.dias)], ["Horas", dados.totais.horas]];
  const xs = [M, M + 220];
  colunas.forEach(([rotulo, valor], i) => {
    texto(ctx, rotulo, xs[i], 290 + desce, { peso: 500, tam: 24, cor: COR.mudo });
    texto(ctx, valor, xs[i], 350 + desce, { peso: 400, tam: 48 });
  });
  ctx.fillStyle = COR.linha;
  ctx.fillRect(M, 396 + desce, L - 2 * M, 2);

  // Cabeçalho da tabela
  const periodos = dados.layout === "periodos";
  const X = { dia: M, faixa: M + 136, pausa: M + 396, opcao: M + 508, horas: L - M };
  const larguraOpcao = X.horas - 96 - X.opcao;
  // NGC: três colunas iguais entre o dia e o total.
  const larguraPeriodo = (X.horas - 120 - X.faixa) / 3;
  const XP = [0, 1, 2].map((i) => X.faixa + i * larguraPeriodo);
  const cabecalho = periodos
    ? [["DIA", X.dia, "left"], ["MANHÃ", XP[0], "left"], ["TARDE", XP[1], "left"], ["NOITE", XP[2], "left"], ["TOTAL", X.horas, "right"]]
    : [["DIA", X.dia, "left"], ["HORÁRIO", X.faixa, "left"], ["PAUSA", X.pausa, "left"], [dados.coluna, X.opcao, "left"], ["HORAS", X.horas, "right"]];
  ctx.letterSpacing = "2px";
  cabecalho.forEach(([t, x, a]) => texto(ctx, t, x, 452 + desce, { peso: 600, tam: 20, cor: COR.mudo, alinhar: a }));
  ctx.letterSpacing = "0px";

  const marcaMais1 = (x, base) => {
    ctx.strokeStyle = COR.azul; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.roundRect(x, base - 24, 44, 28, 7); ctx.stroke();
    texto(ctx, "+1", x + 22, base - 3, { peso: 700, tam: 18, cor: COR.azul, alinhar: "center" });
  };

  // Linhas
  let y = topoTabela;
  dados.linhas.forEach((l, i) => {
    const h = ALTURA_LINHA + (l.obs ? ALTURA_OBS : 0);
    if (i % 2 === 0) { ctx.fillStyle = COR.listra; ctx.fillRect(M - 16, y, L - 2 * M + 32, h); }
    const base = y + 49;
    texto(ctx, l.dia, X.dia, base, { peso: 700, tam: 32 });
    texto(ctx, l.semana, X.dia + 50, base, { peso: 500, tam: 22, cor: COR.mudo });
    if (periodos) {
      l.periodos.forEach((t, k) => {
        const vazio = t === "—";
        const w = texto(ctx, t, XP[k], base, { peso: vazio ? 500 : 600, tam: 26, cor: vazio ? COR.mudo : COR.texto });
        if (k === 2 && l.mais1 && !vazio) marcaMais1(XP[k] + w + 8, base);
      });
    } else {
      const w = texto(ctx, l.faixa, X.faixa, base, { peso: 600, tam: 30 });
      if (l.mais1) marcaMais1(X.faixa + w + 12, base);
      texto(ctx, l.pausa, X.pausa, base, { peso: 500, tam: 24, cor: COR.mudo });
      ctx.font = fonte(500, 22);
      texto(ctx, cortar(ctx, l.opcao, larguraOpcao), X.opcao, base, { peso: 500, tam: 22 });
    }
    if (l.semFim) texto(ctx, "sem fim", X.horas, base, { peso: 600, tam: 26, cor: COR.perigo, alinhar: "right" });
    else texto(ctx, l.horas, X.horas, base, { peso: 600, tam: 30, alinhar: "right" });
    if (l.obs) {
      ctx.font = fonte(500, 22);
      texto(ctx, cortar(ctx, l.obs, L - 2 * M - 136), X.faixa, base + 32, { peso: 500, tam: 22, cor: COR.mudo });
    }
    y += h;
  });

  // Rodapé
  ctx.fillStyle = COR.linha;
  ctx.fillRect(M, y + 24, L - 2 * M, 2);
  texto(ctx, `Gerado em ${dados.geradoEm} pelo app Ponto Automisa`, M, y + 84, { peso: 500, tam: 22, cor: COR.mudo });
  texto(ctx, `${dados.linhas.length} registros`, L - M, y + 84, { peso: 500, tam: 22, cor: COR.mudo, alinhar: "right" });

  return new Promise((resolve, reject) =>
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("Não foi possível criar a imagem do comprovante."))), "image/png"));
}
