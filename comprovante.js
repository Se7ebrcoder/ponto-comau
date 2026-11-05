// Desenha o comprovante do mês (todos os dias registrados) numa única imagem PNG.
// Gerado no próprio aparelho, sem enviar nada para a internet.

const L = 1080;          // largura da imagem (px)
const M = 64;            // margem lateral
const COR = {
  fundo: "#ffffff", texto: "#111111", mudo: "#625a5a", linha: "#ece8e8", listra: "#f8f6f6",
  destaque: "#c8102e", dinheiro: "#1d7a3e", perigo: "#b3261e",
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
 *   totais: { dias, horas, valor }, linhas: [{ dia, semana, faixa, mais1, pausa, horas, valor, obs, semFim }]
 * }
 * Retorna um Blob PNG.
 */
export async function desenharComprovante(dados) {
  // Garante a fonte do app carregada antes de desenhar (funciona offline: a fonte vem do próprio app).
  try { await Promise.all([400, 600, 700].map((p) => document.fonts.load(fonte(p, 28)))); } catch { /* usa a fonte do sistema */ }

  const alturaLinhas = dados.linhas.reduce((s, l) => s + ALTURA_LINHA + (l.obs ? ALTURA_OBS : 0), 0);
  const topoTabela = 470;
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

  // Totais
  const colunas = [["Dias", String(dados.totais.dias)], ["Horas", dados.totais.horas], ["Total", dados.totais.valor]];
  const xs = [M, M + 220, M + 500];
  colunas.forEach(([rotulo, valor], i) => {
    texto(ctx, rotulo, xs[i], 290, { peso: 500, tam: 24, cor: COR.mudo });
    texto(ctx, valor, xs[i], 350, { peso: 400, tam: 48, cor: i === 2 ? COR.dinheiro : COR.texto });
  });
  ctx.fillStyle = COR.linha;
  ctx.fillRect(M, 396, L - 2 * M, 2);

  // Cabeçalho da tabela
  const X = { dia: M, faixa: M + 136, pausa: M + 500, horas: M + 726, valor: L - M };
  ctx.letterSpacing = "2px";
  [["DIA", X.dia, "left"], ["HORÁRIO", X.faixa, "left"], ["PAUSA", X.pausa, "left"], ["HORAS", X.horas, "right"], ["VALOR", X.valor, "right"]]
    .forEach(([t, x, a]) => texto(ctx, t, x, 452, { peso: 600, tam: 20, cor: COR.mudo, alinhar: a }));
  ctx.letterSpacing = "0px";

  // Linhas
  let y = topoTabela;
  dados.linhas.forEach((l, i) => {
    const h = ALTURA_LINHA + (l.obs ? ALTURA_OBS : 0);
    if (i % 2 === 0) { ctx.fillStyle = COR.listra; ctx.fillRect(M - 16, y, L - 2 * M + 32, h); }
    const base = y + 49;
    texto(ctx, l.dia, X.dia, base, { peso: 700, tam: 32 });
    texto(ctx, l.semana, X.dia + 50, base, { peso: 500, tam: 22, cor: COR.mudo });
    const w = texto(ctx, l.faixa, X.faixa, base, { peso: 600, tam: 30 });
    if (l.mais1) {
      const cx = X.faixa + w + 12;
      ctx.strokeStyle = COR.destaque; ctx.lineWidth = 2;
      ctx.beginPath(); ctx.roundRect(cx, base - 24, 44, 28, 7); ctx.stroke();
      texto(ctx, "+1", cx + 22, base - 3, { peso: 700, tam: 18, cor: COR.destaque, alinhar: "center" });
    }
    texto(ctx, l.pausa, X.pausa, base, { peso: 500, tam: 26, cor: COR.mudo });
    if (l.semFim) {
      texto(ctx, "sem fim", X.valor, base, { peso: 600, tam: 26, cor: COR.perigo, alinhar: "right" });
    } else {
      texto(ctx, l.horas, X.horas, base, { peso: 600, tam: 30, alinhar: "right" });
      texto(ctx, l.valor, X.valor, base, { peso: 600, tam: 28, cor: COR.dinheiro, alinhar: "right" });
    }
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
