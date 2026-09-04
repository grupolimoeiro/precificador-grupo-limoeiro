const pptxgen = require("pptxgenjs");
const assets = require("./assets-data.json");

// ─────────────────────────────────────────────────────────────
// Layout constants (standard 13.333 x 7.5in widescreen canvas —
// matches the source .pptx templates scaled 1:2)
// ─────────────────────────────────────────────────────────────
const GREEN = "C3E535";
const BLACK = "000000";
const GRAY = "555555";
const FONT = "Helvetica Neue";
const FONT_BOLD = "Helvetica Neue";

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const TITLE_X = 0.55, TITLE_Y = 0.25, TITLE_W = 12.2, TITLE_H = 1.15;
const CAT_X = 1.25;
const ITEM_ICON_X = 1.25, ITEM_TEXT_X = 1.56, ITEM_TEXT_W = 11.2;
const CONTENT_TOP = 1.75, CONTENT_BOTTOM = 6.85;
const FOOTER_Y = 7.0;

function fmtMoney(n) {
  return "R$ " + Math.round(n).toLocaleString("pt-BR");
}

// Estimates vertical height (inches) a single item (name + optional note) needs.
// Short notes render inline after the name (single line); long combined text wraps.
function itemHeight(item) {
  const combined = (item.name || "").length + (item.note ? item.note.length + 2 : 0);
  const CHARS_PER_LINE = 108;
  if (combined <= CHARS_PER_LINE) return 0.42;
  const extraLines = Math.ceil((combined - CHARS_PER_LINE) / CHARS_PER_LINE);
  return 0.42 + extraLines * 0.24;
}

function categoryHeaderHeight() {
  return 0.42;
}

function addBackground(slide, key) {
  slide.addImage({ data: "data:image/png;base64," + assets[key], x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
}

function addContentTitle(slide, text) {
  slide.addText(text, {
    x: TITLE_X, y: TITLE_Y, w: TITLE_W, h: TITLE_H,
    fontFace: FONT_BOLD, fontSize: 30, bold: true, color: BLACK,
    align: "left", valign: "top",
  });
}

// Renders category blocks (each: {title, items:[{name, note}]}) across as many
// slides as needed, auto-paginating like the original hand-built decks.
function renderCategorySlides(pptx, planTitle, categories) {
  let slide = null;
  let y = CONTENT_TOP;

  function newSlide() {
    slide = pptx.addSlide();
    addBackground(slide, "content_bg");
    addContentTitle(slide, planTitle);
    y = CONTENT_TOP;
  }

  newSlide();

  categories.forEach((cat) => {
    if (!cat.items.length) return;

    const catHeight =
      categoryHeaderHeight() + cat.items.reduce((s, it) => s + itemHeight(it), 0);

    if (y + catHeight > CONTENT_BOTTOM && y > CONTENT_TOP) {
      newSlide();
    }

    slide.addText(cat.title, {
      x: CAT_X, y, w: TITLE_W - (CAT_X - TITLE_X), h: 0.45,
      fontFace: FONT_BOLD, fontSize: 16, bold: true, color: BLACK,
    });
    y += categoryHeaderHeight();

    cat.items.forEach((item) => {
      const h = itemHeight(item);

      // if a single item doesn't fit remaining space, push to a new slide
      if (y + h > CONTENT_BOTTOM && y > CONTENT_TOP) {
        newSlide();
        slide.addText(cat.title + " (cont.)", {
          x: CAT_X, y, w: TITLE_W - (CAT_X - TITLE_X), h: 0.45,
          fontFace: FONT_BOLD, fontSize: 16, bold: true, color: BLACK,
        });
        y += categoryHeaderHeight();
      }

      slide.addImage({
        data: "data:image/png;base64," + assets.check,
        x: ITEM_ICON_X, y: y + 0.02, w: 0.22, h: 0.22,
      });

      const runs = [{ text: item.name, options: { bold: true, color: BLACK } }];
      if (item.note) {
        runs.push({ text: "  " + item.note, options: { bold: false, color: GRAY } });
      }
      slide.addText(runs, {
        x: ITEM_TEXT_X, y, w: ITEM_TEXT_W, h,
        fontFace: FONT, fontSize: 12, valign: "top", align: "left",
      });

      y += h;
    });

    y += 0.22;
  });
}

function renderPriceSlide(pptx, data) {
  const slide = pptx.addSlide();
  addBackground(slide, "price_bg");

  const rightX = 8.1, rightW = 4.3;

  slide.addText(`Valor válido no prazo de ${data.validity || "15 dias"}.`, {
    x: rightX, y: 0.75, w: rightW, h: 0.7,
    fontFace: FONT, fontSize: 13, color: BLACK,
  });

  slide.addText(`Plano ${data.planName}`, {
    x: rightX, y: 2.05, w: rightW, h: 0.35,
    fontFace: FONT, fontSize: 12, color: BLACK,
  });
  slide.addText("Valor Semestral", {
    x: rightX, y: 2.35, w: rightW, h: 0.45,
    fontFace: FONT_BOLD, fontSize: 18, bold: true, color: BLACK,
  });

  slide.addText("6x", {
    x: rightX, y: 3.05, w: rightW, h: 0.4,
    fontFace: FONT_BOLD, fontSize: 14, bold: true, color: BLACK,
  });
  slide.addText(fmtMoney(data.totalMensal), {
    x: rightX, y: 3.3, w: rightW, h: 1.05,
    fontFace: FONT_BOLD, fontSize: 54, bold: true, color: BLACK,
  });
  slide.addText(`${fmtMoney(data.totalSemestral)} em seis meses`, {
    x: rightX, y: 4.25, w: rightW, h: 0.4,
    fontFace: FONT_BOLD, fontSize: 15, bold: true, color: BLACK,
  });

  if (data.totalAddons) {
    slide.addText(`+ ${fmtMoney(data.totalAddons)} em projetos avulsos`, {
      x: rightX, y: 4.75, w: rightW, h: 0.4,
      fontFace: FONT, fontSize: 12, color: BLACK,
    });
  }

  slide.addText(
    [
      { text: "Captações e ações fora de Franca e do horário comercial terão custos adicionais.\n", options: {} },
      { text: "(Verificar disponibilidade e custos)", options: {} },
    ],
    {
      x: rightX, y: 6.3, w: rightW, h: 1.0,
      fontFace: FONT, fontSize: 11, color: GRAY,
    }
  );
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  let data;
  try {
    data = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
  } catch (e) {
    res.status(400).json({ error: "JSON inválido" });
    return;
  }

  try {
    const pptx = new pptxgen();
    pptx.defineLayout({ name: "LIMOEIRO", width: SLIDE_W, height: SLIDE_H });
    pptx.layout = "LIMOEIRO";

    // 1) capa + slides institucionais (estáticos, iguais em todos os planos)
    ["cover", "inst2", "inst3", "inst4", "inst5"].forEach((key) => {
      const slide = pptx.addSlide();
      addBackground(slide, key);
    });

    // 2) slide divisor do plano — arte original do template (não recriada),
    // uma imagem por plano (Start/Gestão/Supra) extraída dos .pptx enviados.
    const dividerKey = "divider_" + (data.planKey || "start");
    const divider = pptx.addSlide();
    addBackground(divider, assets[dividerKey] ? dividerKey : "divider_start");

    // 3) slides de conteúdo (categorias + itens personalizados)
    const planTitle = `Plano ${data.planName || ""} Personalizado`;
    renderCategorySlides(pptx, planTitle, data.categories || []);

    if (data.addons && data.addons.length) {
      renderCategorySlides(pptx, `${planTitle} — Avulsos`, [
        { title: "Projetos Avulsos", items: data.addons },
      ]);
    }

    // 4) slide de investimento
    renderPriceSlide(pptx, data);

    // 5) slide de encerramento (estático)
    const closing = pptx.addSlide();
    addBackground(closing, "closing");

    const buffer = await pptx.write({ outputType: "nodebuffer" });

    const safeName = (data.cliente || "Cliente").replace(/[\\/:*?"<>|]/g, "").trim() || "Cliente";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="Grupo Limoeiro - ${safeName}.pptx"`);
    res.status(200).send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Erro ao gerar o arquivo.", message: String(err && err.message || err) });
  }
};
