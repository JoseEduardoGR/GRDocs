const PptxGenJS = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";

// ── Paleta ──────────────────────────────────────────────────────────────────
const C = {
  burgundy:     "6B0F1A",
  burgundyMid:  "8B1A2A",
  burgundyLight:"C94040",
  gold:         "D4AF37",
  goldLight:    "F0D060",
  cream:        "FAF6EE",
  creamDark:    "F0E8D8",
  charcoal:     "1C1C1C",
  darkBg:       "1A0A0D",
  textDark:     "2C1A1A",
  textMid:      "5C3A3A",
  white:        "FFFFFF",
  grayLight:    "F5F0EB",
  grayMid:      "D8CCC0",
  tableBand:    "FAF0E8",
};

// ── Helper: shadow fresco ────────────────────────────────────────────────────
const mkShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.18 });
const mkShadowSoft = () => ({ type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.10 });
const mkShadowGold = () => ({ type: "outer", blur: 10, offset: 3, angle: 135, color: C.gold, opacity: 0.22 });

// ── Helper: fondo oscuro ─────────────────────────────────────────────────────
function darkBg(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.darkBg }, line: { type: "none" } });
  // Textura sutil: faja lateral izquierda
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.gold }, line: { type: "none" } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 9.94, y: 0, w: 0.06, h: 5.625, fill: { color: C.gold }, line: { type: "none" } });
}

// ── Helper: fondo claro ──────────────────────────────────────────────────────
function lightBg(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.cream }, line: { type: "none" } });
  // Barra superior burdeos
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.18, fill: { color: C.burgundy }, line: { type: "none" } });
  // Barra inferior dorada
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.44, w: 10, h: 0.06, fill: { color: C.gold }, line: { type: "none" } });
}

// ── Helper: header de sección ────────────────────────────────────────────────
function sectionHeader(slide, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 0.22, w: 3.5, h: 0.06, fill: { color: C.gold }, line: { type: "none" } });
  slide.addText(title, {
    x: 0.45, y: 0.30, w: 9.1, h: 0.58,
    fontSize: 26, fontFace: "Georgia", bold: true, color: C.burgundy,
    valign: "middle",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.45, y: 0.88, w: 9.1, h: 0.32,
      fontSize: 12, fontFace: "Calibri", color: C.textMid, italic: true,
    });
  }
}

// ── Helper: footer ───────────────────────────────────────────────────────────
function footer(slide, text) {
  slide.addText(text || "Historia del Violín · De Stradivari a la Actualidad", {
    x: 0.45, y: 5.28, w: 9.1, h: 0.25,
    fontSize: 8, fontFace: "Calibri", color: C.textMid, align: "center",
  });
}

function footerDark(slide, text) {
  slide.addText(text || "Historia del Violín · De Stradivari a la Actualidad", {
    x: 0.45, y: 5.28, w: 9.1, h: 0.25,
    fontSize: 8, fontFace: "Calibri", color: C.grayMid, align: "center",
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — PORTADA
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  darkBg(slide);

  // Faja central decorativa
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.85, w: 10, h: 2.0, fill: { color: "2D0810" }, line: { type: "none" } });

  // Línea dorada superior de la faja
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.85, w: 10, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });
  // Línea dorada inferior de la faja
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 3.81, w: 10, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });

  // Ornamento central dorado (pequeño rombo)
  slide.addShape(pres.shapes.RECTANGLE, { x: 4.84, y: 1.76, w: 0.30, h: 0.30,
    fill: { color: C.gold }, line: { type: "none" },
    rotate: 45,
  });

  // Título principal
  slide.addText("Historia del Violín", {
    x: 0.5, y: 1.92, w: 9.0, h: 1.0,
    fontSize: 48, fontFace: "Georgia", bold: true, color: C.white,
    align: "center", valign: "middle",
    shadow: mkShadowGold(),
  });

  // Subtítulo
  slide.addText("De Stradivari a la Actualidad", {
    x: 0.5, y: 2.92, w: 9.0, h: 0.62,
    fontSize: 22, fontFace: "Georgia", color: C.gold,
    align: "center", valign: "middle", italic: true,
  });

  // Línea decorativa bajo subtítulo
  slide.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 3.62, w: 3.0, h: 0.03, fill: { color: C.grayMid }, line: { type: "none" } });

  // Tagline
  slide.addText('"El violín es el corazón del alma orquestal"', {
    x: 0.5, y: 3.72, w: 9.0, h: 0.38,
    fontSize: 12, fontFace: "Calibri", color: C.grayMid,
    align: "center", italic: true,
  });

  // Fecha / institución
  slide.addText("Presentación Académica · 2025", {
    x: 0.5, y: 5.20, w: 9.0, h: 0.28,
    fontSize: 10, fontFace: "Calibri", color: C.grayMid,
    align: "center",
  });
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — ÍNDICE
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "Contenido de la Presentación", "Un recorrido por los siglos de historia del violín");

  // Grid 2×4 de tarjetas de índice
  const items = [
    { num: "01", label: "Introducción — ¿Qué es el violín?" },
    { num: "02", label: "Orígenes históricos · Siglo XVI" },
    { num: "03", label: "La era dorada — Antonio Stradivari" },
    { num: "04", label: "Anatomía del violín" },
    { num: "05", label: "Técnicas de construcción tradicionales" },
    { num: "06", label: "Grandes violinistas de la historia" },
    { num: "07", label: "El violín en diferentes géneros" },
    { num: "08", label: "Violín moderno vs. barroco" },
  ];

  const cols = 2;
  const cardW = 4.20;
  const cardH = 0.62;
  const gapX = 0.30;
  const gapY = 0.18;
  const startX = 0.45;
  const startY = 1.30;

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    // Fondo tarjeta
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white }, line: { color: C.grayMid, pt: 1 },
      shadow: mkShadowSoft(),
    });
    // Acento izquierdo
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.07, h: cardH,
      fill: { color: C.burgundy }, line: { type: "none" },
    });
    // Número
    slide.addText(item.num, {
      x: x + 0.14, y: y + 0.10, w: 0.52, h: 0.42,
      fontSize: 18, fontFace: "Georgia", bold: true, color: C.gold,
      valign: "middle",
    });
    // Etiqueta
    slide.addText(item.label, {
      x: x + 0.68, y: y + 0.08, w: cardW - 0.80, h: 0.46,
      fontSize: 12, fontFace: "Calibri", color: C.textDark,
      valign: "middle",
    });
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — INTRODUCCIÓN
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "¿Qué es el Violín?", "El rey de los instrumentos de cuerda");

  // Columna izquierda: texto principal
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.28, w: 5.40, h: 3.80,
    fill: { color: C.white }, line: { color: C.grayMid, pt: 1 },
    shadow: mkShadowSoft(),
  });

  const introText = [
    { text: "El violín es un instrumento musical de cuerda frotada con arco, compuesto por cuatro cuerdas afinadas en quintas:", options: { fontSize: 12, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "Sol · Re · La · Mi", options: { fontSize: 13, fontFace: "Georgia", bold: true, color: C.burgundy, align: "center", breakLine: true } },
    { text: "Es el instrumento de tessitura más aguda dentro de la familia de las cuerdas de la orquesta sinfónica, conformada también por la viola, el violonchelo y el contrabajo.", options: { fontSize: 11, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "Sus dimensiones oscilan entre 59 y 60 cm de longitud total. Su caja de resonancia, cuidadosamente tallada en madera de arce y abeto, es la responsable de amplificar y dar color al sonido.", options: { fontSize: 11, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "Curiosidades clave:", options: { fontSize: 12, fontFace: "Calibri", bold: true, color: C.burgundy, breakLine: true } },
    { text: "Cuerdas afinadas en quintas perfectas", options: { bullet: true, fontSize: 11, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "Rango sonoro de más de 4 octavas", options: { bullet: true, fontSize: 11, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "Más de 70 piezas de madera en su construcción", options: { bullet: true, fontSize: 11, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "El arco usa crines de cola de caballo", options: { bullet: true, fontSize: 11, fontFace: "Calibri", color: C.textDark } },
  ];
  slide.addText(introText, {
    x: 0.62, y: 1.42, w: 5.10, h: 3.52,
    valign: "top", paraSpaceAfter: 4,
  });

  // Columna derecha: tarjetas de datos
  const facts = [
    { label: "Origen", value: "Italia, ~1550" },
    { label: "Familia", value: "Cuerda frotada" },
    { label: "Cuerdas", value: "4 (Sol·Re·La·Mi)" },
    { label: "Madera", value: "Abeto & Arce" },
    { label: "Peso aprox.", value: "400–500 g" },
  ];

  facts.forEach((f, i) => {
    const fy = 1.28 + i * 0.74;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 6.10, y: fy, w: 3.45, h: 0.62,
      fill: { color: i % 2 === 0 ? C.burgundy : C.burgundyMid },
      line: { type: "none" },
      shadow: mkShadow(),
    });
    slide.addText(f.label, {
      x: 6.22, y: fy + 0.04, w: 1.30, h: 0.54,
      fontSize: 10, fontFace: "Calibri", color: C.gold, bold: true, valign: "middle",
    });
    slide.addText(f.value, {
      x: 7.60, y: fy + 0.04, w: 1.80, h: 0.54,
      fontSize: 12, fontFace: "Georgia", color: C.white, bold: true, valign: "middle",
    });
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — ORÍGENES HISTÓRICOS
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "Orígenes Históricos", "El nacimiento del violín en el Renacimiento italiano · Siglo XVI");

  // Timeline horizontal
  const events = [
    { year: "~1520", event: "Primeras\nrebecas y\nvihuelas de\narco en Italia" },
    { year: "~1550", event: "Andrea\nAmati crea\nel primer\nviolín moderno" },
    { year: "1564", event: "Amati\nconstruye\nviolines para\nCarlos IX" },
    { year: "1590", event: "Gasparo\nda Salò\nperfecciona\nel diseño" },
    { year: "~1620", event: "Giovanni P.\nMaggini eleva\nla luthería\ncremanese" },
  ];

  // Línea de tiempo base
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 2.85, w: 8.90, h: 0.05,
    fill: { color: C.gold }, line: { type: "none" },
  });

  const totalW = 8.90;
  const segW = totalW / (events.length - 1);

  events.forEach((ev, i) => {
    const cx = 0.55 + i * segW;

    // Círculo (rectángulo cuadrado para simular punto)
    slide.addShape(pres.shapes.OVAL, {
      x: cx - 0.18, y: 2.68, w: 0.36, h: 0.36,
      fill: { color: C.burgundy }, line: { color: C.gold, pt: 2 },
      shadow: mkShadow(),
    });

    // Año
    slide.addText(ev.year, {
      x: cx - 0.55, y: 3.12, w: 1.10, h: 0.30,
      fontSize: 10, fontFace: "Georgia", bold: true, color: C.burgundy,
      align: "center",
    });

    // Descripción alternando arriba/abajo
    if (i % 2 === 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx - 0.70, y: 1.35, w: 1.40, h: 1.25,
        fill: { color: C.white }, line: { color: C.grayMid, pt: 1 },
        shadow: mkShadowSoft(),
      });
      slide.addText(ev.event, {
        x: cx - 0.65, y: 1.40, w: 1.30, h: 1.15,
        fontSize: 9.5, fontFace: "Calibri", color: C.textDark,
        align: "center", valign: "middle",
      });
      // Línea vertical conector
      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx - 0.01, y: 2.60, w: 0.02, h: 0.10,
        fill: { color: C.grayMid }, line: { type: "none" },
      });
    } else {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx - 0.70, y: 3.50, w: 1.40, h: 1.25,
        fill: { color: C.grayLight }, line: { color: C.grayMid, pt: 1 },
        shadow: mkShadowSoft(),
      });
      slide.addText(ev.event, {
        x: cx - 0.65, y: 3.55, w: 1.30, h: 1.15,
        fontSize: 9.5, fontFace: "Calibri", color: C.textDark,
        align: "center", valign: "middle",
      });
    }
  });

  // Nota al pie
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 4.88, w: 9.1, h: 0.38,
    fill: { color: C.creamDark }, line: { color: C.grayMid, pt: 1 },
  });
  slide.addText("El epicentro de la luthería: Cremona (Italia), ciudad que vio nacer las tres familias maestras — Amati, Guarneri y Stradivari.", {
    x: 0.55, y: 4.92, w: 8.90, h: 0.30,
    fontSize: 9.5, fontFace: "Calibri", italic: true, color: C.textMid, valign: "middle",
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — ANTONIO STRADIVARI
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();

  // Fondo degradado oscuro con textura
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "1A0A0D" }, line: { type: "none" } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.585, w: 10, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });

  // Panel izquierdo oscuro
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 4.20, h: 5.625, fill: { color: "2D0810" }, line: { type: "none" } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 4.14, y: 0, w: 0.06, h: 5.625, fill: { color: C.gold }, line: { type: "none" } });

  // Nombre y fechas en panel izquierdo
  slide.addText("ANTONIO", {
    x: 0.20, y: 0.55, w: 3.80, h: 0.60,
    fontSize: 32, fontFace: "Georgia", bold: true, color: C.gold,
    align: "center", shadow: mkShadowGold(),
  });
  slide.addText("STRADIVARI", {
    x: 0.20, y: 1.10, w: 3.80, h: 0.70,
    fontSize: 32, fontFace: "Georgia", bold: true, color: C.white,
    align: "center",
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.80, y: 1.82, w: 2.60, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });
  slide.addText("1644 – 1737", {
    x: 0.20, y: 1.90, w: 3.80, h: 0.40,
    fontSize: 16, fontFace: "Calibri", color: C.grayMid, align: "center", italic: true,
  });
  slide.addText("Cremona, Italia", {
    x: 0.20, y: 2.28, w: 3.80, h: 0.30,
    fontSize: 13, fontFace: "Calibri", color: C.grayMid, align: "center",
  });

  // Datos rápidos panel izquierdo
  const quickFacts = [
    { label: "Instrumentos creados", val: "~1,100" },
    { label: "Supervivientes hoy", val: "~650" },
    { label: "Violines", val: "~500" },
    { label: "Valor estimado", val: ">$10M c/u" },
  ];
  quickFacts.forEach((f, i) => {
    const qy = 2.72 + i * 0.60;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: qy, w: 3.70, h: 0.50, fill: { color: "3A0E18" }, line: { type: "none" } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: qy, w: 0.06, h: 0.50, fill: { color: C.gold }, line: { type: "none" } });
    slide.addText(f.label, { x: 0.40, y: qy + 0.06, w: 2.0, h: 0.38, fontSize: 9.5, fontFace: "Calibri", color: C.grayMid, valign: "middle" });
    slide.addText(f.val, { x: 2.50, y: qy + 0.06, w: 1.30, h: 0.38, fontSize: 13, fontFace: "Georgia", bold: true, color: C.gold, valign: "middle", align: "right" });
  });

  // Panel derecho: texto
  sectionHeaderDark(slide, "La Era Dorada del Violín");

  function sectionHeaderDark(sl, title) {
    sl.addShape(pres.shapes.RECTANGLE, { x: 4.40, y: 0.26, w: 2.80, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });
    sl.addText(title, { x: 4.35, y: 0.34, w: 5.45, h: 0.52, fontSize: 20, fontFace: "Georgia", bold: true, color: C.white, valign: "middle" });
  }

  const bodyText = [
    { text: "El periodo de máxima producción de Stradivari, conocido como su \"Periodo de Oro\" (1700–1720), marcó el culmen de la luthería occidental.", options: { fontSize: 11, fontFace: "Calibri", color: C.grayMid, breakLine: true } },
    { text: "Su secreto jamás revelado", options: { fontSize: 12, fontFace: "Georgia", bold: true, color: C.gold, breakLine: true } },
    { text: "El barniz de los Stradivarius — de tonalidad rojiza intensa — sigue siendo objeto de estudio científico. Hipótesis modernas sugieren un tratamiento especial de la madera con minerales o sales de potasio.", options: { fontSize: 11, fontFace: "Calibri", color: C.grayMid, breakLine: true } },
    { text: "El \"Mesías\" de 1716", options: { fontSize: 12, fontFace: "Georgia", bold: true, color: C.gold, breakLine: true } },
    { text: "Considerado el Stradivarius mejor conservado de la historia, es uno de los pocos que jamás ha sido tocado en público. Hoy reside en el Ashmolean Museum de Oxford.", options: { fontSize: 11, fontFace: "Calibri", color: C.grayMid, breakLine: true } },
    { text: "Legado vivo", options: { fontSize: 12, fontFace: "Georgia", bold: true, color: C.gold, breakLine: true } },
    { text: "Artistas como Itzhak Perlman, Hilary Hahn y Anne-Sophie Mutter han tocado Stradivarius. Ninguna tecnología moderna ha logrado replicar su timbre inigualable.", options: { fontSize: 11, fontFace: "Calibri", color: C.grayMid } },
  ];

  slide.addText(bodyText, {
    x: 4.35, y: 0.95, w: 5.45, h: 4.40,
    valign: "top", paraSpaceAfter: 5,
  });

  footerDark(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — ANATOMÍA DEL VIOLÍN
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "Anatomía del Violín", "Las partes principales y su función acústica");

  // Diagrama esquemático (representación con formas geométricas)
  // Cuerpo central del violín
  const vx = 1.80, vy = 1.20;

  // Caja de resonancia (forma aproximada con rectángulo redondeado)
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: vx + 0.30, y: vy + 0.55, w: 1.50, h: 2.80,
    fill: { color: C.burgundy }, line: { color: C.burgundyLight, pt: 1.5 },
    rectRadius: 0.18, shadow: mkShadow(),
  });
  // Efes (f-holes) simulados
  slide.addShape(pres.shapes.OVAL, { x: vx + 0.44, y: vy + 1.50, w: 0.12, h: 0.40, fill: { color: C.darkBg }, line: { type: "none" } });
  slide.addShape(pres.shapes.OVAL, { x: vx + 1.54, y: vy + 1.50, w: 0.12, h: 0.40, fill: { color: C.darkBg }, line: { type: "none" } });
  // Puente
  slide.addShape(pres.shapes.RECTANGLE, { x: vx + 0.80, y: vy + 1.72, w: 0.50, h: 0.08, fill: { color: C.gold }, line: { type: "none" } });
  // Clavijero (cuello + voluta)
  slide.addShape(pres.shapes.RECTANGLE, { x: vx + 0.76, y: vy + 0.08, w: 0.58, h: 0.52, fill: { color: C.burgundyMid }, line: { type: "none" } });
  slide.addShape(pres.shapes.OVAL, { x: vx + 0.84, y: vy + 0.00, w: 0.42, h: 0.22, fill: { color: C.burgundyLight }, line: { type: "none" } });
  // Clavijas (4 rectángulos pequeños)
  [0.70, 0.83, 0.96, 1.09].forEach(pos => {
    slide.addShape(pres.shapes.RECTANGLE, { x: vx + pos, y: vy + 0.10, w: 0.08, h: 0.24, fill: { color: C.gold }, line: { type: "none" } });
  });
  // Cuerdas (4 líneas)
  [0.82, 0.91, 1.00, 1.09].forEach(cx2 => {
    slide.addShape(pres.shapes.RECTANGLE, { x: vx + cx2 - 0.01, y: vy + 0.30, w: 0.02, h: 3.00, fill: { color: C.goldLight }, line: { type: "none" } });
  });
  // Cordal
  slide.addShape(pres.shapes.RECTANGLE, { x: vx + 0.76, y: vy + 3.18, w: 0.58, h: 0.20, fill: { color: C.burgundyMid }, line: { type: "none" } });
  // Botón
  slide.addShape(pres.shapes.OVAL, { x: vx + 0.96, y: vy + 3.36, w: 0.18, h: 0.18, fill: { color: C.burgundy }, line: { color: C.gold, pt: 1 } });

  // Etiquetas con líneas
  const labels = [
    { text: "Voluta / Clavijero", lx: 0.20, ly: vy + 0.06 },
    { text: "Diapasón (ébano)", lx: 0.20, ly: vy + 0.68 },
    { text: "Tapa armónica (abeto)", lx: 0.20, ly: vy + 1.28 },
    { text: "Efes (f-holes)", lx: 0.20, ly: vy + 1.60 },
    { text: "Puente (arce)", lx: 0.20, ly: vy + 1.95 },
    { text: "Cordal (ébano)", lx: 0.20, ly: vy + 2.60 },
    { text: "Botón", lx: 0.20, ly: vy + 3.05 },
  ];

  labels.forEach((lb, i) => {
    const lineColor = i % 2 === 0 ? C.burgundy : C.gold;
    slide.addShape(pres.shapes.RECTANGLE, { x: lb.lx + 1.18, y: lb.ly + 0.08, w: vx + 0.72 - lb.lx - 1.18, h: 0.02, fill: { color: lineColor }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { x: lb.lx + 1.18 + (vx + 0.72 - lb.lx - 1.18) - 0.08, y: lb.ly + 0.04, w: 0.10, h: 0.10, fill: { color: lineColor }, line: { type: "none" } });
    slide.addText(lb.text, {
      x: lb.lx, y: lb.ly, w: 1.20, h: 0.26,
      fontSize: 9, fontFace: "Calibri", color: C.textDark, align: "right", valign: "middle",
    });
  });

  // Panel derecho: descripción de materiales
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.20, y: 1.22, w: 4.35, h: 3.98, fill: { color: C.white }, line: { color: C.grayMid, pt: 1 }, shadow: mkShadowSoft() });

  slide.addText("Materiales & Función", {
    x: 5.35, y: 1.30, w: 4.05, h: 0.38,
    fontSize: 14, fontFace: "Georgia", bold: true, color: C.burgundy,
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: 1.66, w: 4.00, h: 0.03, fill: { color: C.gold }, line: { type: "none" } });

  const matText = [
    { text: "Tapa: Abeto europeo (Picea excelsa)", options: { bullet: true, fontSize: 10.5, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "Seleccionado por su densidad uniforme y elasticidad para la transmisión de vibraciones.", options: { fontSize: 9.5, fontFace: "Calibri", color: C.textMid, breakLine: true } },
    { text: "Fondo y aros: Arce ondulado (Acer platanoides)", options: { bullet: true, fontSize: 10.5, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "Su veta flameada aporta belleza visual y rigidez estructural.", options: { fontSize: 9.5, fontFace: "Calibri", color: C.textMid, breakLine: true } },
    { text: "Alma y barra armónica (internas)", options: { bullet: true, fontSize: 10.5, fontFace: "Calibri", color: C.textDark, breakLine: true } },
    { text: "El alma de abeto transmite las vibraciones del puente al fondo. La barra armónica refuerza la tapa bajo la cuerda Sol.", options: { fontSize: 9.5, fontFace: "Calibri", color: C.textMid, breakLine: true } },
    { text: "Barniz: El gran misterio", options: { bullet: true, fontSize: 10.5, fontFace: "Georgia", bold: true, color: C.burgundy, breakLine: true } },
    { text: "No solo protege la madera — afecta directamente el timbre y la resonancia. Los ingredientes exactos del barniz de Stradivari permanecen desconocidos.", options: { fontSize: 9.5, fontFace: "Calibri", color: C.textMid } },
  ];

  slide.addText(matText, {
    x: 5.35, y: 1.75, w: 4.05, h: 3.30,
    valign: "top", paraSpaceAfter: 3,
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — TÉCNICAS DE CONSTRUCCIÓN
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "Técnicas de Construcción Tradicionales", "Un proceso artesanal de más de 200 horas de trabajo");

  // Grid 2×3 de tarjetas de proceso
  const steps = [
    { num: "1", title: "Selección de la madera", desc: "La madera se seca al aire libre durante 10–20 años. Solo el 5% de los árboles cortados supera los estrictos criterios del luthier." },
    { num: "2", title: "Tallado de la tapa y fondo", desc: "Mediante gubias y cepillos especiales, el luthier excava la madera hasta lograr un grosor variable entre 2 y 5 mm según la zona." },
    { num: "3", title: "Doblado de los aros", desc: "Los aros de arce se calientan sobre un hierro curvo hasta alcanzar la forma característica de la cintura del violín sin romperse." },
    { num: "4", title: "Encolado y montaje", desc: "Se usa cola de piel de conejo. Esta cola, reversible con calor, permite desmontajes para reparaciones futuras sin dañar la madera." },
    { num: "5", title: "Instalación del alma y barra", desc: "El alma se coloca con un compás especial a través del f-hole. Su posición exacta altera dramáticamente el sonido del instrumento." },
    { num: "6", title: "Barnizado a mano", desc: "Hasta 30 capas de barniz, cada una pulida. El luthier utiliza aceites, resinas y pigmentos naturales en una fórmula personal e intransferible." },
  ];

  const cols = 3;
  const cW = 2.80;
  const cH = 1.62;
  const gX = 0.22;
  const gY = 0.22;
  const sX = 0.38;
  const sY = 1.32;

  steps.forEach((s, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = sX + col * (cW + gX);
    const y = sY + row * (cH + gY);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cW, h: cH,
      fill: { color: C.white }, line: { color: C.grayMid, pt: 1 },
      shadow: mkShadowSoft(),
    });
    // Número badge
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.12, y: y + 0.12, w: 0.42, h: 0.42,
      fill: { color: C.burgundy }, line: { type: "none" },
    });
    slide.addText(s.num, {
      x: x + 0.12, y: y + 0.12, w: 0.42, h: 0.42,
      fontSize: 13, fontFace: "Georgia", bold: true, color: C.white,
      align: "center", valign: "middle",
    });
    // Título
    slide.addText(s.title, {
      x: x + 0.60, y: y + 0.14, w: cW - 0.72, h: 0.38,
      fontSize: 10.5, fontFace: "Georgia", bold: true, color: C.burgundy, valign: "middle",
    });
    // Descripción
    slide.addText(s.desc, {
      x: x + 0.12, y: y + 0.58, w: cW - 0.24, h: 0.96,
      fontSize: 9.5, fontFace: "Calibri", color: C.textDark, valign: "top",
    });
  });

  // Nota al pie
  slide.addText("Un luthier experimentado tarda entre 200 y 300 horas en construir un solo violín de concierto.", {
    x: 0.45, y: 5.05, w: 9.10, h: 0.28,
    fontSize: 9.5, fontFace: "Calibri", italic: true, color: C.textMid, align: "center",
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — GRANDES VIOLINISTAS (TIMELINE)
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "Grandes Violinistas de la Historia", "Un legado de virtuosismo a través de los siglos");

  // Timeline vertical con dos columnas alternadas
  const violinists = [
    { period: "S. XVIII", name: "Arcangelo Corelli", years: "1653–1713", note: "Padre del violín barroco. Definió la técnica del arco y la ornamentación." },
    { period: "S. XVIII", name: "Giuseppe Tartini", years: "1692–1770", note: "Compuso el \"Trino del Diablo\", aún hoy considerado una de las piezas más difíciles." },
    { period: "S. XIX", name: "Niccolò Paganini", years: "1782–1840", note: "Técnica sobrehumana. Inventó nuevos recursos: pizzicato, armónicos, scordatura." },
    { period: "S. XIX", name: "Joseph Joachim", years: "1831–1907", note: "Intérprete de Brahms y Beethoven. Fundador de la escuela alemana moderna." },
    { period: "S. XX", name: "Jascha Heifetz", years: "1901–1987", note: "Considerado por muchos el mayor violinista del siglo XX. Técnica perfecta." },
    { period: "S. XX", name: "Itzhak Perlman", years: "1945–", note: "Maestro viviente. Combina virtuosismo con una musicalidad profundamente emotiva." },
  ];

  // Línea central vertical
  slide.addShape(pres.shapes.RECTANGLE, { x: 4.96, y: 1.22, w: 0.06, h: 3.80, fill: { color: C.gold }, line: { type: "none" } });

  const itemH = 0.72;
  const startY = 1.28;

  violinists.forEach((v, i) => {
    const isLeft = i % 2 === 0;
    const y = startY + i * (itemH + 0.10);

    // Punto en la línea
    slide.addShape(pres.shapes.OVAL, {
      x: 4.86, y: y + 0.24, w: 0.26, h: 0.26,
      fill: { color: C.burgundy }, line: { color: C.gold, pt: 1.5 },
    });

    // Período
    slide.addText(v.period, {
      x: isLeft ? 4.52 : 5.18, y: y + 0.26, w: 0.80, h: 0.22,
      fontSize: 8, fontFace: "Calibri", color: C.gold, bold: true,
      align: isLeft ? "right" : "left",
    });

    // Tarjeta
    const cx = isLeft ? 0.40 : 5.42;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y, w: 4.40, h: itemH,
      fill: { color: isLeft ? C.white : C.grayLight },
      line: { color: C.grayMid, pt: 1 },
      shadow: mkShadowSoft(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: isLeft ? cx : cx + 4.33, y, w: 0.07, h: itemH,
      fill: { color: C.burgundy }, line: { type: "none" },
    });
    slide.addText(v.name, {
      x: cx + 0.14, y: y + 0.06, w: 4.10, h: 0.28,
      fontSize: 11, fontFace: "Georgia", bold: true, color: C.burgundy,
    });
    slide.addText(v.years, {
      x: cx + 0.14, y: y + 0.30, w: 1.20, h: 0.20,
      fontSize: 9, fontFace: "Calibri", color: C.gold, italic: true,
    });
    slide.addText(v.note, {
      x: cx + 0.14, y: y + 0.48, w: 4.10, h: 0.22,
      fontSize: 8.5, fontFace: "Calibri", color: C.textMid,
    });
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — TABLA COMPARATIVA: GÉNEROS MUSICALES
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "El Violín en Diferentes Géneros Musicales", "Versatilidad que trasciende épocas y culturas");

  const tableRows = [
    // Header
    [
      { text: "Género",       options: { bold: true, color: C.white, fontFace: "Georgia",  fontSize: 11, fill: { color: C.burgundy } } },
      { text: "Rol",          options: { bold: true, color: C.white, fontFace: "Georgia",  fontSize: 11, fill: { color: C.burgundy } } },
      { text: "Técnica principal", options: { bold: true, color: C.white, fontFace: "Georgia", fontSize: 11, fill: { color: C.burgundy } } },
      { text: "Referente",    options: { bold: true, color: C.white, fontFace: "Georgia",  fontSize: 11, fill: { color: C.burgundy } } },
      { text: "Afinación",    options: { bold: true, color: C.white, fontFace: "Georgia",  fontSize: 11, fill: { color: C.burgundy } } },
    ],
    [
      { text: "Música Clásica",  options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Melodía principal / tutti", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Arco, vibrato, pizzicato", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Heifetz, Perlman", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Estándar",    options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
    ],
    [
      { text: "Música Barroca", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Basso continuo / solista", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Sin barbilla, arco barroco", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Corelli, Vivaldi",   options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Scordatura",  options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
    ],
    [
      { text: "Jazz",           options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Improvisación solista", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Swing, blues scales", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Stéphane Grappelli", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Estándar",    options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
    ],
    [
      { text: "Folk / Celtic",  options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Melodía / danza", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Drones, ornamentos, dobles cuerdas", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Violin \"fiddle\"",  options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "GDAE / alt.",  options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
    ],
    [
      { text: "Rock / Indie",   options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Textura / solista amplificado", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Amplificado, efectos, distorsión", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Yellowcard, Arcade Fire", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
      { text: "Estándar",    options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.white } } },
    ],
    [
      { text: "Flamenco",       options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Armonía y contramelodia", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Rasgueado arco, ornam. flamencos", options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Ara Malikian",   options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
      { text: "Estándar",    options: { fontFace: "Calibri", fontSize: 10, color: C.textDark, fill: { color: C.tableBand } } },
    ],
  ];

  slide.addTable(tableRows, {
    x: 0.40, y: 1.25, w: 9.20,
    colW: [1.40, 2.00, 2.30, 1.90, 1.00],
    rowH: 0.44,
    border: { pt: 1, color: C.grayMid },
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — VIOLÍN MODERNO VS. BARROCO
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "Violín Moderno vs. Violín Barroco", "Dos mundos sonoros, una misma alma de madera");

  // División central
  slide.addShape(pres.shapes.RECTANGLE, { x: 4.94, y: 1.22, w: 0.12, h: 3.90, fill: { color: C.gold }, line: { type: "none" } });

  // Headers de columnas
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.40, y: 1.22, w: 4.46, h: 0.52, fill: { color: C.burgundy }, line: { type: "none" }, shadow: mkShadow() });
  slide.addText("VIOLÍN MODERNO", { x: 0.40, y: 1.22, w: 4.46, h: 0.52, fontSize: 15, fontFace: "Georgia", bold: true, color: C.white, align: "center", valign: "middle" });

  slide.addShape(pres.shapes.RECTANGLE, { x: 5.14, y: 1.22, w: 4.46, h: 0.52, fill: { color: C.burgundyMid }, line: { type: "none" }, shadow: mkShadow() });
  slide.addText("VIOLÍN BARROCO", { x: 5.14, y: 1.22, w: 4.46, h: 0.52, fontSize: 15, fontFace: "Georgia", bold: true, color: C.gold, align: "center", valign: "middle" });

  // Comparaciones
  const comparisons = [
    { label: "Cuerdas", mod: "Metálicas o sintéticas (Dominant, Evah Pirazzi)", bar: "Tripa animal (gut) sin enrollar" },
    { label: "Arco", mod: "Arco de Tourte: palo recto, presión constante", bar: "Arco convexo, más ligero, articulación natural" },
    { label: "Barbilla", mod: "Sí — permite vibrato amplificado y presión", bar: "No — el violín reposa sobre el hombro" },
    { label: "Diapasón", mod: "Ébano, más largo e inclinado hacia el puente", bar: "Madera clara, más corto, sin inclinación" },
    { label: "Tensión", mod: "Mayor tensión: mayor volumen para salas grandes", bar: "Menor tensión: timbre más cálido y flexible" },
    { label: "Afinación", mod: "A=440 Hz (estándar moderno)", bar: "A=415 Hz (un semitono más bajo)" },
    { label: "Vibrato", mod: "Continuo, expresivo, integral a la técnica", bar: "Ornamental, selectivo, no sistemático" },
  ];

  const rowH = 0.48;
  const startY = 1.82;

  comparisons.forEach((c, i) => {
    const ry = startY + i * rowH;
    const bg = i % 2 === 0 ? C.white : C.grayLight;

    slide.addShape(pres.shapes.RECTANGLE, { x: 0.40, y: ry, w: 4.46, h: rowH, fill: { color: bg }, line: { color: C.grayMid, pt: 0.5 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.14, y: ry, w: 4.46, h: rowH, fill: { color: i % 2 === 0 ? C.creamDark : C.white }, line: { color: C.grayMid, pt: 0.5 } });

    // Etiqueta centrada sobre el separador dorado
    slide.addShape(pres.shapes.RECTANGLE, { x: 4.50, y: ry + 0.10, w: 1.00, h: 0.28, fill: { color: C.gold }, line: { type: "none" } });
    slide.addText(c.label, { x: 4.50, y: ry + 0.10, w: 1.00, h: 0.28, fontSize: 8, fontFace: "Calibri", bold: true, color: C.darkBg, align: "center", valign: "middle" });

    slide.addText(c.mod, { x: 0.50, y: ry + 0.07, w: 3.90, h: 0.34, fontSize: 9.5, fontFace: "Calibri", color: C.textDark, valign: "middle" });
    slide.addText(c.bar, { x: 5.24, y: ry + 0.07, w: 3.90, h: 0.34, fontSize: 9.5, fontFace: "Calibri", color: C.textDark, valign: "middle" });
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — CURIOSIDADES Y DATOS
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  lightBg(slide);
  sectionHeader(slide, "Curiosidades y Datos Fascinantes", "Lo que quizás no sabías sobre el rey de los instrumentos");

  const facts = [
    {
      icon: "$",
      title: "El violín más caro del mundo",
      desc: "El \"Lady Blunt\" de Stradivari (1721) se subastó en 2011 por 15,9 millones de dólares. Su dueña original fue la nieta de Lord Byron.",
    },
    {
      icon: "70",
      title: "Más de 70 piezas de madera",
      desc: "Un violín completo está compuesto por entre 70 y 85 piezas individuales de madera, ensambladas exclusivamente con cola de piel animal.",
    },
    {
      icon: "♩",
      title: "El Stradivarius que nunca fue tocado",
      desc: "El \"Mesías\" (1716) está en perfecto estado porque rara vez ha sido tocado. Tiene más de 300 años y sigue siendo casi nuevo.",
    },
    {
      icon: "AI",
      title: "IA vs. Stradivarius",
      desc: "En estudios de doble ciego, músicos y público entrenado no pudieron distinguir consistentemente un Stradivarius de un violín moderno de alta gama.",
    },
    {
      icon: "200",
      title: "200 horas de trabajo",
      desc: "Un luthier experto invierte entre 200 y 300 horas en construir un violín de concierto. Cada barniz se aplica en 20-30 capas sucesivas.",
    },
    {
      icon: "♭",
      title: "El diapasón bajó un semitono",
      desc: "Los violines barrocos se afinaban a A=415 Hz. La estandarización moderna en A=440 Hz ocurrió en 1939 en la Conferencia Internacional de Londres.",
    },
  ];

  const cols = 3;
  const fW = 2.82;
  const fH = 1.62;
  const gX = 0.20;
  const gY = 0.24;
  const sX = 0.42;
  const sY = 1.30;

  facts.forEach((f, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = sX + col * (fW + gX);
    const y = sY + row * (fH + gY);

    // Tarjeta
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: fW, h: fH,
      fill: { color: C.white }, line: { color: C.grayMid, pt: 1 },
      shadow: mkShadowSoft(),
    });

    // Banner superior de color
    const topColor = i % 2 === 0 ? C.burgundy : C.burgundyMid;
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: fW, h: 0.46, fill: { color: topColor }, line: { type: "none" } });

    // Ícono / símbolo
    slide.addText(f.icon, {
      x: x + 0.12, y: y + 0.06, w: 0.50, h: 0.34,
      fontSize: 16, fontFace: "Georgia", bold: true, color: C.gold,
      valign: "middle",
    });

    // Título en banner
    slide.addText(f.title, {
      x: x + 0.66, y: y + 0.06, w: fW - 0.78, h: 0.34,
      fontSize: 10, fontFace: "Georgia", bold: true, color: C.white,
      valign: "middle",
    });

    // Descripción
    slide.addText(f.desc, {
      x: x + 0.12, y: y + 0.52, w: fW - 0.24, h: 1.02,
      fontSize: 9.5, fontFace: "Calibri", color: C.textDark, valign: "top",
    });
  });

  footer(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 12 — CIERRE
// ════════════════════════════════════════════════════════════════════════════
(function () {
  const slide = pres.addSlide();
  darkBg(slide);

  // Panel central decorativo
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.80, w: 10, h: 2.10, fill: { color: "2D0810" }, line: { type: "none" } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.80, w: 10, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 3.86, w: 10, h: 0.04, fill: { color: C.gold }, line: { type: "none" } });

  // Ornamentos dorados
  slide.addShape(pres.shapes.OVAL, { x: 4.82, y: 1.70, w: 0.36, h: 0.36, fill: { color: C.gold }, line: { type: "none" } });
  slide.addShape(pres.shapes.OVAL, { x: 4.82, y: 3.84, w: 0.36, h: 0.36, fill: { color: C.gold }, line: { type: "none" } });

  // Título de cierre
  slide.addText("Gracias", {
    x: 0.5, y: 0.22, w: 9.0, h: 1.42,
    fontSize: 62, fontFace: "Georgia", bold: true, color: C.gold,
    align: "center", valign: "middle",
    shadow: mkShadowGold(),
  });

  // Mensaje central
  slide.addText("Historia del Violín", {
    x: 0.5, y: 1.92, w: 9.0, h: 0.58,
    fontSize: 24, fontFace: "Georgia", color: C.white,
    align: "center", valign: "middle", bold: true,
  });

  slide.addText("De Stradivari a la Actualidad", {
    x: 0.5, y: 2.48, w: 9.0, h: 0.40,
    fontSize: 16, fontFace: "Calibri", color: C.grayMid,
    align: "center", valign: "middle", italic: true,
  });

  slide.addText("2025", {
    x: 0.5, y: 2.88, w: 9.0, h: 0.32,
    fontSize: 13, fontFace: "Calibri", color: C.gold,
    align: "center",
  });

  // Cita final
  slide.addShape(pres.shapes.RECTANGLE, { x: 2.00, y: 4.04, w: 6.00, h: 0.88, fill: { color: "2D0810" }, line: { color: C.gold, pt: 0.8 } });
  slide.addText('"El violín habla donde las palabras callan."', {
    x: 2.10, y: 4.10, w: 5.80, h: 0.50,
    fontSize: 12, fontFace: "Georgia", italic: true, color: C.grayMid,
    align: "center", valign: "middle",
  });
  slide.addText("— Anónimo", {
    x: 2.10, y: 4.60, w: 5.80, h: 0.22,
    fontSize: 9.5, fontFace: "Calibri", color: C.gold,
    align: "right",
  });

  footerDark(slide);
})();

// ════════════════════════════════════════════════════════════════════════════
// GUARDAR ARCHIVO
// ════════════════════════════════════════════════════════════════════════════
const outputPath = path.join("pptx", "cache", "output", "7d8c0e0e-78a7-4150-b87f-75ee030a13b4.pptx");
const outputDir = path.dirname(outputPath);

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

pres.writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`✅ Archivo generado exitosamente: ${outputPath}`);
  })
  .catch((err) => {
    console.error("❌ Error al generar el archivo:", err);
    process.exit(1);
  });