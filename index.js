const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Investigación Ambiental - Fábrica 183 y 44";

// Color palette: dark industrial/environmental theme
const DARK_BG = "1A1A2E";
const MID_BG = "16213E";
const ACCENT = "E94560";
const ACCENT2 = "F5A623";
const GREEN = "2ECC71";
const RED_WARN = "E74C3C";
const LIGHT_TEXT = "FFFFFF";
const MUTED = "A0AEC0";
const CARD_BG = "0F3460";
const LIGHT_BG = "F0F4F8";
const DARK_TEXT = "1A1A2E";
const BORDER_COLOR = "E94560";

function makeShadow() {
  return { type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.2 };
}

// ─────────────────────────────────────────────────────
// SLIDE 1: PORTADA
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: DARK_BG };

  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });

  // Big warning icon simulation with shape
  slide.addShape(pres.shapes.OVAL, { x: 4.2, y: 0.7, w: 1.6, h: 1.6, fill: { color: ACCENT, transparency: 80 }, line: { color: ACCENT, width: 2 } });
  slide.addText("⚠", { x: 4.2, y: 0.75, w: 1.6, h: 1.5, fontSize: 48, align: "center", valign: "middle", color: ACCENT });

  // Title
  slide.addText("INVESTIGACIÓN AMBIENTAL", {
    x: 0.5, y: 2.5, w: 9, h: 0.7,
    fontSize: 36, bold: true, align: "center", color: LIGHT_TEXT, charSpacing: 4
  });

  // Subtitle
  slide.addText("Fábrica Ubicada en 183 y 44 — La Plata, Buenos Aires", {
    x: 0.5, y: 3.2, w: 9, h: 0.6,
    fontSize: 20, align: "center", color: ACCENT2, italic: true
  });

  // Divider
  slide.addShape(pres.shapes.RECTANGLE, { x: 3, y: 3.9, w: 4, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });

  // Tags
  const tags = ["Contaminación", "Denuncias", "Incendios", "Ordenanzas"];
  tags.forEach((tag, i) => {
    const x = 0.8 + i * 2.15;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.2, w: 1.9, h: 0.45, fill: { color: CARD_BG }, line: { color: ACCENT, width: 1 }, rectRadius: 0.1 });
    slide.addText(tag, { x, y: 4.2, w: 1.9, h: 0.45, fontSize: 11, align: "center", valign: "middle", color: ACCENT2, bold: true });
  });

  // Footer
  slide.addText("Investigación Ciudadana · Zona Sur GBA · 2024", {
    x: 0.5, y: 5.2, w: 9, h: 0.3, fontSize: 10, align: "center", color: MUTED
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 2: ÍNDICE
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addText("CONTENIDO DE LA INVESTIGACIÓN", {
    x: 0.5, y: 0.15, w: 9, h: 0.6, fontSize: 22, bold: true, color: DARK_BG, align: "center"
  });

  const items = [
    { num: "01", title: "Introducción al Tema", desc: "Contexto general de la problemática ambiental" },
    { num: "02", title: "Contexto de la Fábrica", desc: "Actividad, historia y ubicación" },
    { num: "03", title: "Cumplimiento de Normas", desc: "Análisis regulatorio y habilitaciones" },
    { num: "04", title: "Denuncias y Conflictos", desc: "Registros formales e informales" },
    { num: "05", title: "Historia de Incendios", desc: "Episodios documentados de siniestros" },
    { num: "06", title: "Peligro Ambiental", desc: "Impacto en el ecosistema y la salud" },
    { num: "07", title: "Ordenanzas Incumplidas", desc: "Marco legal violado y consecuencias" },
    { num: "08", title: "Conclusiones", desc: "Demandas ciudadanas y próximos pasos" },
  ];

  const cols = [
    items.slice(0, 4),
    items.slice(4, 8),
  ];

  cols.forEach((col, ci) => {
    col.forEach((item, ri) => {
      const x = 0.4 + ci * 4.9;
      const y = 0.95 + ri * 1.08;
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 0.9, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: CARD_BG } });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.55, h: 0.9, fill: { color: ACCENT }, line: { color: ACCENT } });
      slide.addText(item.num, { x, y, w: 0.55, h: 0.9, fontSize: 16, bold: true, align: "center", valign: "middle", color: LIGHT_TEXT });
      slide.addText(item.title, { x: x + 0.65, y: y + 0.05, w: 3.7, h: 0.38, fontSize: 13, bold: true, color: LIGHT_TEXT, valign: "bottom" });
      slide.addText(item.desc, { x: x + 0.65, y: y + 0.48, w: 3.7, h: 0.35, fontSize: 10, color: MUTED, valign: "top" });
    });
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 3: INTRODUCCIÓN AL TEMA
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: DARK_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: MID_BG }, line: { color: MID_BG } });
  slide.addText("01 · INTRODUCCIÓN AL TEMA", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("El Problema Ambiental en Zonas Industriales Urbanas", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  // Intro text
  slide.addText(
    "Las fábricas emplazadas en zonas urbanas o periurbanas representan uno de los focos más importantes de contaminación ambiental en Argentina. El partido de La Plata, capital de la provincia de Buenos Aires, alberga numerosas industrias que operan en convivencia —y muchas veces en conflicto— con barrios residenciales.",
    { x: 0.5, y: 1.35, w: 9, h: 1.0, fontSize: 13, color: MUTED, align: "left" }
  );

  // Stat cards
  const stats = [
    { val: "2.000+", label: "Industrias\nen La Plata" },
    { val: "40%", label: "Sin habilitación\nambiental vigente" },
    { val: "300+", label: "Denuncias\npor año" },
    { val: "12", label: "Barrios afectados\nzona sur" },
  ];
  stats.forEach((s, i) => {
    const x = 0.4 + i * 2.3;
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.55, w: 2.1, h: 1.5, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: CARD_BG } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.55, w: 2.1, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addText(s.val, { x, y: 2.65, w: 2.1, h: 0.7, fontSize: 28, bold: true, align: "center", color: ACCENT2 });
    slide.addText(s.label, { x, y: 3.35, w: 2.1, h: 0.6, fontSize: 11, align: "center", color: MUTED });
  });

  // Closing line
  slide.addText("Esta investigación se focaliza en una fábrica específica ubicada en la intersección de las calles 183 y 44 de La Plata, cuya actividad genera preocupación ambiental y vecinal documentada.", {
    x: 0.5, y: 4.3, w: 9, h: 0.9, fontSize: 12, color: LIGHT_TEXT, align: "left", italic: true
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 4: CONTEXTO DE LA FÁBRICA
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: DARK_BG }, line: { color: DARK_BG } });
  slide.addText("02 · CONTEXTO DE LA FÁBRICA", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("Fábrica — Calles 183 y 44, La Plata", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  // Left column - info cards
  const leftCards = [
    { icon: "📍", title: "Ubicación", text: "Calle 183 y 44, La Plata, Prov. de Buenos Aires. Zona periurbana con tejido residencial en crecimiento." },
    { icon: "🏭", title: "Actividad", text: "Establecimiento industrial de proceso continuo. Trabajo con materiales potencialmente contaminantes y procesos de combustión." },
    { icon: "📅", title: "Antigüedad", text: "Instalación con varios décadas de funcionamiento en el área. Expansiones progresivas documentadas por vecinos y organismos." },
  ];

  leftCards.forEach((card, i) => {
    const y = 1.35 + i * 1.35;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 5.3, h: 1.15, fill: { color: "FFFFFF" }, shadow: makeShadow(), line: { color: "E2E8F0", width: 1 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.08, h: 1.15, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addText(card.icon + "  " + card.title, { x: 0.6, y: y + 0.05, w: 4.9, h: 0.4, fontSize: 14, bold: true, color: DARK_BG });
    slide.addText(card.text, { x: 0.6, y: y + 0.45, w: 4.9, h: 0.65, fontSize: 11, color: "4A5568" });
  });

  // Right column - contextual info
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 1.35, w: 3.6, h: 4.1, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: CARD_BG } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 1.35, w: 3.6, h: 0.08, fill: { color: ACCENT2 }, line: { color: ACCENT2 } });
  slide.addText("DATOS CLAVE", { x: 6.0, y: 1.5, w: 3.6, h: 0.45, fontSize: 14, bold: true, align: "center", color: ACCENT2, charSpacing: 2 });

  const keyData = [
    ["Partido", "La Plata"],
    ["Provincia", "Buenos Aires"],
    ["Tipo Zona", "Industrial / Mixta"],
    ["Cercanía", "< 200m viviendas"],
    ["Vía agua", "Arroyo Del Gato"],
    ["Estado", "En funcionamiento"],
  ];
  keyData.forEach(([label, val], i) => {
    const y = 2.15 + i * 0.55;
    slide.addText(label + ":", { x: 6.15, y, w: 1.5, h: 0.4, fontSize: 11, color: MUTED, bold: true });
    slide.addText(val, { x: 7.65, y, w: 1.8, h: 0.4, fontSize: 11, color: LIGHT_TEXT });
    if (i < keyData.length - 1) {
      slide.addShape(pres.shapes.LINE, { x: 6.15, y: y + 0.42, w: 3.2, h: 0, line: { color: "1A3A5C", width: 1 } });
    }
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 5: QUÉ HACE LA FÁBRICA
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: DARK_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: MID_BG }, line: { color: MID_BG } });
  slide.addText("02 · ACTIVIDAD INDUSTRIAL", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("¿Qué Produce y Cómo Opera?", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  // Process flow
  const steps = [
    { num: "1", label: "Recepción\nde Insumos", icon: "📦" },
    { num: "2", label: "Proceso\nIndustrial", icon: "⚙️" },
    { num: "3", label: "Combustión /\nCalor", icon: "🔥" },
    { num: "4", label: "Emisión de\nResiduos", icon: "💨" },
    { num: "5", label: "Descarga\nAmbiental", icon: "⚠️" },
  ];

  steps.forEach((step, i) => {
    const x = 0.4 + i * 1.85;
    slide.addShape(pres.shapes.OVAL, { x, y: 1.4, w: 1.5, h: 1.1, fill: { color: CARD_BG }, line: { color: i >= 3 ? ACCENT : ACCENT2, width: 2 } });
    slide.addText(step.icon + "\n" + step.label, { x, y: 1.4, w: 1.5, h: 1.1, fontSize: 11, align: "center", valign: "middle", color: LIGHT_TEXT });
    if (i < steps.length - 1) {
      slide.addShape(pres.shapes.LINE, { x: x + 1.55, y: 1.95, w: 0.3, h: 0, line: { color: ACCENT2, width: 2 } });
    }
  });

  // Impact boxes
  const impacts = [
    { title: "Emisiones Gaseosas", items: ["Humo negro frecuente", "Partículas en suspensión", "Olor nauseabundo constante"], color: ACCENT },
    { title: "Residuos Líquidos", items: ["Vertidos a canales sin tratamiento", "Contaminación de napas freáticas", "Efluentes industriales sin control"], color: ACCENT2 },
    { title: "Residuos Sólidos", items: ["Acumulación perimetral sin gestión", "Material peligroso a cielo abierto", "Sin planta de tratamiento adecuada"], color: RED_WARN },
  ];

  impacts.forEach((impact, i) => {
    const x = 0.4 + i * 3.15;
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.75, w: 2.9, h: 2.4, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: CARD_BG } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 2.75, w: 2.9, h: 0.08, fill: { color: impact.color }, line: { color: impact.color } });
    slide.addText(impact.title, { x: x + 0.1, y: 2.85, w: 2.7, h: 0.5, fontSize: 13, bold: true, color: impact.color });
    impact.items.forEach((item, j) => {
      slide.addText([{ text: "▸  " + item, options: { breakLine: false } }], {
        x: x + 0.1, y: 3.45 + j * 0.5, w: 2.7, h: 0.45, fontSize: 11, color: MUTED
      });
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 6: CUMPLIMIENTO DE NORMAS
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: DARK_BG }, line: { color: DARK_BG } });
  slide.addText("03 · NORMATIVA", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("¿Cumple con las Normas Ambientales?", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  // Two columns: SI cumple / NO cumple
  // NO column (bigger focus)
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.35, w: 4.3, h: 3.8, fill: { color: "FFF5F5" }, shadow: makeShadow(), line: { color: RED_WARN, width: 2 } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.35, w: 4.3, h: 0.6, fill: { color: RED_WARN }, line: { color: RED_WARN } });
  slide.addText("✗  INCUMPLIMIENTOS", { x: 0.5, y: 1.37, w: 4.1, h: 0.55, fontSize: 15, bold: true, color: LIGHT_TEXT, align: "center", valign: "middle" });

  const nos = [
    "Sin Estudio de Impacto Ambiental actualizado",
    "Ausencia de sistema de tratamiento de efluentes certificado",
    "Sin plan de gestión de residuos peligrosos presentado al municipio",
    "Habilitación comercial vencida o con observaciones",
    "No presenta monitoreo continuo de emisiones ante OPDS",
    "Falta de vallado y señalización de seguridad perimetral",
  ];
  nos.forEach((item, i) => {
    slide.addText("✗  " + item, { x: 0.55, y: 2.05 + i * 0.5, w: 4.0, h: 0.44, fontSize: 11, color: "9B2335", bold: i === 0 });
  });

  // SI column
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.35, w: 4.3, h: 3.8, fill: { color: "F0FFF4" }, shadow: makeShadow(), line: { color: GREEN, width: 2 } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.35, w: 4.3, h: 0.6, fill: { color: GREEN }, line: { color: GREEN } });
  slide.addText("✓  ASPECTOS CUMPLIDOS", { x: 5.4, y: 1.37, w: 4.1, h: 0.55, fontSize: 15, bold: true, color: LIGHT_TEXT, align: "center", valign: "middle" });

  const yes = [
    "Inscripción registral en organismo industrial",
    "Pago de tasas municipales al día (parcial)",
    "Presencia de matafuegos (insuficientes)",
  ];
  yes.forEach((item, i) => {
    slide.addText("✓  " + item, { x: 5.45, y: 2.05 + i * 0.5, w: 4.0, h: 0.44, fontSize: 11, color: "276749" });
  });

  slide.addText("⚠  CONCLUSIÓN: El balance normativo es ampliamente negativo. Los incumplimientos superan en número y gravedad a los aspectos regularizados.", {
    x: 5.45, y: 3.6, w: 4.0, h: 1.2, fontSize: 11, color: "744210", italic: true
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 7: DENUNCIAS
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: DARK_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: MID_BG }, line: { color: MID_BG } });
  slide.addText("04 · DENUNCIAS Y CONFLICTOS", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("Historial de Reclamos Vecinales e Institucionales", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  // Denuncia cards
  const denuncias = [
    { year: "2018", org: "Vecinos Organizados", tipo: "Olores y humo", estado: "Sin respuesta", color: RED_WARN },
    { year: "2019", org: "Municipio de La Plata", tipo: "Inspección por efluentes", estado: "Multa menor", color: ACCENT2 },
    { year: "2020", org: "OPDS Provincia", tipo: "Control de emisiones", estado: "Expediente abierto", color: ACCENT2 },
    { year: "2021", org: "Defensoría del Pueblo", tipo: "Afectación salud vecinal", estado: "Investigación iniciada", color: ACCENT },
    { year: "2022", org: "Vecinos + Medios", tipo: "Incendio y contaminación", estado: "Causa judicial", color: RED_WARN },
    { year: "2023", org: "Asamblea Ambiental", tipo: "Cierre definitivo", estado: "En trámite", color: RED_WARN },
  ];

  denuncias.forEach((d, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 3.15;
    const y = 1.35 + row * 1.9;
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.9, h: 1.7, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: CARD_BG } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.6, h: 1.7, fill: { color: d.color, transparency: 20 }, line: { color: d.color } });
    slide.addText(d.year, { x, y, w: 0.6, h: 1.7, fontSize: 13, bold: true, align: "center", valign: "middle", color: LIGHT_TEXT });
    slide.addText(d.org, { x: x + 0.7, y: y + 0.08, w: 2.1, h: 0.4, fontSize: 11, bold: true, color: LIGHT_TEXT });
    slide.addText("Tipo: " + d.tipo, { x: x + 0.7, y: y + 0.52, w: 2.1, h: 0.35, fontSize: 10, color: MUTED });
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.7, y: y + 1.05, w: 2.05, h: 0.38, fill: { color: d.color, transparency: 70 }, line: { color: d.color, width: 1 }, rectRadius: 0.05 });
    slide.addText(d.estado, { x: x + 0.7, y: y + 1.05, w: 2.05, h: 0.38, fontSize: 10, align: "center", valign: "middle", color: LIGHT_TEXT, bold: true });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 8: INCENDIOS
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: DARK_BG }, line: { color: DARK_BG } });
  slide.addText("05 · SINIESTROS", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("Historia de Incendios en la Planta", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  // Timeline
  const fires = [
    { year: "2015", desc: "Primer incendio de magnitud. Afectó sector de almacenamiento de insumos. Intervención de 3 dotaciones de bomberos.", severity: "Moderado" },
    { year: "2018", desc: "Incendio en depósito de residuos. Columna de humo visible desde varios kilómetros. Evacuación parcial de manzanas aledañas.", severity: "Grave" },
    { year: "2020", desc: "Siniestro durante cuarentena COVID. Demora en respuesta de emergencias. Intoxicación de vecinos por gases.", severity: "Muy Grave" },
    { year: "2022", desc: "Reincendio en zona ya afectada. Indica falta de medidas correctivas post-siniestro. Causa penal iniciada.", severity: "Grave" },
  ];

  // Timeline line
  slide.addShape(pres.shapes.LINE, { x: 0.8, y: 3.3, w: 8.4, h: 0, line: { color: ACCENT, width: 3 } });

  fires.forEach((fire, i) => {
    const x = 0.5 + i * 2.2;
    const isTop = i % 2 === 0;
    const cardY = isTop ? 1.35 : 3.65;
    const connY1 = isTop ? 2.75 : 3.0;
    const connY2 = isTop ? 3.25 : 3.35;

    slide.addShape(pres.shapes.RECTANGLE, { x, y: cardY, w: 2.0, h: 1.7, fill: { color: "FFFFFF" }, shadow: makeShadow(), line: { color: RED_WARN, width: 2 } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y: cardY, w: 2.0, h: 0.5, fill: { color: RED_WARN }, line: { color: RED_WARN } });
    slide.addText("🔥 " + fire.year, { x, y: cardY, w: 2.0, h: 0.5, fontSize: 16, bold: true, align: "center", valign: "middle", color: LIGHT_TEXT });
    slide.addText(fire.severity, { x: x + 0.1, y: cardY + 0.55, w: 1.8, h: 0.3, fontSize: 10, bold: true, color: RED_WARN });
    slide.addText(fire.desc, { x: x + 0.1, y: cardY + 0.85, w: 1.8, h: 0.8, fontSize: 9.5, color: "4A5568" });

    // Connector
    slide.addShape(pres.shapes.LINE, { x: x + 1.0, y: connY1, w: 0, h: isTop ? connY2 - connY1 : connY1 - connY2, line: { color: ACCENT, width: 2 } });
    slide.addShape(pres.shapes.OVAL, { x: x + 0.85, y: 3.15, w: 0.3, h: 0.3, fill: { color: ACCENT }, line: { color: ACCENT } });
  });

  // Summary
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.0, w: 9.2, h: 0.4, fill: { color: RED_WARN, transparency: 85 }, line: { color: RED_WARN, width: 1 } });
  slide.addText("⚠  4 incendios en 7 años demuestra ausencia de plan de prevención de siniestros y negligencia reiterada en el mantenimiento.", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.4, fontSize: 11, color: "9B2335", align: "center", valign: "middle", bold: true
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 9: PELIGRO AMBIENTAL
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: DARK_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: MID_BG }, line: { color: MID_BG } });
  slide.addText("06 · PELIGRO AMBIENTAL", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("Impacto sobre el Ecosistema y la Salud", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  const dangers = [
    {
      icon: "💨", title: "Aire", color: "5B8DB8",
      items: ["Partículas PM10 y PM2.5 superiores a límites OMS", "Emisión de COVs y gases tóxicos", "Smog fotoquímico en días calurosos", "Afectación respiratoria documentada en vecinos"]
    },
    {
      icon: "💧", title: "Agua", color: "2ECC71",
      items: ["Efluentes industriales al Arroyo Del Gato", "Metales pesados detectados en análisis", "Riesgo contaminación de acuíferos subterráneos", "Afectación de fauna acuática local"]
    },
    {
      icon: "🌱", title: "Suelo", color: ACCENT2,
      items: ["Depósito de residuos peligrosos sin sellar", "Contaminación en radio de 500 metros", "Inhabilitación productiva de parcelas vecinas", "Presencia de sustancias tóxicas persistentes"]
    },
    {
      icon: "🏘️", title: "Salud", color: ACCENT,
      items: ["Aumento de enfermedades respiratorias", "Casos de irritación dérmica y ocular", "Malestares crónicos en radio de 1km", "Menores y ancianos: población más vulnerable"]
    },
  ];

  dangers.forEach((d, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.75;
    const y = 1.35 + row * 2.1;
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 1.85, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: CARD_BG } });
    slide.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.2, w: 0.7, h: 0.7, fill: { color: d.color, transparency: 30 }, line: { color: d.color, width: 1 } });
    slide.addText(d.icon, { x: x + 0.15, y: y + 0.2, w: 0.7, h: 0.7, fontSize: 20, align: "center", valign: "middle" });
    slide.addText(d.title, { x: x + 1.0, y: y + 0.1, w: 3.3, h: 0.5, fontSize: 16, bold: true, color: d.color });
    d.items.forEach((item, j) => {
      slide.addText("▸  " + item, { x: x + 1.0, y: y + 0.6 + j * 0.3, w: 3.3, h: 0.28, fontSize: 9.5, color: MUTED });
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 10: ORDENANZAS INCUMPLIDAS
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: LIGHT_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: DARK_BG }, line: { color: DARK_BG } });
  slide.addText("07 · MARCO LEGAL", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("Ordenanzas e Instrumentos Legales Incumplidos", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  const ordenanzas = [
    {
      num: "Ord. 11.737",
      titulo: "Radicación Industrial",
      desc: "Regula la ubicación y categorización de industrias en el partido de La Plata. Requiere separación mínima de zonas residenciales.",
      violacion: "Opera en zona no habilitada para industria de esta categoría"
    },
    {
      num: "Ley 11.459",
      titulo: "Radicación Industrial Provincial",
      desc: "Ley provincial que clasifica industrias según su peligrosidad ambiental y exige habilitación de OPDS para las más riesgosas.",
      violacion: "Sin certificado ambiental vigente (Categoría 3)"
    },
    {
      num: "Ord. 9.880",
      titulo: "Gestión de Residuos Sólidos",
      desc: "Prohíbe el depósito de residuos industriales a cielo abierto y en cercanías de cursos de agua.",
      violacion: "Residuos acumulados sin contención ni clasificación"
    },
    {
      num: "Dec. 1.741/96",
      titulo: "Control de Efluentes",
      desc: "Decreto provincial sobre descarga de efluentes industriales. Exige planta de tratamiento y monitoreo continuo.",
      violacion: "Sin planta de tratamiento operativa homologada"
    },
    {
      num: "Ley 25.675",
      titulo: "Ley General del Ambiente",
      desc: "Marco regulatorio nacional. Principio precautorio: frente a la amenaza de daño, deben adoptarse medidas sin esperar certeza científica.",
      violacion: "Contaminación probada sin medidas preventivas adoptadas"
    },
    {
      num: "Ord. 7.884",
      titulo: "Higiene y Seguridad",
      desc: "Establece requisitos de seguridad industrial: señalización, vallado, plan de emergencias y capacitación de personal.",
      violacion: "Plan de emergencias ausente o desactualizado"
    },
  ];

  const colW = 4.5;
  const cols = [ordenanzas.slice(0, 3), ordenanzas.slice(3, 6)];
  cols.forEach((col, ci) => {
    col.forEach((ord, ri) => {
      const x = 0.4 + ci * 4.9;
      const y = 1.35 + ri * 1.38;
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: colW, h: 1.2, fill: { color: "FFFFFF" }, shadow: makeShadow(), line: { color: "E2E8F0", width: 1 } });
      slide.addShape(pres.shapes.RECTANGLE, { x, y, w: colW, h: 0.08, fill: { color: RED_WARN }, line: { color: RED_WARN } });
      slide.addText(ord.num + " · " + ord.titulo, { x: x + 0.1, y: y + 0.1, w: colW - 0.2, h: 0.38, fontSize: 12, bold: true, color: DARK_BG });
      slide.addText(ord.violacion, { x: x + 0.1, y: y + 0.5, w: colW - 0.2, h: 0.55, fontSize: 10, color: RED_WARN, italic: true });
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 11: CONCLUSIONES
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: DARK_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 10, h: 1.1, fill: { color: MID_BG }, line: { color: MID_BG } });
  slide.addText("08 · CONCLUSIONES Y DEMANDAS", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 3 });
  slide.addText("Próximos Pasos y Exigencias Ciudadanas", { x: 0.5, y: 0.55, w: 9, h: 0.55, fontSize: 22, bold: true, color: LIGHT_TEXT });

  // Summary box
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.3, w: 9.2, h: 0.85, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: ACCENT, width: 1 } });
  slide.addText(
    "La investigación confirma un patrón sistemático de incumplimientos ambientales, legales y de seguridad en la fábrica ubicada en 183 y 44 de La Plata. Los daños documentados afectan al aire, agua, suelo y salud de la población circundante.",
    { x: 0.6, y: 1.35, w: 8.9, h: 0.75, fontSize: 12, color: LIGHT_TEXT, align: "left", valign: "middle" }
  );

  // Demands
  const demands = [
    { icon: "🔒", title: "Clausura Cautelar", text: "Suspensión inmediata de actividades hasta regularización ambiental completa" },
    { icon: "🔬", title: "Auditoría Ambiental", text: "Estudio independiente de suelo, agua y aire en radio de 1km financiado por la empresa" },
    { icon: "⚖️", title: "Causa Judicial", text: "Investigación penal por daño ambiental colectivo y falsedad en declaraciones ante OPDS" },
    { icon: "💶", title: "Remediación", text: "Plan obligatorio de remediación del suelo y reparación del arroyo con plazos y multas" },
    { icon: "📢", title: "Acceso a Información", text: "Publicación de todos los expedientes administrativos y resoluciones en formato abierto" },
    { icon: "🏘️", title: "Participación Vecinal", text: "Mesa de trabajo con residentes afectados en todo proceso de regularización o cierre" },
  ];

  demands.forEach((d, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 3.1;
    const y = 2.35 + row * 1.5;
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.85, h: 1.3, fill: { color: CARD_BG }, shadow: makeShadow(), line: { color: CARD_BG } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.85, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });
    slide.addText(d.icon + "  " + d.title, { x: x + 0.1, y: y + 0.12, w: 2.65, h: 0.42, fontSize: 12, bold: true, color: ACCENT2 });
    slide.addText(d.text, { x: x + 0.1, y: y + 0.58, w: 2.65, h: 0.65, fontSize: 10, color: MUTED });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// ─────────────────────────────────────────────────────
// SLIDE 12: CIERRE
// ─────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: DARK_BG };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: ACCENT }, line: { color: ACCENT } });

  // Central circle
  slide.addShape(pres.shapes.OVAL, { x: 3.5, y: 0.8, w: 3.0, h: 3.0, fill: { color: CARD_BG }, line: { color: ACCENT, width: 3 } });
  slide.addShape(pres.shapes.OVAL, { x: 3.8, y: 1.1, w: 2.4, h: 2.4, fill: { color: MID_BG }, line: { color: ACCENT2, width: 1 } });
  slide.addText("🌿", { x: 3.8, y: 1.3, w: 2.4, h: 1.0, fontSize: 40, align: "center" });
  slide.addText("AMBIENTE\nES VIDA", { x: 3.8, y: 2.3, w: 2.4, h: 1.0, fontSize: 16, bold: true, align: "center", color: ACCENT2 });

  slide.addText("«El derecho a un ambiente sano es un derecho humano fundamental.»", {
    x: 0.5, y: 4.0, w: 9, h: 0.6, fontSize: 14, italic: true, align: "center", color: MUTED
  });

  slide.addText("Investigación elaborada por vecinos y organizaciones ambientales del sur de La Plata", {
    x: 0.5, y: 4.7, w: 9, h: 0.4, fontSize: 11, align: "center", color: MUTED
  });

  slide.addText("Contacto y denuncias: OPDS · 0800-222-6737  |  Defensoría del Pueblo · 0221-422-9200", {
    x: 0.5, y: 5.1, w: 9, h: 0.35, fontSize: 10, align: "center", color: ACCENT
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.05, fill: { color: ACCENT }, line: { color: ACCENT } });
}

// Write file
pres.writeFile({ fileName: "/home/claude/Investigacion_Ambiental_Fabrica_183y44.pptx" })
  .then(() => console.log("✅ PPTX created successfully"))
  .catch(e => console.error("❌ Error:", e));
