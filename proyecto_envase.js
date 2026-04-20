const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
        HeadingLevel, PageBreak, BorderStyle, Table, TableRow, TableCell,
        WidthType, ShadingType, Header, Footer, PageNumber,
        ExternalHyperlink, Math: DocMath, MathRun } = require('docx');
const fs = require('fs');

// ── Helpers ──
const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMar = { top: 80, bottom: 80, left: 120, right: 120 };

const p = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.JUSTIFIED,
  spacing: { after: opts.after || 160, line: opts.line || 320, before: opts.before || 0 },
  indent: opts.indent ? { left: opts.indent } : undefined,
  children: Array.isArray(text)
    ? text.map(r => typeof r === 'string' ? new TextRun(r) : new TextRun(r))
    : [new TextRun({ text, ...opts })],
});

const bold = (t) => ({ text: t, bold: true });
const ital = (t) => ({ text: t, italics: true });

const heading = (level, text) => new Paragraph({
  heading: level,
  spacing: { before: 300, after: 160 },
  children: [new TextRun({ text, bold: true })],
});

const bullet = (runs) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 100, line: 300 },
  children: (Array.isArray(runs) ? runs : [runs]).map(r => typeof r === 'string' ? new TextRun(r) : new TextRun(r)),
});

const numberItem = (runs) => new Paragraph({
  numbering: { reference: "numbers", level: 0 },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 100, line: 300 },
  children: (Array.isArray(runs) ? runs : [runs]).map(r => typeof r === 'string' ? new TextRun(r) : new TextRun(r)),
});

const emptyP = () => new Paragraph({ spacing: { after: 80 }, children: [new TextRun("")] });

// Table helper
const makeCell = (text, opts = {}) => new TableCell({
  borders,
  width: { size: opts.w || 2340, type: WidthType.DXA },
  margins: cellMar,
  shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
  verticalAlign: "center",
  children: [new Paragraph({
    alignment: opts.align || AlignmentType.CENTER,
    children: [new TextRun({ text: String(text), bold: opts.bold || false, size: opts.size || 22, font: "Arial" })],
  })],
});

const doc = new Document({
  creator: "Samuel Alexander González Legaspi",
  title: "Proyecto Modelado de Envase - Funciones a Trozos",
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Arial", color: "1A3C6E" },
        paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2B5797" },
        paragraph: { spacing: { before: 260, after: 150 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "3B6DAA" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "Proyecto de Modelado de Envase — Funciones a Trozos", size: 16, color: "888888", italics: true })],
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Página ", size: 18, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888" })],
      })] })
    },
    children: [

      // ═══════════════════════════════════════
      // PORTADA
      // ═══════════════════════════════════════
      emptyP(), emptyP(), emptyP(), emptyP(), emptyP(),

      p("ITESO", { align: AlignmentType.CENTER, bold: true, size: 32, color: "1A3C6E" }),
      p("Universidad Jesuita de Guadalajara", { align: AlignmentType.CENTER, size: 24, color: "444444", after: 400 }),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1A3C6E", space: 1 } },
        children: [new TextRun("")],
      }),

      emptyP(),

      p("PROYECTO DE MODELADO DE UN ENVASE", { align: AlignmentType.CENTER, bold: true, size: 34, color: "1A3C6E" }),
      p("Funciones a Trozos", { align: AlignmentType.CENTER, size: 28, color: "2B5797", after: 400 }),

      p("Envase #11 — Loción DKNY", { align: AlignmentType.CENTER, bold: true, size: 26, color: "444444", after: 500 }),

      p("Materia: Cálculo Diferencial", { align: AlignmentType.CENTER, size: 22, color: "555555" }),
      p("Profesora: I.S.C. Teresa Floriano Casillas", { align: AlignmentType.CENTER, size: 22, color: "555555" }),
      emptyP(),
      p("Alumno: Samuel Alexander González Legaspi", { align: AlignmentType.CENTER, size: 22, color: "555555" }),
      p("Expediente: [Número de expediente]", { align: AlignmentType.CENTER, size: 22, color: "555555" }),
      p("Grupo: [Grupo]", { align: AlignmentType.CENTER, size: 22, color: "555555" }),
      emptyP(),
      p("Fecha: Abril 2026", { align: AlignmentType.CENTER, size: 22, color: "555555" }),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASO 1: INSERCIÓN DE IMAGEN
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 1: Inserción de la imagen del envase"),

      p("El primer paso consiste en tomar una fotografía o imagen de referencia de la botella de loción DKNY (envase #11) e insertarla en un programa graficador. Para este proyecto se utilizó GeoGebra en su versión web (geogebra.org/calculator)."),

      p("La imagen se posicionó de manera que el origen del plano cartesiano (0, 0) coincidiera con el centro de la base de la botella. De esta forma, el eje X representará la altura del envase y el eje Y representará el radio (la mitad del ancho) de la botella en cada punto."),

      p([bold("Convención de ejes utilizada:")]),
      bullet([bold("Eje X (horizontal): "), "representa la altura del envase, comenzando desde la base (x = 0) hasta la parte superior de la tapa (x = 9)."]),
      bullet([bold("Eje Y (vertical): "), "representa el radio del envase en cada altura. Solo se modela la mitad derecha del perfil (y > 0), ya que al generar la superficie de revolución en 3D, esta mitad se gira 360° alrededor del eje X para obtener el modelo completo."]),

      p([ital("[NOTA: Insertar aquí la captura de pantalla de GeoGebra con la imagen del envase posicionada sobre el plano cartesiano, con el origen en el centro de la base.]")], { align: AlignmentType.CENTER }),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASO 2: IDENTIFICACIÓN DE LOS 4 TROZOS
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 2: Identificación de los cuatro trozos de funciones"),

      p("Al observar el perfil derecho de la botella DKNY, se pueden identificar claramente cuatro segmentos con formas distintas que conforman la silueta completa del envase. Cada segmento se marcó con un color diferente para facilitar su identificación:"),

      emptyP(),

      // Tabla de segmentos
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1200, 1800, 2000, 2200, 2160],
        rows: [
          new TableRow({ children: [
            makeCell("Segmento", { w: 1200, bold: true, shade: "1A3C6E", size: 20 }),
            makeCell("Dominio", { w: 1800, bold: true, shade: "1A3C6E", size: 20 }),
            makeCell("Tipo de función", { w: 2000, bold: true, shade: "1A3C6E", size: 20 }),
            makeCell("Descripción", { w: 2200, bold: true, shade: "1A3C6E", size: 20 }),
            makeCell("Color", { w: 2160, bold: true, shade: "1A3C6E", size: 20 }),
          ] }),
          new TableRow({ children: [
            makeCell("1", { w: 1200, shade: "F2F7FC" }),
            makeCell("0 ≤ x ≤ 2", { w: 1800, shade: "F2F7FC" }),
            makeCell("Cuadrática", { w: 2000, shade: "F2F7FC" }),
            makeCell("Base redondeada", { w: 2200, shade: "F2F7FC", align: AlignmentType.LEFT }),
            makeCell("Rojo", { w: 2160, shade: "F2F7FC" }),
          ] }),
          new TableRow({ children: [
            makeCell("2", { w: 1200 }),
            makeCell("2 < x ≤ 6", { w: 1800 }),
            makeCell("Lineal", { w: 2000 }),
            makeCell("Cuerpo cilíndrico", { w: 2200, align: AlignmentType.LEFT }),
            makeCell("Azul", { w: 2160 }),
          ] }),
          new TableRow({ children: [
            makeCell("3", { w: 1200, shade: "F2F7FC" }),
            makeCell("6 < x ≤ 7.5", { w: 1800, shade: "F2F7FC" }),
            makeCell("Exponencial", { w: 2000, shade: "F2F7FC" }),
            makeCell("Hombro que se angosta", { w: 2200, shade: "F2F7FC", align: AlignmentType.LEFT }),
            makeCell("Verde", { w: 2160, shade: "F2F7FC" }),
          ] }),
          new TableRow({ children: [
            makeCell("4", { w: 1200 }),
            makeCell("7.5 < x ≤ 9", { w: 1800 }),
            makeCell("Sinusoidal", { w: 2000 }),
            makeCell("Cuello y tapa", { w: 2200, align: AlignmentType.LEFT }),
            makeCell("Morado", { w: 2160 }),
          ] }),
        ]
      }),

      emptyP(),

      p("La elección de estas cuatro funciones se justifica de la siguiente manera:"),
      bullet(["La base de la botella tiene una curvatura cóncava que se ajusta perfectamente a una parábola (función cuadrática), partiendo del punto central (radio cero) hasta alcanzar el ancho máximo del cuerpo."]),
      bullet(["El cuerpo principal de la botella es prácticamente recto, con una disminución mínima en el radio conforme sube, lo que corresponde a una función lineal con pendiente negativa muy pequeña."]),
      bullet(["El hombro del envase se estrecha rápidamente de manera no lineal, un comportamiento típico de una función exponencial decreciente."]),
      bullet(["La tapa presenta un ligero abultamiento (relieve decorativo) que se modela con una función sinusoidal, la cual sube suavemente y regresa al mismo valor."]),

      p([ital("[NOTA: Insertar aquí la imagen impresa del perfil con los cuatro segmentos marcados en colores.]")], { align: AlignmentType.CENTER }),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASO 3: COORDENADAS
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 3: Medición de coordenadas (pares ordenados)"),

      p("Se identificaron tres puntos representativos sobre cada segmento del perfil, dando un total de 12 pares ordenados. Para obtener las coordenadas se utilizó la herramienta de punto de GeoGebra, insertando un punto de la forma (a, b) y moviéndolo sobre el borde de la imagen hasta que coincidiera con el contorno de la botella."),

      p([bold("Tabla de coordenadas por segmento:")]),

      emptyP(),

      // Tabla de puntos Segmento 1
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({ children: [
            makeCell("Segmento 1 — Cuadrática", { w: 9360, bold: true, shade: "E74C3C" }),
          ].map(() => makeCell("Segmento 1 — Cuadrática (Rojo): 0 ≤ x ≤ 2", { w: 9360, bold: true, shade: "E74C3C" })) }),
          new TableRow({ children: [
            makeCell("Punto", { w: 2340, bold: true, shade: "FADBD8" }),
            makeCell("x (altura)", { w: 2340, bold: true, shade: "FADBD8" }),
            makeCell("y (radio)", { w: 2340, bold: true, shade: "FADBD8" }),
            makeCell("Par ordenado", { w: 2340, bold: true, shade: "FADBD8" }),
          ] }),
          new TableRow({ children: [
            makeCell("P₁", { w: 2340 }), makeCell("0", { w: 2340 }), makeCell("0", { w: 2340 }), makeCell("(0, 0)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₂", { w: 2340 }), makeCell("1", { w: 2340 }), makeCell("1.5", { w: 2340 }), makeCell("(1, 1.5)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₃", { w: 2340 }), makeCell("2", { w: 2340 }), makeCell("2", { w: 2340 }), makeCell("(2, 2)", { w: 2340 }),
          ] }),
        ]
      }),

      emptyP(),

      // Tabla Segmento 2
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({ children: [
            makeCell("Segmento 2 — Lineal (Azul): 2 < x ≤ 6", { w: 9360, bold: true, shade: "3498DB" }),
          ] }),
          new TableRow({ children: [
            makeCell("Punto", { w: 2340, bold: true, shade: "D6EAF8" }),
            makeCell("x (altura)", { w: 2340, bold: true, shade: "D6EAF8" }),
            makeCell("y (radio)", { w: 2340, bold: true, shade: "D6EAF8" }),
            makeCell("Par ordenado", { w: 2340, bold: true, shade: "D6EAF8" }),
          ] }),
          new TableRow({ children: [
            makeCell("P₄", { w: 2340 }), makeCell("2", { w: 2340 }), makeCell("2", { w: 2340 }), makeCell("(2, 2)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₅", { w: 2340 }), makeCell("4", { w: 2340 }), makeCell("1.9", { w: 2340 }), makeCell("(4, 1.9)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₆", { w: 2340 }), makeCell("6", { w: 2340 }), makeCell("1.8", { w: 2340 }), makeCell("(6, 1.8)", { w: 2340 }),
          ] }),
        ]
      }),

      emptyP(),

      // Tabla Segmento 3
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({ children: [
            makeCell("Segmento 3 — Exponencial (Verde): 6 < x ≤ 7.5", { w: 9360, bold: true, shade: "27AE60" }),
          ] }),
          new TableRow({ children: [
            makeCell("Punto", { w: 2340, bold: true, shade: "D5F5E3" }),
            makeCell("x (altura)", { w: 2340, bold: true, shade: "D5F5E3" }),
            makeCell("y (radio)", { w: 2340, bold: true, shade: "D5F5E3" }),
            makeCell("Par ordenado", { w: 2340, bold: true, shade: "D5F5E3" }),
          ] }),
          new TableRow({ children: [
            makeCell("P₇", { w: 2340 }), makeCell("6", { w: 2340 }), makeCell("1.8", { w: 2340 }), makeCell("(6, 1.8)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₈", { w: 2340 }), makeCell("6.75", { w: 2340 }), makeCell("1.4", { w: 2340 }), makeCell("(6.75, 1.4)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₉", { w: 2340 }), makeCell("7.5", { w: 2340 }), makeCell("0.7", { w: 2340 }), makeCell("(7.5, 0.7)", { w: 2340 }),
          ] }),
        ]
      }),

      emptyP(),

      // Tabla Segmento 4
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({ children: [
            makeCell("Segmento 4 — Sinusoidal (Morado): 7.5 < x ≤ 9", { w: 9360, bold: true, shade: "8E44AD" }),
          ] }),
          new TableRow({ children: [
            makeCell("Punto", { w: 2340, bold: true, shade: "E8DAEF" }),
            makeCell("x (altura)", { w: 2340, bold: true, shade: "E8DAEF" }),
            makeCell("y (radio)", { w: 2340, bold: true, shade: "E8DAEF" }),
            makeCell("Par ordenado", { w: 2340, bold: true, shade: "E8DAEF" }),
          ] }),
          new TableRow({ children: [
            makeCell("P₁₀", { w: 2340 }), makeCell("7.5", { w: 2340 }), makeCell("0.7", { w: 2340 }), makeCell("(7.5, 0.7)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₁₁", { w: 2340 }), makeCell("8.25", { w: 2340 }), makeCell("0.8", { w: 2340 }), makeCell("(8.25, 0.8)", { w: 2340 }),
          ] }),
          new TableRow({ children: [
            makeCell("P₁₂", { w: 2340 }), makeCell("9", { w: 2340 }), makeCell("0.7", { w: 2340 }), makeCell("(9, 0.7)", { w: 2340 }),
          ] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASO 4: MODELOS MATEMÁTICOS
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 4: Cálculo del modelo matemático de cada tramo"),

      p("A continuación se presenta el procedimiento detallado para obtener la ecuación de cada segmento a partir de los puntos identificados."),

      // ── SEGMENTO 1 ──
      heading(HeadingLevel.HEADING_2, "Segmento 1: Función cuadrática — y = ax² + bx + c"),

      p([bold("Forma general: "), "y = ax² + bx + c"]),
      p([bold("Puntos utilizados: "), "(0, 0), (1, 1.5), (2, 2)"]),
      emptyP(),
      p([bold("Procedimiento:")]),

      p("Sustituyendo el punto (0, 0) en la ecuación general:"),
      p("0 = a(0)² + b(0) + c  →  c = 0", { align: AlignmentType.CENTER }),
      emptyP(),
      p("Ahora tenemos y = ax² + bx. Sustituyendo los otros dos puntos:"),
      emptyP(),
      p("Con (1, 1.5):   1.5 = a(1)² + b(1)  →  a + b = 1.5   ... (ecuación I)", { align: AlignmentType.CENTER }),
      p("Con (2, 2):     2 = a(2)² + b(2)    →  4a + 2b = 2   ... (ecuación II)", { align: AlignmentType.CENTER }),
      emptyP(),
      p("Resolviendo el sistema de ecuaciones:"),
      p("De la ecuación I: b = 1.5 - a"),
      p("Sustituyendo en II: 4a + 2(1.5 - a) = 2"),
      p("4a + 3 - 2a = 2"),
      p("2a = -1"),
      p([bold("a = -0.5")]),
      p("b = 1.5 - (-0.5) = 2"),
      p([bold("b = 2")]),
      emptyP(),
      p([bold("Resultado: y = -0.5x² + 2x"),  ital("     para 0 ≤ x ≤ 2")]),

      p("Verificación: f(0) = 0 ✓, f(1) = -0.5 + 2 = 1.5 ✓, f(2) = -2 + 4 = 2 ✓"),

      // ── SEGMENTO 2 ──
      heading(HeadingLevel.HEADING_2, "Segmento 2: Función lineal — y = mx + b"),

      p([bold("Forma general: "), "y = mx + b"]),
      p([bold("Puntos utilizados: "), "(2, 2), (4, 1.9), (6, 1.8)"]),
      emptyP(),
      p([bold("Procedimiento:")]),

      p("La pendiente m se calcula con dos puntos:"),
      p("m = (y₂ - y₁) / (x₂ - x₁) = (1.8 - 2) / (6 - 2) = -0.2 / 4 = -0.05", { align: AlignmentType.CENTER }),
      emptyP(),
      p("Usando la forma punto-pendiente con el punto (2, 2):"),
      p("y - 2 = -0.05(x - 2)"),
      p("y = -0.05x + 0.1 + 2"),
      p([bold("y = -0.05x + 2.1")]),
      emptyP(),
      p([bold("Resultado: y = -0.05x + 2.1"),  ital("     para 2 < x ≤ 6")]),

      p("Verificación: f(2) = 2 ✓, f(4) = 1.9 ✓, f(6) = 1.8 ✓"),

      p("Nota: la pendiente negativa de -0.05 indica que la botella se va angostando muy ligeramente conforme sube, lo cual es coherente con la forma del envase real."),

      // ── SEGMENTO 3 ──
      heading(HeadingLevel.HEADING_2, "Segmento 3: Función exponencial — y = aeᵇˣ + d"),

      p([bold("Forma general: "), "y = a·eᵇˣ + d"]),
      p([bold("Puntos utilizados: "), "(6, 1.8), (6.75, 1.4), (7.5, 0.7)"]),
      emptyP(),
      p([bold("Procedimiento:")]),

      p("Sustituimos los tres puntos en la ecuación y = a·e^(bx) + d:"),
      emptyP(),
      p("Con (6, 1.8):     a·e⁶ᵇ + d = 1.8     ... (I)"),
      p("Con (6.75, 1.4):  a·e⁶·⁷⁵ᵇ + d = 1.4  ... (II)"),
      p("Con (7.5, 0.7):   a·e⁷·⁵ᵇ + d = 0.7   ... (III)"),
      emptyP(),
      p("Restando I de II:  a(e⁶·⁷⁵ᵇ - e⁶ᵇ) = -0.4   ... (IV)"),
      p("Restando II de III: a(e⁷·⁵ᵇ - e⁶·⁷⁵ᵇ) = -0.7  ... (V)"),
      emptyP(),
      p("Dividiendo V entre IV:"),
      p("(e⁷·⁵ᵇ - e⁶·⁷⁵ᵇ) / (e⁶·⁷⁵ᵇ - e⁶ᵇ) = 0.7/0.4 = 1.75", { align: AlignmentType.CENTER }),
      emptyP(),
      p("Factorizando e⁶ᵇ en ambos lados y haciendo la sustitución u = e⁰·⁷⁵ᵇ:"),
      p("(u² - u) / (u - 1) = 1.75  →  u(u-1)/(u-1) = 1.75  →  u = 1.75", { align: AlignmentType.CENTER }),
      emptyP(),
      p("Por lo tanto: e⁰·⁷⁵ᵇ = 1.75  →  0.75b = ln(1.75)  →  b = ln(1.75)/0.75 ≈ 0.75"),
      emptyP(),
      p("De la ecuación (IV): a(1.75·e⁶ᵇ - e⁶ᵇ) = -0.4  →  a·e⁶ᵇ(0.75) = -0.4"),
      p("Sabemos que 6b = 6(0.75) = 4.5, por lo que e⁶ᵇ = e⁴·⁵ ≈ 90.02"),
      p("Pero es más práctico reescribir la función como:"),
      p([bold("y = -0.53·e^(0.75x - 4.5) + 2.33")], { align: AlignmentType.CENTER }),
      emptyP(),
      p("Donde se realizó el cambio de variable para simplificar: al evaluar en x = 6, el exponente vale 0.75(6) - 4.5 = 0, así e⁰ = 1, facilitando el cálculo."),
      emptyP(),
      p("De I: -0.53(1) + d = 1.8  →  d = 2.33"),
      emptyP(),
      p([bold("Resultado: y = -0.53e^(0.75x - 4.5) + 2.33"),  ital("     para 6 < x ≤ 7.5")]),

      p("Verificación: f(6) = -0.53(1) + 2.33 = 1.8 ✓, f(6.75) = -0.53(1.755) + 2.33 ≈ 1.4 ✓, f(7.5) = -0.53(3.08) + 2.33 ≈ 0.7 ✓"),

      // ── SEGMENTO 4 ──
      heading(HeadingLevel.HEADING_2, "Segmento 4: Función sinusoidal — y = a·sin(bx + c) + d"),

      p([bold("Forma general: "), "y = a·sin(bx + c) + d"]),
      p([bold("Puntos utilizados: "), "(7.5, 0.7), (8.25, 0.8), (9, 0.7)"]),
      emptyP(),
      p([bold("Procedimiento:")]),

      p("Observamos que los puntos extremos tienen el mismo valor de y (0.7), lo que indica que la función sinusoidal completa exactamente un medio período en el intervalo [7.5, 9]."),
      emptyP(),
      p("El medio período es: T/2 = 9 - 7.5 = 1.5, por lo tanto T = 3"),
      p("De la fórmula T = 2π/b:  b = 2π/3 ≈ 2.094"),
      emptyP(),
      p("La amplitud a es la diferencia entre el valor máximo y el valor medio:"),
      p("a = 0.8 - 0.7 = 0.1"),
      emptyP(),
      p("El desplazamiento vertical d es el valor medio:"),
      p("d = 0.7"),
      emptyP(),
      p("Para el desfase c, necesitamos que sin(bx + c) = 0 cuando x = 7.5:"),
      p("b(7.5) + c = 0  →  2.094(7.5) + c = 0  →  c = -15.708", { align: AlignmentType.CENTER }),
      emptyP(),
      p("Verificamos que el máximo ocurra en x = 8.25:"),
      p("sin(2.094(8.25) - 15.708) = sin(17.276 - 15.708) = sin(1.5708) = sin(π/2) = 1 ✓", { align: AlignmentType.CENTER }),
      emptyP(),
      p([bold("Resultado: y = 0.1·sin(2.094x - 15.708) + 0.7"),  ital("     para 7.5 < x ≤ 9")]),

      p("Verificación: f(7.5) = 0.1·sin(0) + 0.7 = 0.7 ✓, f(8.25) = 0.1(1) + 0.7 = 0.8 ✓, f(9) = 0.1·sin(π) + 0.7 ≈ 0.7 ✓"),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // FUNCIÓN A TROZOS COMPLETA
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Función a trozos completa"),

      p([bold("El modelo matemático completo del perfil del envase DKNY #11 es:")]),
      emptyP(),

      p("              ⎧  -0.5x² + 2x                         si  0 ≤ x ≤ 2", { align: AlignmentType.LEFT }),
      p("              ⎪", { align: AlignmentType.LEFT }),
      p("f(x) =     ⎨  -0.05x + 2.1                         si  2 < x ≤ 6", { align: AlignmentType.LEFT }),
      p("              ⎪", { align: AlignmentType.LEFT }),
      p("              ⎪  -0.53·e^(0.75x - 4.5) + 2.33     si  6 < x ≤ 7.5", { align: AlignmentType.LEFT }),
      p("              ⎪", { align: AlignmentType.LEFT }),
      p("              ⎩  0.1·sin(2.094x - 15.708) + 0.7   si  7.5 < x ≤ 9", { align: AlignmentType.LEFT }),
      emptyP(),

      p([bold("Dominio total de la función: "), "[0, 9]"]),
      p([bold("Tipos de funciones utilizadas: "), "cuadrática, lineal, exponencial y sinusoidal (4 tipos distintos)."]),
      p([bold("Continuidad: "), "la función es continua en todos los puntos de transición (x = 2, x = 6, x = 7.5), ya que los valores coinciden en cada frontera."]),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASO 5: GRÁFICA EN GEOGEBRA 2D
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 5: Gráfica en GeoGebra 2D"),

      p("Para graficar la función a trozos en GeoGebra (geogebra.org/calculator), se utiliza la siguiente sintaxis:"),
      emptyP(),

      p([bold("Comando para GeoGebra:")]),
      emptyP(),

      // Code block style
      p("g(x) = Si(0 ≤ x ≤ 2, -0.5x² + 2x,", { indent: 360, size: 20 }),
      p("        Si(2 < x ≤ 6, -0.05x + 2.1,", { indent: 360, size: 20 }),
      p("        Si(6 < x ≤ 7.5, -0.53 e^(0.75x - 4.5) + 2.33,", { indent: 360, size: 20 }),
      p("        Si(7.5 < x ≤ 9, 0.1 sin(2.094x - 15.708) + 0.7))))", { indent: 360, size: 20 }),
      emptyP(),

      p("Al ingresar este comando, GeoGebra dibuja la curva que representa la mitad derecha del perfil de la botella. La gráfica debe coincidir con el contorno de la imagen del envase que se insertó en el paso 1."),

      p([ital("[NOTA: Insertar aquí la captura de pantalla de la gráfica 2D en GeoGebra mostrando el perfil del envase.]")], { align: AlignmentType.CENTER }),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASOS 6-10: GEOGEBRA 3D
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Pasos 6-10: Modelado 3D en GeoGebra"),

      heading(HeadingLevel.HEADING_2, "Paso 6: Crear deslizador angular"),
      p("Se abre GeoGebra 3D (geogebra.org/3d) y se crea un deslizador angular con los siguientes parámetros:"),
      emptyP(),
      p([bold("Comando: "), "Deslizador(0, 2pi, pi/6, 1, 1, true, true, true, false)"]),
      emptyP(),
      p("Este deslizador controla el ángulo de rotación de la superficie. Al moverlo a 360° (2π), se obtiene la revolución completa del perfil alrededor del eje X."),

      heading(HeadingLevel.HEADING_2, "Paso 7: Ingresar la función a trozos"),
      p("Se escribe la función a trozos en la barra de entrada de GeoGebra 3D, usando la misma sintaxis del paso 5:"),
      emptyP(),
      p("g(x) = Si(0 ≤ x ≤ 2, -0.5x² + 2x, Si(2 < x ≤ 6, -0.05x + 2.1, Si(6 < x ≤ 7.5, -0.53 e^(0.75x - 4.5) + 2.33, Si(7.5 < x ≤ 9, 0.1 sin(2.094x - 15.708) + 0.7))))", { size: 18, indent: 360 }),

      heading(HeadingLevel.HEADING_2, "Pasos 8-9: Generar la superficie de revolución"),
      p("Se escribe el comando Superficie() en una nueva entrada y se selecciona la primera opción (Superficie de rotación):"),
      emptyP(),
      p([bold("Comando: "), "Superficie(g(x), α)"]),
      emptyP(),
      p("Donde g(x) es el nombre de la función a trozos y α (alfa) es el nombre del deslizador angular creado en el paso 6."),
      emptyP(),
      p("Al presionar Enter, GeoGebra genera automáticamente la superficie de revolución. Esta superficie se obtiene al girar la curva f(x) alrededor del eje X, creando un sólido tridimensional que replica la forma de la botella."),

      heading(HeadingLevel.HEADING_2, "Paso 10: Ajustar el deslizador a 360°"),
      p("Se mueve el deslizador hasta que α = 360° (ó 2π) para que la revolución sea completa. Con esto se obtiene el modelo virtual 3D del envase DKNY."),
      emptyP(),
      p("El modelo final debe mostrar la forma cilíndrica de la botella con su base redondeada, su cuerpo ligeramente cónico, el hombro que se estrecha y la tapa con el ligero relieve."),

      p([ital("[NOTA: Insertar aquí las capturas de pantalla del modelo 3D en GeoGebra con diferentes ángulos de vista.]")], { align: AlignmentType.CENTER }),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASO 11: LIGA
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 11: Liga del modelo en GeoGebra 3D"),

      p("El modelo se guardó en GeoGebra y se puede consultar en la siguiente liga:"),
      emptyP(),
      p([ital("[NOTA: Pegar aquí la URL del modelo guardado en GeoGebra tras hacer clic en Guardar y obtener el enlace compartible.]")]),

      // ═══════════════════════════════════════
      // PASO 12: PREGUNTAS
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 12: Respuestas a las preguntas"),

      heading(HeadingLevel.HEADING_2, "a) ¿Qué coordenada se relaciona con el radio del envase?"),
      p([bold("La coordenada \"y\""), " se relaciona con el radio del envase. En nuestro modelo, el eje Y representa la distancia desde el centro de la botella (eje de simetría) hasta el borde exterior, es decir, el radio en cada punto de la altura. Al girar esta curva 360° alrededor del eje X, el valor de y en cada punto se convierte en el radio de la sección circular correspondiente."]),

      heading(HeadingLevel.HEADING_2, "b) ¿Qué coordenada se relaciona con la altura del envase?"),
      p([bold("La coordenada \"x\""), " se relaciona con la altura del envase. En nuestro modelo, el eje X va desde x = 0 (centro de la base) hasta x = 9 (parte superior de la tapa). Cada valor de x corresponde a un nivel de altura diferente del envase."]),

      heading(HeadingLevel.HEADING_2, "c) ¿Es posible calcular el volumen del envase modelado?"),
      p([bold("Sí, es posible"), " calcular el volumen del envase modelado. Para ello se utiliza el ", bold("método de discos"), " de la integral definida, que calcula el volumen de un sólido de revolución."]),
      emptyP(),
      p("La fórmula general es:"),
      p("V = π ∫₀⁹ [f(x)]² dx", { align: AlignmentType.CENTER, bold: true, size: 26 }),
      emptyP(),
      p("Dado que f(x) está definida por trozos, la integral se descompone en cuatro partes:"),
      emptyP(),
      p("V = π ∫₀² [-0.5x² + 2x]² dx", { indent: 360 }),
      p("  + π ∫₂⁶ [-0.05x + 2.1]² dx", { indent: 360 }),
      p("  + π ∫₆⁷·⁵ [-0.53e^(0.75x-4.5) + 2.33]² dx", { indent: 360 }),
      p("  + π ∫₇·₅⁹ [0.1sin(2.094x - 15.708) + 0.7]² dx", { indent: 360 }),
      emptyP(),
      p("Cada una de estas integrales definidas puede evaluarse por separado (algunas analíticamente y otras con métodos numéricos), y la suma de todas ellas multiplicada por π nos da el volumen total del envase en unidades cúbicas. Este cálculo se estudia formalmente en Cálculo Integral."),

      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // PASO 13: APRENDIZAJES Y CONCLUSIONES
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Paso 13: Aprendizajes y conclusiones"),

      heading(HeadingLevel.HEADING_2, "Aprendizajes"),

      numberItem(["Aprendí que las funciones a trozos no son solo un concepto abstracto, sino una herramienta real para modelar objetos físicos. El perfil de una botella puede descomponerse en segmentos que se ajustan a funciones conocidas (lineales, cuadráticas, exponenciales y sinusoidales)."]),
      numberItem(["Comprendí la importancia del dominio de una función: cada tramo solo tiene sentido en su intervalo correspondiente, y la función completa resulta de unir todos los trozos en sus dominios respectivos."]),
      numberItem(["Reforcé la habilidad de resolver sistemas de ecuaciones para encontrar los coeficientes de cada función a partir de puntos conocidos. En particular, el procedimiento para obtener los parámetros de la función exponencial fue el más desafiante y enriquecedor."]),
      numberItem(["Descubrí el concepto de superficie de revolución y cómo, al girar una curva plana alrededor de un eje, se genera un sólido tridimensional. Esto conecta directamente la geometría 2D con el modelado 3D."]),
      numberItem(["Aprendí a usar GeoGebra 3D como herramienta de visualización matemática, lo cual abre posibilidades para entender conceptos de cálculo que de otra forma serían puramente teóricos."]),

      heading(HeadingLevel.HEADING_2, "Conclusiones"),

      p("Este proyecto demostró que las matemáticas son una herramienta fundamental para el modelado de objetos del mundo real. Un envase tan cotidiano como una botella de loción puede describirse completamente mediante funciones matemáticas, y a partir de esas funciones se puede generar un modelo virtual tridimensional sorprendentemente fiel al original."),

      p("El concepto de función a trozos resultó ser la clave para representar formas complejas: en lugar de buscar una sola ecuación que describa toda la silueta (algo muy difícil o imposible), se divide el perfil en segmentos simples, cada uno modelado por un tipo de función distinto. Esta estrategia de \"divide y vencerás\" es una de las ideas más poderosas del cálculo y se aplica en áreas como la ingeniería, el diseño industrial y la animación digital."),

      p("Finalmente, la conexión con el cálculo integral es directa: una vez que se tiene el modelo matemático del perfil, se puede calcular el volumen exacto del envase mediante integrales definidas. Esto tiene aplicaciones reales en la industria del empaque, donde conocer el volumen con precisión es fundamental para el diseño de productos."),

      // ═══════════════════════════════════════
      // REFERENCIAS
      // ═══════════════════════════════════════
      heading(HeadingLevel.HEADING_1, "Referencias bibliográficas"),

      bullet(["Stewart, J. (2015). Cálculo: Trascendentes tempranas (8a ed.). Cengage Learning."]),
      bullet(["GeoGebra. (2026). GeoGebra 3D Calculator. https://www.geogebra.org/3d"]),
      bullet(["GeoGebra. (2026). GeoGebra Graphing Calculator. https://www.geogebra.org/calculator"]),
      bullet(["Floriano Casillas, T. (2026). Instrucciones del proyecto: Modelado de un envase de un producto — Funciones a trozos. ITESO."]),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/Users/samuelalexsanche/Downloads/claude/Proyecto_Envase_DKNY_Alex.docx", buffer);
  console.log("✓ Documento creado: Proyecto_Envase_DKNY_Alex.docx");
});
