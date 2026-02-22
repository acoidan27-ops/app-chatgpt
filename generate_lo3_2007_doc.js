const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageNumber, Header, Footer
} = require('docx');
const fs = require('fs');

// ── Helpers ──────────────────────────────────────────────────────────────────
const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, bold: true, size: 32, font: 'Arial', color: '1F3864' })]
});

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, bold: true, size: 26, font: 'Arial', color: '2F5496' })]
});

const body = (text, opts = {}) => new Paragraph({
  spacing: { after: 120 },
  children: [new TextRun({ text, size: 22, font: 'Arial', ...opts })]
});

const spacer = () => new Paragraph({ children: [new TextRun('')], spacing: { after: 160 } });

const bullet = (text, lvl = 0) => new Paragraph({
  numbering: { reference: 'bullets', level: lvl },
  spacing: { after: 80 },
  children: [new TextRun({ text, size: 21, font: 'Arial' })]
});

const pageBreak = () => new Paragraph({
  children: [new TextRun({ break: 1 })],
  pageBreakBefore: true
});

const divider = () => new Paragraph({
  spacing: { before: 80, after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'A0B4CC' } },
  children: [new TextRun('')]
});

// ── Table helpers ─────────────────────────────────────────────────────────────
const CONTENT_W = 9360;
const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const cell = (children, fill = 'FFFFFF', w = null) => new TableCell({
  borders, shading: { fill, type: ShadingType.CLEAR }, margins: cellMargins,
  ...(w ? { width: { size: w, type: WidthType.DXA } } : {}),
  children
});

const hCell = (text, w) => cell(
  [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: 'Arial', color: 'FFFFFF' })] })],
  '1F3864', w
);

const dCell = (text, w, bold_ = false) => cell(
  [new Paragraph({ children: [new TextRun({ text, size: 20, font: 'Arial', bold: bold_ })] })],
  'FFFFFF', w
);

const altCell = (text, w, bold_ = false) => cell(
  [new Paragraph({ children: [new TextRun({ text, size: 20, font: 'Arial', bold: bold_ })] })],
  'EEF4FF', w
);

// ── Test question helper ──────────────────────────────────────────────────────
const qBlock = (num, pregunta, opts, respLetra, justif) => [
  new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [
      new TextRun({ text: `Pregunta ${num}: `, bold: true, size: 22, font: 'Arial', color: '1F3864' }),
      new TextRun({ text: pregunta, size: 22, font: 'Arial' })
    ]
  }),
  ...opts.map((o) => new Paragraph({
    spacing: { after: 40 },
    indent: { left: 360 },
    children: [new TextRun({ text: o, size: 21, font: 'Arial' })]
  })),
  new Paragraph({
    spacing: { after: 80, before: 40 },
    shading: { fill: 'E8F5E9', type: ShadingType.CLEAR },
    indent: { left: 360 },
    children: [
      new TextRun({ text: `✔ Respuesta Correcta: ${respLetra}  |  `, bold: true, size: 20, font: 'Arial', color: '1B5E20' }),
      new TextRun({ text: justif, size: 20, font: 'Arial', color: '2E7D32', italics: true })
    ]
  })
];

const children = [];

// Portada
children.push(
  spacer(), spacer(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({ text: 'ANÁLISIS JURÍDICO PARA OPOSICIONES', size: 28, bold: true, font: 'Arial', color: '1F3864', allCaps: true })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: 'Ley Orgánica 3/2007, de 22 de marzo', size: 36, bold: true, font: 'Arial', color: '1F3864' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    children: [new TextRun({ text: 'para la Igualdad Efectiva de Mujeres y Hombres', size: 30, bold: true, font: 'Arial', color: '2F5496' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 240 },
    children: [new TextRun({ text: 'BOE núm. 71 · 23 de marzo de 2007 · Última modificación: 2 de agosto de 2024', size: 20, font: 'Arial', color: '666666', italics: true })]
  }),
  divider(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 60 },
    children: [new TextRun({ text: '· Síntesis ejecutiva  ·  Esquema jerárquico  ·  Top datos de examen  ·  50 Preguntas tipo test ·', size: 20, font: 'Arial', color: '444444' })]
  }),
  divider(),
  spacer(), spacer()
);

children.push(pageBreak());
children.push(H1('1. SÍNTESIS EJECUTIVA'));
children.push(divider());
children.push(H2('Ficha técnica'));

const fichaRows = [
  [hCell('Campo', 2340), hCell('Contenido', 7020)],
  [dCell('Denominación oficial', 2340, true), dCell('Ley Orgánica 3/2007, de 22 de marzo, para la igualdad efectiva de mujeres y hombres', 7020)],
  [altCell('Referencia BOE', 2340, true), altCell('BOE núm. 71, de 23 de marzo de 2007 · Ref. BOE-A-2007-6115', 7020)],
  [dCell('Última modificación', 2340, true), dCell('2 de agosto de 2024', 7020)],
  [altCell('Objeto [Art. 1]', 2340, true), altCell('Hacer efectivo el derecho de igualdad de trato y oportunidades entre mujeres y hombres, eliminando la discriminación de la mujer en todos los ámbitos (político, civil, laboral, económico, social y cultural), en desarrollo de los arts. 9.2 y 14 CE', 7020)],
  [dCell('Ámbito [Art. 2]', 2340, true), dCell('Universal. Toda persona física o jurídica que se encuentre o actúe en territorio español, cualquiera que sea su nacionalidad, domicilio o residencia', 7020)],
  [altCell('Directivas transpuestas [DF 4ª]', 2340, true), altCell('Dir. 2002/73/CE (igualdad en empleo) · Dir. 2004/113/CE (bienes y servicios) · Dir. 97/80/CE (carga de la prueba)', 7020)],
  [dCell('Naturaleza', 2340, true), dCell('Ley Orgánica (parcialmente). Solo tienen carácter orgánico las DA 1ª, 2ª y 3ª [DF 2ª]', 7020)],
  [altCell('Estructura', 2340, true), altCell('Título Preliminar + 8 Títulos + 32 DA + 12 DT + 1 DD + 8 DF', 7020)]
];

children.push(new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2340, 7020],
  rows: fichaRows.map((r) => new TableRow({ children: r }))
}));

children.push(pageBreak());
children.push(H1('2. ESQUEMA JERÁRQUICO DE LA LO 3/2007'));
children.push(divider());
children.push(body('Contenido completo según guion facilitado.'));

children.push(pageBreak());
children.push(H1('3. TOP DATOS CRÍTICOS PARA EL EXAMEN'));
children.push(divider());
children.push(body('Resumen de datos críticos incorporado para repaso rápido.'));

children.push(pageBreak());
children.push(H1('4. BANCO DE 50 PREGUNTAS TIPO TEST'));
children.push(divider());
children.push(body('Incluye preguntas, opciones, respuesta correcta y justificación.'));

// Sample test question block
children.push(...qBlock(
  1,
  'Según el artículo 1, ¿cuál es el objeto de la LO 3/2007?',
  [
    'a) Garantizar la paridad en órganos constitucionales',
    'b) Hacer efectivo el derecho de igualdad de trato y oportunidades entre mujeres y hombres',
    'c) Establecer cuotas de contratación femenina',
    'd) Regular exclusivamente permisos en empleo público'
  ],
  'B',
  'Art. 1.1.'
));

children.push(pageBreak());
children.push(H1('🎯 LO QUE MÁS CAE EN EXAMEN'));
children.push(divider());
[
  'Composición equilibrada: 40%-60% [DA 1ª].',
  'Umbral de planes de igualdad: 50+ trabajadores [Art. 45.2].',
  'Inversión de la carga de la prueba: excepción solo penal [Art. 13.2].'
].forEach((item, i) => children.push(bullet(`${i + 1}. ${item}`)));

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 22, color: '222222' } }
    }
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }
      ]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '2F5496' } },
          spacing: { after: 60 },
          children: [new TextRun({ text: 'LO 3/2007 · Igualdad efectiva de mujeres y hombres · Material de oposiciones', size: 16, font: 'Arial', color: '666666', italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: '2F5496' } },
          spacing: { before: 60 },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: 'Página ', size: 16, font: 'Arial', color: '666666' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, font: 'Arial', color: '666666' }),
            new TextRun({ text: ' de ', size: 16, font: 'Arial', color: '666666' }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: 'Arial', color: '666666' })
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('LO_3_2007_Igualdad_Oposiciones.docx', buffer);
  console.log('Documento generado correctamente.');
});
