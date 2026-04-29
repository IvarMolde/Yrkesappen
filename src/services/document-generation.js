'use strict';

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, TabStopType, TabStopPosition, PageNumber, Header, Footer,
} = require('docx');
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const C = {
  primary: '005F73',
  secondary: '0A9396',
  accent: 'E9C46A',
  bgLight: 'F8F9FA',
  bgGray: 'E9ECEF',
  textDark: '1B1B1B',
  textMid: '495057',
  white: 'FFFFFF',
};

async function buildDocx(data, hjelpesprak, plassering, grammatikkData) {
  const { yrke, niva, intro, seksjoner, ordliste } = data;
  const showHelp = hjelpesprak && hjelpesprak !== 'ingen';
  const ordlisteAtEnd = showHelp && plassering === 'slutt';

  const border1 = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const allBorders = { top: border1, bottom: border1, left: border1, right: border1 };
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  function hLine() {
    return new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.primary, space: 1 } },
      spacing: { after: 120 },
      children: [],
    });
  }

  function sectionHeader(text) {
    return [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 80 },
        children: [new TextRun({ text, bold: true, size: 40, color: C.primary, font: 'Calibri' })],
      }),
      hLine(),
    ];
  }

  function tekstHeader(nr, tittel) {
    return [
      new Paragraph({
        spacing: { before: 360, after: 0 },
        shading: { fill: C.primary, type: ShadingType.CLEAR },
        children: [new TextRun({ text: `  📄 Tekst ${nr}  `, bold: true, size: 24, color: C.white, font: 'Calibri' })],
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: tittel, bold: true, size: 32, color: C.secondary, font: 'Calibri' })],
      }),
    ];
  }

  function oppgaveHeader(nr, tittel, instruksjon, tilknyttetTekst, oppgavetype) {
    const typeIkon = { leseforståelse: '📖', grammatikk: '✏️', vokabular: '🔤', skriv_muntlig: '💬' }[oppgavetype] || '📝';
    const visTekst = tilknyttetTekst && tilknyttetTekst !== 'Generell';
    return [
      new Paragraph({
        spacing: { before: 300, after: 60 },
        shading: { fill: C.primary, type: ShadingType.CLEAR },
        keepNext: true,
        keepLines: true,
        children: [new TextRun({ text: `  Oppgave ${nr}  `, bold: true, size: 26, color: C.white, font: 'Calibri' })],
      }),
      ...(visTekst ? [new Paragraph({
        spacing: { before: 40, after: 40 },
        shading: { fill: C.secondary, type: ShadingType.CLEAR },
        keepNext: true,
        keepLines: true,
        children: [new TextRun({ text: `  ${typeIkon} ${tilknyttetTekst}  `, size: 20, color: C.white, font: 'Calibri' })],
      })] : []),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        keepNext: true,
        keepLines: true,
        children: [new TextRun({ text: tittel, bold: true, size: 28, color: C.textDark, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { after: 120 },
        keepNext: true,
        keepLines: true,
        children: [new TextRun({ text: instruksjon, italics: true, size: 24, color: C.textMid, font: 'Calibri' })],
      }),
    ];
  }

  function svarLinje() {
    return new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: '_'.repeat(58), size: 22, color: 'AAAAAA', font: 'Calibri' })],
    });
  }

  const titleBlock = [
    new Paragraph({
      shading: { fill: C.primary, type: ShadingType.CLEAR },
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: `  ${yrke.toUpperCase()}  `, bold: true, size: 56, color: C.white, font: 'Calibri' })],
    }),
    new Paragraph({
      shading: { fill: C.secondary, type: ShadingType.CLEAR },
      spacing: { after: 0 },
      children: [
        new TextRun({ text: `  Arbeidshefte – Norsknivå ${niva}`, size: 28, color: C.white, font: 'Calibri' }),
        new TextRun({ text: '   |   Molde voksenopplæringssenter', size: 24, color: C.bgGray, font: 'Calibri' }),
      ],
    }),
    new Paragraph({ spacing: { after: 200 }, children: [] }),
  ];

  const introBlock = [
    ...sectionHeader('Innledning'),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: intro, size: 24, font: 'Calibri' })] }),
  ];

  const seksjonerBlock = [];
  let firstText = true;
  for (const seksjon of seksjoner) {
    if (seksjon.type === 'tekst') {
      if (firstText) { seksjonerBlock.push(...sectionHeader('Fagtekster og oppgaver')); firstText = false; }
      seksjonerBlock.push(...tekstHeader(seksjon.nummer, seksjon.tittel));
      seksjonerBlock.push(new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: seksjon.innhold, size: 24, font: 'Calibri' })] }));
    } else if (seksjon.type === 'oppgave') {
      const deloRows = seksjon.delopgaver.map((d, i) => {
        const fill = i % 2 === 0 ? C.white : C.bgGray;
        return new TableRow({
          cantSplit: true,
          children: [
            new TableCell({ borders: noBorders, width: { size: 800, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: `${d.bokstav})`, bold: true, size: 24, color: C.primary, font: 'Calibri' })] })] }),
            new TableCell({ borders: noBorders, width: { size: 8200, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 60, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d.tekst, size: 24, font: 'Calibri' })] }), svarLinje()] }),
          ],
        });
      });
      seksjonerBlock.push(
        ...oppgaveHeader(seksjon.nummer, seksjon.tittel, seksjon.instruksjon, seksjon.tilknyttet_tekst, seksjon.oppgavetype),
        new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [800, 8200], rows: deloRows }),
        new Paragraph({ spacing: { after: 120 }, children: [] }),
      );
    }
  }

  const colCount = showHelp && !ordlisteAtEnd ? 3 : 2;
  const colWidths = colCount === 3 ? [2700, 3500, 2800] : [3300, 5700];

  function makeHeaderCell(text, w) {
    return new TableCell({ borders: allBorders, width: { size: w, type: WidthType.DXA }, shading: { fill: C.primary, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 22, color: C.white, font: 'Calibri' })] })] });
  }

  const headerCells = [makeHeaderCell('Norsk', colWidths[0]), makeHeaderCell('Forklaring', colWidths[1])];
  if (colCount === 3) headerCells.push(makeHeaderCell(hjelpesprak, colWidths[2]));

  const ordRows = ordliste.map((o, i) => {
    const fill = i % 2 === 0 ? C.white : C.bgGray;
    const mc = (text, w, opts = {}) => new TableCell({ borders: allBorders, width: { size: w, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text, size: 22, font: 'Calibri', ...opts })] })] });
    const cells = [mc(o.norsk, colWidths[0], { bold: true, color: C.secondary }), mc(o.forklaring, colWidths[1])];
    if (colCount === 3) cells.push(mc(o.oversettelse || '', colWidths[2], { italics: true }));
    return new TableRow({ children: cells });
  });

  const ordlisteBlock = [
    ...sectionHeader('Viktige ord og uttrykk'),
    new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: colWidths, rows: [new TableRow({ children: headerCells }), ...ordRows] }),
    new Paragraph({ spacing: { after: 200 }, children: [] }),
  ];

  const extraOrdliste = ordlisteAtEnd ? [
    ...sectionHeader(`Ordliste – ${hjelpesprak}`),
    new Table({
      width: { size: 9000, type: WidthType.DXA }, columnWidths: [4500, 4500],
      rows: [
        new TableRow({ children: [makeHeaderCell('Norsk', 4500), makeHeaderCell(hjelpesprak, 4500)] }),
        ...ordliste.map((o, i) => {
          const fill = i % 2 === 0 ? C.white : C.bgGray;
          return new TableRow({ children: [
            new TableCell({ borders: allBorders, width: { size: 4500, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: o.norsk, size: 22, bold: true, color: C.secondary, font: 'Calibri' })] })] }),
            new TableCell({ borders: allBorders, width: { size: 4500, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: o.oversettelse || '', size: 22, italics: true, font: 'Calibri' })] })] }),
          ] });
        }),
      ],
    }),
  ] : [];

  const grammatikkBlock = [];
  if (grammatikkData && grammatikkData.oppgaver) {
    grammatikkBlock.push(...sectionHeader(`Grammatikk: ${grammatikkData.tema}`));
    grammatikkBlock.push(
      new Paragraph({ spacing: { before: 0, after: 0 }, shading: { fill: C.primary, type: ShadingType.CLEAR }, children: [new TextRun({ text: '  📘 Grammatikkforklaring  ', bold: true, size: 24, color: C.white, font: 'Calibri' })] }),
      new Paragraph({ spacing: { before: 0, after: 240 }, shading: { fill: 'E6F4F6', type: ShadingType.CLEAR }, border: { left: { style: BorderStyle.SINGLE, size: 12, color: C.secondary, space: 8 } }, children: [new TextRun({ text: grammatikkData.forklaring, size: 24, font: 'Calibri', color: C.textDark })] })
    );
  }

  const doc = new Document({
    numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    styles: {
      default: { document: { run: { font: 'Calibri', size: 24 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 40, bold: true, font: 'Calibri', color: C.primary }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, font: 'Calibri', color: C.secondary }, paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
      headers: { default: new Header({ children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.primary, space: 1 } }, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }], children: [new TextRun({ text: `${yrke} – Nivå ${niva}`, size: 18, color: C.textMid, font: 'Calibri' }), new TextRun({ text: '\t', size: 18 }), new TextRun({ text: 'Molde voksenopplæringssenter', size: 18, color: C.textMid, font: 'Calibri' })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.bgGray, space: 1 } }, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }], children: [new TextRun({ text: '© MBO – Molde voksenopplæringssenter', size: 18, color: C.textMid, font: 'Calibri' }), new TextRun({ text: '\tSide ', size: 18, color: C.textMid, font: 'Calibri' }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.textMid, font: 'Calibri' })] })] }) },
      children: [...titleBlock, ...introBlock, ...seksjonerBlock, ...grammatikkBlock, ...ordlisteBlock, ...extraOrdliste],
    }],
  });

  return Packer.toBuffer(doc);
}

async function buildPptx(data, yrke, niva) {
  const { hms = [], egenskaper = [], arbeidsoppgaver = [], utdanning = [] } = data.pptx || {};
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = `${yrke} – Norsknivå ${niva}`;
  pres.author = 'Molde voksenopplæringssenter';

  const s1 = pres.addSlide();
  s1.background = { color: C.primary };
  s1.addText(yrke.toUpperCase(), { x: 0.5, y: 1.3, w: 9, h: 1.5, fontSize: 44, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle' });
  s1.addText(`Norsknivå ${niva}`, { x: 0.5, y: 2.85, w: 9, h: 0.5, fontSize: 20, color: C.accent, fontFace: 'Calibri', align: 'center', valign: 'middle' });

  const s2 = pres.addSlide();
  s2.background = { color: C.bgLight };
  s2.addText('Hva er dette yrket?', { x: 0.3, y: 0.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.primary, fontFace: 'Calibri' });
  s2.addText(arbeidsoppgaver.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })), { x: 0.4, y: 1.1, w: 8.5, h: 3.5, fontSize: 15, color: C.textDark, fontFace: 'Calibri' });

  const s3 = pres.addSlide();
  s3.background = { color: C.bgLight };
  s3.addText('HMS', { x: 0.3, y: 0.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.primary, fontFace: 'Calibri' });
  s3.addText(hms.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })), { x: 0.4, y: 1.1, w: 8.5, h: 3.5, fontSize: 15, color: C.textDark, fontFace: 'Calibri' });

  const s4 = pres.addSlide();
  s4.background = { color: C.bgLight };
  s4.addText('Personlige egenskaper', { x: 0.3, y: 0.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.primary, fontFace: 'Calibri' });
  s4.addText(egenskaper.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })), { x: 0.4, y: 1.1, w: 8.5, h: 3.5, fontSize: 15, color: C.textDark, fontFace: 'Calibri' });

  const s5 = pres.addSlide();
  s5.background = { color: C.bgLight };
  s5.addText('Utdanning og karriere', { x: 0.3, y: 0.2, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.primary, fontFace: 'Calibri' });
  s5.addText(utdanning.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })), { x: 0.4, y: 1.1, w: 8.5, h: 3.5, fontSize: 15, color: C.textDark, fontFace: 'Calibri' });

  const tmpPath = path.join(os.tmpdir(), `pptx-${Date.now()}.pptx`);
  try {
    await pres.writeFile({ fileName: tmpPath });
    return fs.readFileSync(tmpPath);
  } finally {
    if (fs.existsSync(tmpPath)) fs.unlinkSync(tmpPath);
  }
}

module.exports = {
  buildDocx,
  buildPptx,
};
