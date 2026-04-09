'use strict';
const express = require('express');
const https = require('https');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, TabStopType, TabStopPosition, PageNumber, Header, Footer,
} = require('docx');
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── Colours from DESIGN.md ───────────────────────────────────────────────────
const C = {
  primary:   '005F73',
  secondary: '0A9396',
  accent:    'E9C46A',
  bgLight:   'F8F9FA',
  bgGray:    'E9ECEF',
  textDark:  '1B1B1B',
  textMid:   '495057',
  white:     'FFFFFF',
};

// ─── Gemini 2.5 Flash via direkte HTTP ────────────────────────────────────────
function callGemini(prompt) {
  return new Promise((resolve, reject) => {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return reject(new Error('GEMINI_API_KEY mangler i miljøvariabler.'));

    const body = JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 8192 }
    });

    const options = {
      hostname: 'generativelanguage.googleapis.com',
      path: `/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.error) return reject(new Error(parsed.error.message));
          const text = parsed.candidates[0].content.parts[0].text;
          resolve(text);
        } catch (e) {
          reject(new Error('Kunne ikke tolke svar fra Gemini: ' + data.slice(0, 200)));
        }
      });
    });

    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─── Prompt ────────────────────────────────────────────────────────────────────
function buildPrompt(yrke, niva, sprak, plassering) {
  const nivaMap = {
    A1: 'svært enkelt språk, korte setninger maks 40 ord per tekst, grunnleggende hverdagsord',
    A2: 'enkelt språk, korte avsnitt maks 70 ord per tekst, vanlige arbeidslivsord',
    B1: 'moderat komplekst, ca. 120 ord per tekst, fagterminologi forklares i teksten',
    B2: 'mer komplekst, ca. 180 ord per tekst, fagterminologi brukes naturlig',
  };
  const hjelpeTekst = sprak && sprak !== 'ingen'
    ? `Hjelpespråk: ${sprak}. Plassering: ${plassering === 'tosidig'
        ? 'oversettelse i parentes etter norsk term, f.eks. lege (doctor)'
        : 'samle alle oversettelser i ordlisten på slutten'}.`
    : 'Ingen hjelpespråk.';

  return `Du er en erfaren norsklærer og fagpedagog. Lag et komplett arbeidshefte om yrket "${yrke}" på norsknivå ${niva} (${nivaMap[niva]}).

${hjelpeTekst}

VIKTIG: Svar KUN med gyldig JSON. Ingen markdown, ingen tekst utenfor JSON.

{
  "yrke": "${yrke}",
  "niva": "${niva}",
  "intro": "2-3 setninger om yrket på ${niva}-nivå",
  "tekster": [
    { "tittel": "Tittel 1", "innhold": "Tekst 1" },
    { "tittel": "Tittel 2", "innhold": "Tekst 2" },
    { "tittel": "Tittel 3", "innhold": "Tekst 3" }
  ],
  "ordliste": [
    { "norsk": "fagord", "forklaring": "norsk forklaring"${sprak && sprak !== 'ingen' ? ', "oversettelse": "på ' + sprak + '"' : ''} }
  ],
  "oppgaver": [
    {
      "nummer": 1,
      "type": "leseforståelse",
      "tittel": "Oppgavetittel",
      "instruksjon": "Les teksten og svar på spørsmålene.",
      "delopgaver": [
        { "bokstav": "a", "tekst": "Spørsmål..." },
        { "bokstav": "b", "tekst": "..." },
        { "bokstav": "c", "tekst": "..." },
        { "bokstav": "d", "tekst": "..." },
        { "bokstav": "e", "tekst": "..." }
      ]
    }
  ],
  "pptx": {
    "nokkelord": ["8 viktige fagord for yrket"],
    "hms": ["4 viktige HMS-punkter"],
    "egenskaper": ["5 personlige egenskaper"],
    "arbeidsoppgaver": ["4 typiske arbeidsoppgaver"],
    "utdanning": ["3 utdanningsveier eller karrieremuligheter"]
  }
}

KRAV TIL OPPGAVENE (10-13 stk.):
- Leseforståelse: 3 oppgaver (spørsmål til tekstene)
- Vokabular/ord: 2-3 oppgaver (ordforklaring, fyll inn riktig ord, koble ord og forklaring)
- Grammatikk: 3-4 oppgaver knyttet til yrket (verb, setningsbygning V2, adjektiv, preposisjoner)
- Muntlig/skriv: 1-2 oppgaver
- Alle oppgaver har 5 delopgaver a-e
- Ordlisten: 10-15 ord
Lag faglig korrekt, pedagogisk og engasjerende innhold.`;
}

// ─── DOCX builder ──────────────────────────────────────────────────────────────
async function buildDocx(data, hjelpesprak, plassering) {
  const { yrke, niva, intro, tekster, ordliste, oppgaver } = data;
  const showHelp = hjelpesprak && hjelpesprak !== 'ingen';
  const ordlisteAtEnd = showHelp && plassering === 'slutt';

  const border1 = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const allBorders = { top: border1, bottom: border1, left: border1, right: border1 };
  const noBorder  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
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

  function oppgaveHeader(nr, tittel, instruksjon) {
    return [
      new Paragraph({
        spacing: { before: 300, after: 60 },
        shading: { fill: C.primary, type: ShadingType.CLEAR },
        children: [new TextRun({ text: `  Oppgave ${nr}  `, bold: true, size: 26, color: C.white, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: tittel, bold: true, size: 28, color: C.textDark, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { after: 120 },
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
    new Paragraph({ spacing: { after: 240 }, children: [] }),
  ];

  const introBlock = [
    ...sectionHeader('Innledning'),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({ text: intro, size: 24, font: 'Calibri' })],
    }),
  ];

  const teksterBlock = [
    ...sectionHeader('Yrkestekster'),
    ...tekster.flatMap(t => [
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 80 },
        children: [new TextRun({ text: t.tittel, bold: true, size: 32, color: C.secondary, font: 'Calibri' })],
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun({ text: t.innhold, size: 24, font: 'Calibri' })],
      }),
    ]),
  ];

  const colCount = showHelp && !ordlisteAtEnd ? 3 : 2;
  const colWidths = colCount === 3 ? [2700, 3500, 2800] : [3300, 5700];

  function makeHeaderCell(text, w) {
    return new TableCell({
      borders: allBorders,
      width: { size: w, type: WidthType.DXA },
      shading: { fill: C.primary, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 22, color: C.white, font: 'Calibri' })] })],
    });
  }

  const headerCells = [makeHeaderCell('Norsk', colWidths[0]), makeHeaderCell('Forklaring', colWidths[1])];
  if (colCount === 3) headerCells.push(makeHeaderCell(hjelpesprak, colWidths[2]));

  const ordRows = ordliste.map((o, i) => {
    const fill = i % 2 === 0 ? C.white : C.bgGray;
    const makeCell = (text, w, opts = {}) => new TableCell({
      borders: allBorders, width: { size: w, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text, size: 22, font: 'Calibri', ...opts })] })],
    });
    const cells = [
      makeCell(o.norsk, colWidths[0], { bold: true, color: C.secondary }),
      makeCell(o.forklaring, colWidths[1]),
    ];
    if (colCount === 3) cells.push(makeCell(o.oversettelse || '', colWidths[2], { italics: true }));
    return new TableRow({ children: cells });
  });

  const ordlisteBlock = [
    ...sectionHeader('Viktige ord og uttrykk'),
    new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: colWidths, rows: [new TableRow({ children: headerCells }), ...ordRows] }),
    new Paragraph({ spacing: { after: 200 }, children: [] }),
  ];

  const oppgaverBlock = [
    ...sectionHeader('Oppgaver'),
    ...oppgaver.flatMap(o => {
      const deloRows = o.delopgaver.map((d, i) => {
        const fill = i % 2 === 0 ? C.white : C.bgGray;
        return new TableRow({
          children: [
            new TableCell({
              borders: noBorders, width: { size: 800, type: WidthType.DXA },
              shading: { fill, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 60 },
              children: [new Paragraph({ children: [new TextRun({ text: `${d.bokstav})`, bold: true, size: 24, color: C.primary, font: 'Calibri' })] })],
            }),
            new TableCell({
              borders: noBorders, width: { size: 8200, type: WidthType.DXA },
              shading: { fill, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 60, right: 120 },
              children: [
                new Paragraph({ children: [new TextRun({ text: d.tekst, size: 24, font: 'Calibri' })] }),
                svarLinje(),
              ],
            }),
          ],
        });
      });
      return [
        ...oppgaveHeader(o.nummer, o.tittel, o.instruksjon),
        new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [800, 8200], rows: deloRows }),
        new Paragraph({ spacing: { after: 120 }, children: [] }),
      ];
    }),
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
          ]});
        }),
      ],
    }),
  ] : [];

  const doc = new Document({
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      }],
    },
    styles: {
      default: { document: { run: { font: 'Calibri', size: 24 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 40, bold: true, font: 'Calibri', color: C.primary },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Calibri', color: C.secondary },
          paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.primary, space: 1 } },
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            children: [
              new TextRun({ text: `${yrke} – Nivå ${niva}`, size: 18, color: C.textMid, font: 'Calibri' }),
              new TextRun({ text: '\t', size: 18 }),
              new TextRun({ text: 'Molde voksenopplæringssenter', size: 18, color: C.textMid, font: 'Calibri' }),
            ],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.bgGray, space: 1 } },
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            children: [
              new TextRun({ text: '© MBO – Molde voksenopplæringssenter', size: 18, color: C.textMid, font: 'Calibri' }),
              new TextRun({ text: '\tSide ', size: 18, color: C.textMid, font: 'Calibri' }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.textMid, font: 'Calibri' }),
            ],
          })],
        }),
      },
      children: [
        ...titleBlock,
        ...introBlock,
        ...teksterBlock,
        ...ordlisteBlock,
        ...oppgaverBlock,
        ...extraOrdliste,
      ],
    }],
  });

  return Packer.toBuffer(doc);
}

// ─── PPTX builder ──────────────────────────────────────────────────────────────
async function buildPptx(data, yrke, niva) {
  const { nokkelord, hms, egenskaper, arbeidsoppgaver, utdanning } = data.pptx;

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = `${yrke} – Norsknivå ${niva}`;
  pres.author = 'Molde voksenopplæringssenter';

  const makeShadow = () => ({ type: 'outer', blur: 6, offset: 2, angle: 135, color: '000000', opacity: 0.12 });

  function darkSlide() {
    const s = pres.addSlide();
    s.background = { color: C.primary };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.18, fill: { color: C.accent }, line: { color: C.accent } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.445, w: 10, h: 0.18, fill: { color: C.accent }, line: { color: C.accent } });
    return s;
  }

  function lightSlide(titleText) {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.primary }, line: { color: C.primary } });
    s.addText(titleText, { x: 0.3, y: 0.2, w: 9.2, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: 'Calibri', align: 'left', margin: 0 });
    s.addShape(pres.shapes.LINE, { x: 0.3, y: 1.0, w: 9.3, h: 0, line: { color: C.secondary, width: 2 } });
    return s;
  }

  // Slide 1 – Tittel
  {
    const s = darkSlide();
    s.addText(yrke.toUpperCase(), { x: 0.5, y: 1.5, w: 9, h: 1.5, fontSize: 44, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', margin: 0 });
    s.addText(`Norsknivå ${niva}  •  Molde voksenopplæringssenter`, { x: 0.5, y: 3.3, w: 9, h: 0.6, fontSize: 18, color: C.bgGray, fontFace: 'Calibri', align: 'center', margin: 0 });
  }

  // Slide 2 – Om yrket
  {
    const s = lightSlide('Om yrket');
    const items = arbeidsoppgaver.map((t, idx) => ({
      text: t,
      options: { bullet: true, breakLine: idx < arbeidsoppgaver.length - 1, fontSize: 16, color: C.textDark, fontFace: 'Calibri', paraSpaceAfter: 10 },
    }));
    s.addText(items, { x: 0.4, y: 1.2, w: 5.8, h: 3.8, valign: 'top' });
    s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 1.2, w: 2.6, h: 3.8, fill: { color: C.primary }, line: { color: C.primary }, shadow: makeShadow() });
    s.addText(yrke.charAt(0).toUpperCase(), { x: 7.0, y: 1.2, w: 2.6, h: 3.8, fontSize: 100, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
  }

  // Slide 3 – Viktige ord
  {
    const s = lightSlide('Viktige ord og uttrykk');
    nokkelord.slice(0, 8).forEach((ord, i) => {
      const col = i % 4;
      const row = Math.floor(i / 4);
      const x = 0.3 + col * 2.35;
      const y = 1.15 + row * 1.85;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.2, h: 1.55, fill: { color: C.white }, line: { color: C.bgGray, width: 1 }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.09, h: 1.55, fill: { color: C.secondary }, line: { color: C.secondary } });
      s.addText(ord, { x: x + 0.15, y, w: 2.0, h: 1.55, fontSize: 13, bold: true, color: C.textDark, fontFace: 'Calibri', align: 'left', valign: 'middle', wrap: true, margin: 4 });
    });
  }

  // Slide 4 – HMS
  {
    const s = lightSlide('Helse, miljø og sikkerhet (HMS)');
    s.addText('HMS', { x: 5.2, y: 0.9, w: 4.5, h: 4.0, fontSize: 110, bold: true, color: C.bgGray, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
    const dotColors = [C.primary, C.secondary, 'E9C46A', '264653'];
    hms.slice(0, 4).forEach((punkt, i) => {
      const y = 1.15 + i * 1.05;
      s.addShape(pres.shapes.OVAL, { x: 0.4, y, w: 0.55, h: 0.55, fill: { color: dotColors[i] }, line: { color: dotColors[i] } });
      s.addText(String(i + 1), { x: 0.4, y, w: 0.55, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0 });
      s.addText(punkt, { x: 1.1, y: y + 0.03, w: 4.8, h: 0.5, fontSize: 15, color: C.textDark, fontFace: 'Calibri', align: 'left', valign: 'middle', margin: 0 });
    });
  }

  // Slide 5 – Personlige egenskaper
  {
    const s = lightSlide('Personlige egenskaper');
    egenskaper.slice(0, 5).forEach((eg, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = 0.3 + col * 3.15;
      const y = 1.15 + row * 1.8;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 1.55, fill: { color: C.bgGray }, line: { color: C.bgGray } });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: 1.55, fill: { color: C.secondary }, line: { color: C.secondary } });
      s.addText(eg, { x: x + 0.2, y, w: 2.7, h: 1.55, fontSize: 14, bold: true, color: C.textDark, fontFace: 'Calibri', align: 'left', valign: 'middle', wrap: true, margin: 4 });
    });
  }

  // Slide 6 – Utdanning (mørk)
  {
    const s = darkSlide();
    s.addText('Utdanning og karriere', { x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 28, bold: true, color: C.white, fontFace: 'Calibri', align: 'left', margin: 0 });
    s.addShape(pres.shapes.LINE, { x: 0.5, y: 1.1, w: 9, h: 0, line: { color: C.accent, width: 2 } });
    utdanning.slice(0, 3).forEach((u, i) => {
      const x = 0.5 + i * 3.1;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.9, h: 3.7, fill: { color: C.white, transparency: 85 }, line: { color: C.white, width: 1 } });
      s.addText(String(i + 1), { x, y: 1.4, w: 2.9, h: 0.7, fontSize: 32, bold: true, color: C.accent, fontFace: 'Calibri', align: 'center', margin: 0 });
      s.addText(u, { x: x + 0.1, y: 2.2, w: 2.7, h: 2.5, fontSize: 14, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'top', wrap: true, margin: 4 });
    });
  }

  // Slide 7 – Avslutning
  {
    const s = darkSlide();
    s.addText('Lykke til!', { x: 0.5, y: 1.5, w: 9, h: 1.5, fontSize: 54, bold: true, color: C.white, fontFace: 'Calibri', align: 'center', margin: 0 });
    s.addText(`${yrke}  •  Norsknivå ${niva}`, { x: 0.5, y: 3.2, w: 9, h: 0.6, fontSize: 20, color: C.accent, fontFace: 'Calibri', align: 'center', margin: 0 });
    s.addText('Molde voksenopplæringssenter – MBO', { x: 0.5, y: 3.9, w: 9, h: 0.5, fontSize: 14, color: C.bgGray, fontFace: 'Calibri', align: 'center', margin: 0 });
  }

  const tmpPath = `/tmp/pptx-${Date.now()}.pptx`;
  await pres.writeFile({ fileName: tmpPath });
  const buf = fs.readFileSync(tmpPath);
  fs.unlinkSync(tmpPath);
  return buf;
}

// ─── API endpoint ──────────────────────────────────────────────────────────────
app.post('/api/generer', async (req, res) => {
  try {
    const { yrke, niva, sprak, plassering } = req.body;
    if (!yrke || !niva) return res.status(400).json({ feil: 'Yrke og nivå er påkrevd.' });

    const raw = await callGemini(buildPrompt(yrke, niva, sprak, plassering));
    const clean = raw.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '');

    let data;
    try {
      data = JSON.parse(clean);
    } catch (e) {
      console.error('JSON feil:', clean.slice(0, 400));
      return res.status(500).json({ feil: 'Klarte ikke tolke svar fra AI. Prøv igjen.' });
    }

    const [docxBuf, pptxBuf] = await Promise.all([
      buildDocx(data, sprak, plassering),
      buildPptx(data, yrke, niva),
    ]);

    const safeName = yrke.replace(/[^a-zA-ZæøåÆØÅ0-9\-]/g, '_');
    res.json({
      docx: docxBuf.toString('base64'),
      pptx: pptxBuf.toString('base64'),
      filnavn: safeName,
      niva: niva,
    });

  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).json({ feil: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Yrkesappen kjører på port ${PORT}`));
