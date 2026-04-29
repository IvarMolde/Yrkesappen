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

function buildInteractiveHtml(data, yrke, niva, hjelpesprak) {
  const tekster = (data.seksjoner || []).filter((s) => s.type === 'tekst').slice(0, 3);
  const ordliste = Array.isArray(data.ordliste) ? data.ordliste : [];
  const showHelp = hjelpesprak && hjelpesprak !== 'ingen';

  function escapeHtml(text) {
    return String(text || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function splitSentences(text) {
    return String(text || '')
      .split(/(?<=[.!?])\s+/)
      .map((s) => s.trim())
      .filter((s) => s.length > 10);
  }

  function words(text) {
    return String(text || '').match(/[A-Za-zÆØÅæøå]+/g) || [];
  }

  function seededPick(arr, seed, count) {
    const copy = [...arr];
    let x = seed % 2147483647;
    if (x <= 0) x += 2147483646;
    function rnd() {
      x = (x * 16807) % 2147483647;
      return (x - 1) / 2147483646;
    }
    for (let i = copy.length - 1; i > 0; i -= 1) {
      const j = Math.floor(rnd() * (i + 1));
      [copy[i], copy[j]] = [copy[j], copy[i]];
    }
    return copy.slice(0, count);
  }

  const antonymPairs = [
    ['høy', 'lav'], ['stor', 'liten'], ['trygg', 'farlig'], ['tidlig', 'sent'], ['rask', 'langsom'],
    ['sterk', 'svak'], ['ren', 'skitten'], ['åpen', 'lukket'], ['viktig', 'uviktig'], ['rolig', 'stresset'],
  ];

  const taskTypePool = ['fill_missing', 'true_false', 'synonym', 'antonym', 'choose_statement', 'click_word'];
  if (niva !== 'A1') taskTypePool.push('sentence_order');

  function buildTaskByType(type, textObj, textIdx) {
    const sentences = splitSentences(textObj.innhold);
    const vocabWords = ordliste.map((o) => o.norsk).filter(Boolean);

    if (type === 'fill_missing') {
      const items = seededPick(sentences, textIdx + 11, 5).map((sentence, i) => {
        const sentenceWords = words(sentence).filter((w) => w.length > 3);
        const answer = sentenceWords[(i + sentenceWords.length) % sentenceWords.length] || 'ord';
        const prompt = sentence.replace(new RegExp(`\\b${answer}\\b`), '_____');
        return { letter: 'abcde'[i], sentence: prompt, answer };
      });
      return { type, title: 'Skriv inn ordet som mangler', instruction: 'Skriv ordet som mangler i hver setning.', items };
    }

    if (type === 'sentence_order') {
      const items = seededPick(sentences, textIdx + 21, 5).map((sentence, i) => {
        const w = words(sentence).slice(0, 8);
        const shuffled = seededPick(w, textIdx + 100 + i, w.length);
        return { letter: 'abcde'[i], shuffled: shuffled.join(' / '), answer: sentence };
      });
      return { type, title: 'Sorter ordene til riktig setning', instruction: 'Skriv setningen i riktig rekkefølge. Husk inversjon når det trengs.', items };
    }

    if (type === 'true_false') {
      const picked = seededPick(sentences, textIdx + 31, 5);
      const items = picked.map((sentence, i) => {
        const isTrue = i % 2 === 0;
        const statement = isTrue ? sentence : `Ikke: ${sentence}`;
        return { letter: 'abcde'[i], statement, answer: isTrue ? 'sant' : 'usant' };
      });
      return { type, title: 'Sant eller usant', instruction: 'Velg om påstanden er sant eller usant.', items };
    }

    if (type === 'synonym') {
      const candidates = ordliste.filter((o) => o.forklaring && o.norsk).slice(0, 8);
      const items = seededPick(candidates, textIdx + 41, Math.min(5, candidates.length)).map((entry, i) => {
        const options = seededPick(candidates.map((c) => c.norsk), textIdx + 200 + i, Math.min(4, candidates.length));
        if (!options.includes(entry.norsk)) options[0] = entry.norsk;
        return { letter: 'abcde'[i], clue: entry.forklaring, options, answer: entry.norsk };
      });
      while (items.length < 5) {
        items.push({ letter: 'abcde'[items.length], clue: 'Velg ordet som passer best.', options: vocabWords.slice(0, 4), answer: vocabWords[0] || 'ord' });
      }
      return { type, title: 'Finn synonym/ord med samme mening', instruction: 'Velg ordet som passer til forklaringen.', items };
    }

    if (type === 'antonym') {
      const usable = antonymPairs.filter(([a]) => vocabWords.some((w) => w.toLowerCase().includes(a)));
      const base = usable.length > 0 ? usable : antonymPairs;
      const items = seededPick(base, textIdx + 51, 5).map((pair, i) => {
        const [a, b] = pair;
        const options = seededPick(antonymPairs.map((p) => p[1]), textIdx + 300 + i, 4);
        if (!options.includes(b)) options[0] = b;
        return { letter: 'abcde'[i], word: a, options, answer: b };
      });
      return { type, title: 'Finn antonym', instruction: 'Velg ordet som betyr det motsatte.', items };
    }

    if (type === 'click_word') {
      const items = seededPick(sentences, textIdx + 61, 5).map((sentence, i) => {
        const ws = words(sentence).filter((w) => w.length > 3);
        const answer = ws[(i + 1) % ws.length] || ws[0] || 'ord';
        const clue = (ordliste.find((o) => o.norsk.toLowerCase().includes(answer.toLowerCase())) || {}).forklaring || `Klikk på ordet som betyr: ${answer}`;
        return { letter: 'abcde'[i], sentence, clue, answer };
      });
      return { type, title: 'Klikk på riktig ord i teksten', instruction: 'Klikk ordet som passer til betydningen.', items };
    }

    const items = seededPick(sentences, textIdx + 71, 5).map((sentence, i) => {
      const wrongA = `Påstand A: ${sentence}`;
      const wrongB = `Påstand B: Teksten handler ikke om ${yrke}.`;
      const correct = `Påstand C: ${sentence}`;
      return { letter: 'abcde'[i], options: [wrongA, wrongB, correct], answer: correct };
    });
    return { type: 'choose_statement', title: 'Finn riktig påstand', instruction: 'Velg påstanden som stemmer med teksten.', items };
  }

  const textModels = tekster.map((t, i) => {
    const chosenTypes = seededPick(taskTypePool, 700 + i * 17, Math.min(3, taskTypePool.length));
    const tasks = chosenTypes.map((type) => buildTaskByType(type, t, i));
    return { title: t.tittel, content: t.innhold, tasks };
  });

  const wizardTabs = textModels.map((t, i) => `<button class="tab-btn${i === 0 ? ' active' : ''}" data-pane="pane-${i}">Tekst ${i + 1}</button>`).join('');
  const glossary = ordliste.slice(0, 16).map((o) => `<tr><td>${escapeHtml(o.norsk)}</td><td>${escapeHtml(o.forklaring)}</td>${showHelp ? `<td>${escapeHtml(o.oversettelse || '')}</td>` : ''}</tr>`).join('');

  function renderTask(task, textIdx, taskIdx) {
    const taskKey = `t${textIdx}-q${taskIdx}`;
    const rows = task.items.map((item, i) => {
      const id = `${taskKey}-${item.letter}`;
      if (task.type === 'fill_missing' || task.type === 'sentence_order') {
        return `<div class="item"><label><strong>${item.letter})</strong> ${escapeHtml(item.sentence || item.shuffled)}</label><div class="answer-row"><input id="${id}" type="text"><button onclick="checkText('${id}','${escapeHtml(item.answer.toLowerCase())}')">Sjekk</button><span id="${id}-fb" class="fb"></span></div></div>`;
      }
      if (task.type === 'true_false') {
        return `<div class="item"><div><strong>${item.letter})</strong> ${escapeHtml(item.statement)}</div><div class="answer-row"><button onclick="checkChoice('${id}','sant','${item.answer}')">Sant</button><button onclick="checkChoice('${id}','usant','${item.answer}')">Usant</button><span id="${id}-fb" class="fb"></span></div></div>`;
      }
      if (task.type === 'click_word') {
        const buttons = words(item.sentence).slice(0, 10).map((w) => `<button onclick="checkChoice('${id}','${escapeHtml(w.toLowerCase())}','${escapeHtml(item.answer.toLowerCase())}')">${escapeHtml(w)}</button>`).join('');
        return `<div class="item"><div><strong>${item.letter})</strong> ${escapeHtml(item.clue)}</div><p class="mini">${escapeHtml(item.sentence)}</p><div class="answer-row">${buttons}<span id="${id}-fb" class="fb"></span></div></div>`;
      }
      const opts = (item.options || []).map((op) => `<button onclick="checkChoice('${id}','${escapeHtml(op)}','${escapeHtml(item.answer)}')">${escapeHtml(op)}</button>`).join('');
      const prompt = item.clue || `Finn antonym til ordet: ${item.word || ''}`;
      return `<div class="item"><div><strong>${item.letter})</strong> ${escapeHtml(prompt)}</div><div class="answer-row">${opts}<span id="${id}-fb" class="fb"></span></div></div>`;
    }).join('');
    return `<section class="task"><h4>${escapeHtml(task.title)}</h4><p class="instr">${escapeHtml(task.instruction)}</p>${rows}</section>`;
  }

  const panes = textModels.map((t, i) => `
    <section id="pane-${i}" class="pane${i === 0 ? ' active' : ''}">
      <h2>${escapeHtml(t.title)}</h2>
      <article class="text-card">${escapeHtml(t.content)}</article>
      <div class="tasks">
        ${t.tasks.map((task, idx) => renderTask(task, i, idx)).join('')}
      </div>
    </section>
  `).join('');

  const html = `<!doctype html>
<html lang="no">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(yrke)} - interaktive oppgaver (${escapeHtml(niva)})</title>
  <style>
    body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f6f8fa;color:#1b1b1b}
    header{background:#005F73;color:#fff;padding:12px 18px}
    .wrap{max-width:1100px;margin:0 auto;padding:14px}
    .tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}
    .tab-btn{border:1px solid #c9d1d9;background:#fff;padding:8px 12px;border-radius:8px;cursor:pointer}
    .tab-btn.active{background:#005F73;color:#fff;border-color:#005F73}
    .layout{display:grid;grid-template-columns:2.1fr 1fr;gap:12px}
    .pane{display:none}.pane.active{display:block}
    .text-card{background:#fff;border-radius:10px;padding:12px;border:1px solid #d0d7de;line-height:1.5}
    .task{background:#fff;border:1px solid #d0d7de;border-radius:10px;padding:10px;margin-top:10px}
    .instr{color:#495057;margin-top:0}
    .item{border-top:1px dashed #e5e7eb;padding-top:8px;margin-top:8px}
    .answer-row{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:6px}
    button{background:#0A9396;color:#fff;border:none;border-radius:6px;padding:6px 9px;cursor:pointer}
    input{padding:6px 8px;border:1px solid #c9d1d9;border-radius:6px;min-width:220px}
    .fb{font-weight:700}
    .ok{color:#12733f}.no{color:#b42318}
    .side{background:#fff;border:1px solid #d0d7de;border-radius:10px;padding:10px;position:sticky;top:10px;max-height:86vh;overflow:auto}
    table{width:100%;border-collapse:collapse}th,td{border-bottom:1px solid #eee;padding:6px;text-align:left;vertical-align:top}
    .mini{font-size:13px;color:#495057}
    @media (max-width:960px){.layout{grid-template-columns:1fr}.side{position:static;max-height:none}}
  </style>
</head>
<body>
  <header><h1>${escapeHtml(yrke)} - Interaktive oppgaver (${escapeHtml(niva)})</h1></header>
  <div class="wrap">
    <div class="tabs">${wizardTabs}</div>
    <div class="layout">
      <main>${panes}</main>
      <aside class="side">
        <h3>Ordliste</h3>
        <table>
          <thead><tr><th>Norsk</th><th>Forklaring</th>${showHelp ? `<th>${escapeHtml(hjelpesprak)}</th>` : ''}</tr></thead>
          <tbody>${glossary}</tbody>
        </table>
      </aside>
    </div>
  </div>
  <script>
    document.querySelectorAll('.tab-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach((b) => b.classList.remove('active'));
        document.querySelectorAll('.pane').forEach((p) => p.classList.remove('active'));
        btn.classList.add('active');
        const pane = document.getElementById(btn.dataset.pane);
        if (pane) pane.classList.add('active');
      });
    });
    function mark(id, ok) {
      const el = document.getElementById(id + '-fb');
      if (!el) return;
      el.className = 'fb ' + (ok ? 'ok' : 'no');
      el.textContent = ok ? 'Riktig' : 'Feil';
    }
    window.checkText = function(id, expected) {
      const inp = document.getElementById(id);
      const got = (inp && inp.value ? inp.value : '').trim().toLowerCase();
      mark(id, got === String(expected).trim().toLowerCase());
    };
    window.checkChoice = function(id, selected, expected) {
      mark(id, String(selected).trim().toLowerCase() === String(expected).trim().toLowerCase());
    };
  </script>
</body>
</html>`;

  return Buffer.from(html, 'utf8');
}

module.exports = {
  buildDocx,
  buildPptx,
  buildInteractiveHtml,
};
