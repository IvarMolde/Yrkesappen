'use strict';

const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Yrkesappen';
pptx.company = 'Molde voksenopplæringssenter';
pptx.subject = 'Produktpresentasjon';
pptx.title = 'Yrkesappen – pedagogisk og teknisk oversikt';

const C = {
  primary: '005F73',
  secondary: '0A9396',
  accent: 'E9C46A',
  dark: '1B1B1B',
  mid: '495057',
  light: 'F8F9FA',
  white: 'FFFFFF',
};

function header(slide, title, subtitle) {
  slide.background = { color: C.light };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.3, fill: { color: C.primary }, line: { color: C.primary } });
  slide.addText(title, {
    x: 0.5, y: 0.45, w: 12.2, h: 0.55,
    fontFace: 'Calibri', fontSize: 28, bold: true, color: C.primary,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.5, y: 1.0, w: 12.2, h: 0.35,
      fontFace: 'Calibri', fontSize: 14, italic: true, color: C.mid,
    });
  }
}

function bullets(slide, items, x = 0.7, y = 1.6, w = 8.6, h = 4.7) {
  slide.addText(
    items.map((t, i) => ({
      text: t,
      options: {
        bullet: { indent: 18 },
        breakLine: i < items.length - 1,
        paraSpaceAfterPt: 10,
      },
    })),
    {
      x, y, w, h,
      fontFace: 'Calibri', fontSize: 20, color: C.dark, valign: 'top', fit: 'shrink',
    }
  );
}

function addNotes(slide, lines) {
  if (!Array.isArray(lines) || lines.length === 0) return;
  const notes = lines.map((line) => (line.startsWith('- ') ? line : `- ${line}`));
  slide.addNotes(notes);
}

// 1: Tittel
{
  const s = pptx.addSlide();
  s.background = { color: C.primary };
  s.addShape(pptx.ShapeType.rect, { x: 0.8, y: 0.9, w: 11.8, h: 4.8, fill: { color: 'FFFFFF', transparency: 90 }, line: { color: C.white } });
  s.addText('Yrkesappen', {
    x: 1.0, y: 1.45, w: 11.2, h: 1.0,
    fontFace: 'Calibri', fontSize: 52, bold: true, color: C.white, align: 'center',
  });
  s.addText('Fra planlegging til ferdig undervisning på under 1 minutt', {
    x: 1.0, y: 2.55, w: 11.2, h: 0.7,
    fontFace: 'Calibri', fontSize: 22, color: C.accent, align: 'center',
  });
  s.addText('Pedagogisk kvalitet + teknisk stabilitet', {
    x: 1.0, y: 3.35, w: 11.2, h: 0.5,
    fontFace: 'Calibri', fontSize: 17, color: C.white, align: 'center',
  });
  addNotes(s, [
    'Åpne med verdien: mindre tid på produksjon, mer tid til undervisning.',
    'Fortell at løsningen er utviklet med lærere og for lærere.',
    'Poengter at dette ikke erstatter lærerrollen, men forsterker den.',
  ]);
}

// 2: Problem
{
  const s = pptx.addSlide();
  header(s, 'Utfordringen i dag', 'Hva skoler og lærere bruker tid på');
  bullets(s, [
    'Lærere lager ofte materiell manuelt i flere verktøy.',
    'Det tar tid a lage tekster, oppgaver, ordlister og presentasjoner.',
    'Nivåtilpasning (A1-B2) og morsmålsstøtte blir ofte ujevn.',
    'Resultat: Mindre tid til undervisning, mer tid til produksjon.',
  ]);
  s.addShape(pptx.ShapeType.roundRect, {
    x: 9.5, y: 1.8, w: 3.2, h: 2.1,
    fill: { color: 'FFE7E9' }, line: { color: 'F4B7BD' },
  });
  s.addText('Tidstap\n+ varierende\nkvalitet', {
    x: 9.7, y: 2.1, w: 2.8, h: 1.6, align: 'center',
    fontFace: 'Calibri', fontSize: 20, bold: true, color: '9A1B2A',
  });
  addNotes(s, [
    'Bruk et konkret eksempel fra hverdagen: én undervisningsøkt kan ta flere timer å forberede.',
    'Vis at problemet ikke er lærernes faglighet, men tidspress og fragmenterte verktøy.',
  ]);
}

// 3: Løsning
{
  const s = pptx.addSlide();
  header(s, 'Løsningen: Yrkesappen', 'Ett klikk -> komplette undervisningsressurser');
  bullets(s, [
    'Genererer Word-hefte, PowerPoint og interaktiv HTML fra samme faglige kjerne.',
    'Stotter CEFR-niva A1, A2, B1, B2 med tydelig progresjon.',
    'Språkstøtte med ordlisteforklaring på norsk + valgt hjelpespråk.',
    'Ferdig materiell til klassebruk, storskjerm og individuell trening.',
  ]);
  s.addShape(pptx.ShapeType.roundRect, { x: 9.3, y: 1.8, w: 3.5, h: 2.5, fill: { color: 'E6F7F8' }, line: { color: C.secondary } });
  s.addText('3 formater\n1 arbeidsflyt\n0 dobbeltarbeid', {
    x: 9.5, y: 2.1, w: 3.1, h: 2.0, align: 'center',
    fontFace: 'Calibri', fontSize: 20, bold: true, color: C.primary,
  });
  addNotes(s, [
    'Forklar at alt henger sammen: samme innhold brukes i alle formater.',
    'Dette gir rød tråd i undervisningen og mindre forvirring for elevene.',
  ]);
}

// 4: Hvordan det virker pedagogisk
{
  const s = pptx.addSlide();
  header(s, 'Pedagogisk design i praksis', 'Bygget for laringsutbytte og klasseflyt');
  bullets(s, [
    'Tekst + oppgaver henger sammen, sa eleven trener pa relevant innhold.',
    'Variasjon i oppgavetyper gir bedre mestring og motivasjon.',
    'Direkte riktig/feil-tilbakemelding i HTML styrker egenvurdering.',
    'Klarsprak, tydelig progresjon og struktur for nybegynnere til viderekomne.',
  ]);
  s.addShape(pptx.ShapeType.roundRect, { x: 9.2, y: 1.9, w: 3.7, h: 2.6, fill: { color: 'FFF7E2' }, line: { color: C.accent } });
  s.addText('Pedagogisk\nkonsekvent\npa tvers av\nalle formater', {
    x: 9.45, y: 2.15, w: 3.2, h: 2.2, align: 'center',
    fontFace: 'Calibri', fontSize: 18, bold: true, color: '7C5A00',
  });
  addNotes(s, [
    'Beskriv hvordan variasjon i oppgavetyper øker aktivitet og mestring.',
    'Forklar at direkte tilbakemelding i HTML gjør vurdering og oppfølging enklere.',
  ]);
}

// 5: Teknisk forklart enkelt
{
  const s = pptx.addSlide();
  header(s, 'Teknologi forklart enkelt', 'Robust, sikker og enkel å drifte');
  bullets(s, [
    'Systemet kvalitetssikrer alltid data før noe genereres, så resultatene blir stabile.',
    'Rate limiting beskytter API-et mot overbelastning.',
    'Helse-sjekk gjør drift og overvåking enklere i produksjon.',
    'Modulær arkitektur gir raskere feilretting og tryggere videreutvikling.',
  ]);
  s.addShape(pptx.ShapeType.roundRect, { x: 9.2, y: 1.9, w: 3.7, h: 2.6, fill: { color: 'EAF0FF' }, line: { color: 'B8C8FF' } });
  s.addText('Teknisk kvalitet\nuten teknisk\nkompleksitet\nfor brukeren', {
    x: 9.45, y: 2.2, w: 3.2, h: 2.1, align: 'center',
    fontFace: 'Calibri', fontSize: 18, bold: true, color: '2D3E84',
  });
  addNotes(s, [
    'Unngå teknisk sjargong. Selg trygghet: "Dette fungerer stabilt i en travel skolehverdag."',
    'Poengter at sikkerhet og stabilitet er bygget inn, ikke lagt til i etterkant.',
  ]);
}

// 6: Kundefordel / ROI
{
  const s = pptx.addSlide();
  header(s, 'Forretningsverdi for kunden', 'Hvorfor dette er et godt kjop');
  bullets(s, [
    'Mindre tidsbruk pa materiellproduksjon = mer tid til elevoppfolging.',
    'Hoyere kvalitet og jevnere standard i undervisningen.',
    'Skalerbar losning for enkeltlarere og hele skoler/kommuner.',
    'Rask implementering: krever minimal opplaering for oppstart.',
  ]);
  s.addShape(pptx.ShapeType.roundRect, { x: 9.2, y: 1.9, w: 3.7, h: 2.8, fill: { color: 'E8F8EE' }, line: { color: 'B7E6C7' } });
  s.addText('Mer tid\nBedre kvalitet\nSterkere\nlaringsresultat', {
    x: 9.45, y: 2.2, w: 3.2, h: 2.3, align: 'center',
    fontFace: 'Calibri', fontSize: 19, bold: true, color: '196C3A',
  });
  addNotes(s, [
    'Gi et enkelt regnestykke: spart tid per uke per lærer.',
    'Forklar at gevinsten er både økonomisk og pedagogisk.',
  ]);
}

// 7: Avslutning
{
  const s = pptx.addSlide();
  s.background = { color: C.primary };
  s.addText('Klar for neste steg?', {
    x: 0.8, y: 1.2, w: 11.8, h: 0.9,
    fontFace: 'Calibri', fontSize: 42, bold: true, color: C.white, align: 'center',
  });
  s.addText('Demo | Pilot | Innføring', {
    x: 0.8, y: 2.25, w: 11.8, h: 0.6,
    fontFace: 'Calibri', fontSize: 24, color: C.accent, align: 'center',
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 2.8, y: 3.3, w: 7.7, h: 1.5, fill: { color: 'FFFFFF', transparency: 88 }, line: { color: C.white } });
  s.addText('Yrkesappen gir pedagogisk kraft i hverdagen – med stabil teknologi i bunn.', {
    x: 3.1, y: 3.7, w: 7.1, h: 0.8,
    fontFace: 'Calibri', fontSize: 18, bold: true, color: C.white, align: 'center', fit: 'shrink',
  });
  addNotes(s, [
    'Avslutt med en tydelig CTA: "La oss kjøre en pilot med deres fagområde."',
    'Foreslå konkret neste møte med mål, tidslinje og suksesskriterier.',
  ]);
}

const outputName = 'Yrkesappen-kundepresentasjon.pptx';
pptx.writeFile({ fileName: outputName })
  .then(() => console.log(`Generert: ${outputName}`))
  .catch((err) => {
    console.error('Kunne ikke generere presentasjonen:', err);
    process.exit(1);
  });
