# Yrkesappen – Designguide for Word og PowerPoint

Denne filen beskriver designprofilen og pedagogisk struktur for alle dokumenter som genereres av Yrkesappen. `server.js` bruker denne filen som referanse når den bygger DOCX og PPTX.

---

## Overordnede prinsipper

- **Tilgjengelighet (WCAG 2.1 AA)**: Minimum kontrastforhold 4.5:1 for brødtekst, 3:1 for store overskrifter
- **Ryddighet**: Mye luft, tydelig hierarki, aldri overfylt
- **Profesjonelt og imøtekommende**: Passer for voksne innvandrere i yrkesrettet norskopplæring
- **Konsistent**: Samme fargepalett og typografi i begge dokumenttyper – eleven kjenner igjen stilen
- **Pedagogisk tydelighet**: Oppgavenumre, tekstnummer og instruksjoner skal skille seg visuelt fra innholdstekst
- **Visuell sammenheng**: Word-heftet og PowerPoint skal oppleves som ett sammenhengende læringsprodukt

---

## Fargepalett

| Rolle             | Farge         | Hex       | Bruk                                                |
|-------------------|---------------|-----------|-----------------------------------------------------|
| Primærfarge       | Teal          | `#005F73` | Overskrifter, mørke slide-bakgrunner, venstrekanter |
| Sekundærfarge     | Lys teal      | `#0A9396` | Underoverskrifter, tekst-tags, ikoner               |
| Aksentfarge       | Gull/amber    | `#E9C46A` | Oppgavebokser, badges, fokus-merking, bannere       |
| Bakgrunn lys      | Kremhvit      | `#F8F9FA` | Sidebakgrunn i Word, lyse slides                    |
| Bakgrunn grå      | Lys grå       | `#E9ECEF` | Alternerende tabellrader, egenskapskort             |
| Tekstfarge mørk   | Nesten svart  | `#1B1B1B` | Brødtekst                                           |
| Tekstfarge medium | Mørkgrå       | `#495057` | Instruksjoner, bildetekst, subtitler                |
| Hvit              | Hvit          | `#FFFFFF` | Tekst på mørke bakgrunner                           |

**Kontrastsjekk (WCAG AA):**
- `#FFFFFF` på `#005F73` → ratio 7.2:1 ✅
- `#1B1B1B` på `#F8F9FA` → ratio 18.1:1 ✅
- `#1B1B1B` på `#E9ECEF` → ratio 14.2:1 ✅
- `#1B1B1B` på `#E9C46A` → ratio 5.8:1 ✅

---

## Typografi

| Element               | Font    | Størrelse  | Stil    |
|-----------------------|---------|------------|---------|
| Dokumenttittel (DOCX) | Calibri | 28 pt      | Bold    |
| H1 (seksjoner)        | Calibri | 20 pt      | Bold    |
| H2 (underseksjoner)   | Calibri | 16 pt      | Bold    |
| Brødtekst             | Calibri | 12 pt      | Regular |
| Oppgavenummer         | Calibri | 13 pt      | Bold    |
| Instruksjon           | Calibri | 12 pt      | Italic  |
| Ordlistetabell        | Calibri | 11 pt      | Regular |
| Slide-tittel (PPTX)   | Calibri | 24 pt      | Bold    |
| Slide-brødtekst       | Calibri | 13–17 pt   | Regular |
| Slide-undertittel     | Calibri | 12 pt      | Italic  |

---

## Word-dokument (DOCX)

### Formål
Arbeidsheftet er **elevens** primære læringsressurs. Det brukes individuelt eller i par etter at læreren har hatt gjennomgang med PowerPoint-presentasjonen.

### Sideformat
- Papirstørrelse: A4 (11906 × 16838 DXA)
- Marger: 2 cm alle kanter (1134 DXA)
- Topptekst og bunntekst: Ja

### Topptekst
- Venstre: Yrke + nivå (f.eks. «Sykepleier – A2»)
- Høyre: «Molde voksenopplæringssenter»
- Skrift: Calibri 9 pt, farge `#495057`
- Skille: Tynn linje under (BorderStyle.SINGLE, farge `#005F73`, tykkelse 6)

### Bunntekst
- Venstre: «© MBO – Molde voksenopplæringssenter»
- Høyre: Sidetall
- Skrift: Calibri 9 pt, farge `#495057`

### Dokumentstruktur (rekkefølge)
1. **Innledning** – kort introduksjonstekst om yrket
2. **Fagtekster og oppgaver** – flettet struktur:
   - Tekst 1 → Leseforståelse (Oppgave 1) → Grammatikk (Oppgave 2)
   - Tekst 2 → Leseforståelse (Oppgave 3) → Vokabular (Oppgave 4)
   - Tekst 3 → Leseforståelse (Oppgave 5) → Grammatikk (Oppgave 6)
   - Avsluttende vokabularoppgave (Oppgave 7) + Skriv/muntlig (Oppgave 8)
3. **Viktige ord og uttrykk** – tabell med 2 eller 3 kolonner
4. **Ordliste på hjelpespråk** (valgfritt, bakerst)

### Oppgaveboks-design
- Oppgavenummer: Teal boks (`#005F73`), hvit Calibri 13 pt bold
- Teksttilknytning-tag: Sekundærboks (`#0A9396`) med ikon + «Tekst N»
- Tittel: Calibri 14 pt bold, mørk tekst
- Instruksjon: Kursiv, farge `#495057`
- Delopgaver (a–e): Alternerende hvit/grå bakgrunn, svarlinje under hvert delspørsmål

### Ordlistetabell
- Kolonner: «Norsk» | «Forklaring» | (evt.) «[Hjelpespråk]»
- Toppraden: Bakgrunn `#005F73`, hvit bold tekst
- Alternerende rader: hvit / `#E9ECEF`
- Alle celleborder: `#CCCCCC`, 1 pt

---

## PowerPoint (PPTX)

### Formål
Presentasjonen er **lærerens** verktøy for klasseromsgjennomgang *før* elevene åpner Word-heftet. Målet er å:
1. Aktivere forkunnskaper om yrket
2. Forberede elevene på fagord og tekster
3. Skape motivasjon og nysgjerrighet
4. Gi læreren diskusjonsspørsmål og samtalestartere
5. Gjennomgå HMS og personlige egenskaper

### Teknisk oppsett
- Layout: `LAYOUT_16x9` (10" × 5.625")
- Font: Calibri overalt
- Alle tekstbokser: `wrap: true` + `shrinkText: true` (forhindrer avkuttet tekst)
- Alle innholdstekstbokser: `margin: 6–8` (aldri 0 på innhold)
- Minimum boksbredde: 2.0" | Minimum bokshøyde: 0.45"

### Slide-maler

**Mørk slide (darkSlide):**
- Bakgrunn: `#005F73`
- Topp-banner: `#E9C46A`, 0.20" høy, full bredde
- Bunn-banner: `#E9C46A`, 0.20" høy, full bredde
- Brukes til: Tittelslide, Utdanning

**Lys slide (lightSlide):**
- Bakgrunn: `#F8F9FA`
- Venstre fargebolk: `#005F73`, 0.15" bred, full høyde
- Tittelområde: hvit bakgrunn, 0.95" høy
- Tittel: Calibri 24 pt bold, `#005F73`, venstrejustert
- Undertittel: Calibri 12 pt italic, `#495057`
- Skillelinje: `#0A9396`, 2 pt
- Brukes til: Alle faglige innholdsslides

**Fokus-badge** (kun hvis faglig fokus er angitt):
- Plassering: Øverst til høyre på lyse slides
- Størrelse: 2.0" × 0.38", bakgrunn `#E9C46A`
- Tekst: «Fokus: [brukerens fokus]», Calibri 9 pt bold, `#1B1B1B`

---

### Slide-struktur (10 slides, fast rekkefølge)

#### Slide 1 – Tittel (mørk)
- Yrkestittel: 44 pt bold, hvit, sentrert, wrap+shrink
- Nivå: 22 pt, gull, sentrert
- MBO-tekst: 14 pt, `#E9ECEF`, sentrert
- Fokustekst (valgfritt): 13 pt italic, gull, sentrert

**Pedagogisk formål:** Introdusere temaet og skape forventning.

#### Slide 2 – Hva er dette yrket? (lys)
- Venstre (62%): Punktliste med arbeidsoppgaver, 15 pt, wrap+shrink
- Høyre (28%): Primærfargeboks med stor forbokstav (110 pt hvit) + yrkestittel i gull

**Pedagogisk formål:** Aktivere forkunnskaper. Læreren spør: «Hva vet dere om dette yrket?»

#### Slide 3 – Viktige ord og uttrykk (lys)
Uten hjelpespråk: 2×4 kortgrid (8 ord)
- Hvert kort: 4.6" × 1.0", hvit bakgrunn, teal venstrekant, skygge
- Øverst: fagord teal bold 14 pt
- Nederst: forklaring norsk 12 pt grå

Med hjelpespråk: Tospråklig tabell (8 rader)
- Kolonneheader: primærfarge (norsk) + sekundærfarge (hjelpespråk)
- Alternerende hvit/`#E9ECEF` rader, 0.46" høyde per rad
- Norsk: teal bold 13 pt | Oversettelse: kursiv 12 pt

**Pedagogisk formål:** Forberede elevene på ordforrådet. Knyttes direkte til ordlisten i Word.

#### Slide 4 – Tekst 1: [tittel] (lys)
- Venstre panel (58%): Hvit boks med gull venstrekant, diskusjonsspørsmål 18 pt teal bold
- Høyre panel (35%): Teal boks, «Tekst 1» i gull, 4 nøkkelord fra ordlisten

**Pedagogisk formål:** Forberede elevene mentalt på Tekst 1 før de leser.

#### Slide 5 – Tekst 2: [tittel] (lys)
- Samme layout som Slide 4 med nytt spørsmål og neste 4 ord

**Pedagogisk formål:** Aktivere faglig nysgjerrighet for Tekst 2.

#### Slide 6 – Tekst 3: [tittel] (lys)
- Samme layout med siste spørsmål og siste relevante ord

**Pedagogisk formål:** Forberede elevene på den mest krevende teksten.

#### Slide 7 – HMS (lys)
- Bakgrunnsdekor: «HMS» i `#E9ECEF`, 130 pt, transparent, høyre side
- 4 HMS-punkter med farget sirkel (0.65" diameter) og bred tekstboks (5.3" × 0.9")
- Tekstboks alltid generøs høyde (0.9") for å unngå avkuttet tekst

**Pedagogisk formål:** Gjennomgå HMS-regler før elevene møter det i Tekst 3.

#### Slide 8 – Personlige egenskaper (lys)
- 5 kort i 3+2-layout: 3.0" × 1.9" (generøs høyde)
- `#E9ECEF` bakgrunn, teal venstrekant, nummermerke øverst til høyre
- Tekst: 14 pt bold, wrap+shrink, margin 8

**Pedagogisk formål:** Diskutere hvilke egenskaper som er viktige i yrket.

#### Slide 9 – Utdanning og karriere (mørk)
- 3 kolonner (3.0" brede), halvtransparente hvite bokser (4.0" høye)
- Tall i gull 34 pt, gull skillelinje, hvit tekst 14 pt wrap+shrink
- Tekstboks 2.9" høy sikrer at lang tekst vises

**Pedagogisk formål:** Gi perspektiv og motivasjon — hva kan man bli?

#### Slide 10 – La oss snakke norsk! (lys)
- 3 yrkesspesifikke diskusjonsspørsmål (generert av Gemini)
- Brede bokser: 9.3" × 1.25", hvit bakgrunn, teal kant
- Teal tallfelt (0.5" bred) til venstre, spørsmålstekst 17 pt teal bold wrap+shrink

**Pedagogisk formål:** Warm-up/avslutningsaktivitet. Muntlig aktivitet før elevene starter på Word-heftet.

---

## Regler for tekstvisibilitet (kritisk)

Alle tekstbokser i PPTX følger disse reglene for å sikre at ingen tekst kuttes:

| Regel | Verdi | Begrunnelse |
|-------|-------|-------------|
| `wrap: true` | Alltid | Tekst brytes til neste linje |
| `shrinkText: true` | Alltid unntatt dekor | PowerPoint krymper font automatisk |
| `margin` | Min. 4, helst 6–8 | Luft mellom tekst og boksramme |
| Boksbredde | Min. 2.0" | Unngår for smal wrapping |
| Bokshøyde | Min. 0.45" | Plass til minst én linje |
| Tekststørrelse innhold | 12–18 pt | Lesbart fra 3–4 meters avstand |

**Kritiske forbud i PptxGenJS:**
- ALDRI `margin: 0` på innholdstekst
- ALDRI fast bokshøyde under 0.45" for variabelt innhold
- ALDRI dele `options`-objekter mellom shapes (muteres in-place)
- ALDRI `#` foran hex-farger
- ALDRI unicode-bullets (`•`) — bruk `bullet: true`
- ALDRI 8-tegns hex for opacity — bruk `opacity`-property separat

---

## Hjelpespråk

### I Word
- **Tosidig**: Oversettelse i parentes etter norsk term, kursiv teal
- **Ordliste til slutt**: Eget avsnitt «Ordliste – [språk]» bakerst, 2-kolonne tabell

### I PowerPoint
- **Slide 3**: Bytter fra kortgrid til tospråklig tabell
- **Slides 4–6**: Oversettelse i kursiv gull (11 pt) under nøkkelord i høyre panel
- **Prinsipp**: Hjelpespråket er støttende, aldri dominerende over norsk

---

## Faglig fokus

Hvis læreren har angitt fokus (f.eks. «verktøy», «kommunikasjon», «medisiner»):

### I Word
- Gemini integrerer fokuset i alle tre fagtekster, ordlisten og oppgavene

### I PowerPoint
- **Alle lyse slides**: Gull fokus-badge øverst til høyre
- **Slide 1**: Fokusteksten i kursiv gull
- **Formål**: Læreren ser alltid det pedagogiske fokuset og kan styre gjennomgangen

---

## Sammenheng mellom Word og PowerPoint

| PowerPoint-slide           | Tilsvarer i Word                     |
|---------------------------|--------------------------------------|
| Slide 2 – Hva er yrket?   | Innledning                           |
| Slide 3 – Viktige ord     | Ordlisten                            |
| Slide 4 – Tekst 1         | Tekst 1 + Oppgave 1–2                |
| Slide 5 – Tekst 2         | Tekst 2 + Oppgave 3–4                |
| Slide 6 – Tekst 3         | Tekst 3 + Oppgave 5–6                |
| Slide 7 – HMS             | Tekst 3 (HMS-fokus) + Oppgave 6      |
| Slide 8 – Egenskaper      | Ordlisten + Oppgave 8                |
| Slide 9 – Utdanning       | Bakgrunnsinformasjon                 |
| Slide 10 – Snakk norsk!   | Oppgave 8 (muntlig)                  |

---

*Designguide versjon 2.0 – Yrkesappen, Molde voksenopplæringssenter*
