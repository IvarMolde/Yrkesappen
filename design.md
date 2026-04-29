# Yrkesappen – Designguide for Word og PowerPoint

Denne filen beskriver designprofilen for alle dokumenter som genereres av Yrkesappen.
Serveren leser denne filen og bruker den som referanse når den bygger DOCX og PPTX.

---

## Overordnede prinsipper

- **Tilgjengelighet (WCAG 2.1 AA)**: Minimum kontrastforhold 4.5:1 for brødtekst, 3:1 for store overskrifter
- **Ryddighet**: Mye luft, tydelig hierarki, aldri overfylt
- **Profesjonelt og imøtekommende**: Passer for voksne innvandrere i yrkesrettet norskopplæring
- **Konsistent**: Samme fargepalett og typografi i begge dokumenttyper
- **Pedagogisk tydelighet**: Oppgavenumre og instruksjoner skal skille seg visuelt fra oppgavetekst

---

## Fargepalett

| Rolle              | Farge       | Hex       | Bruk                                      |
|--------------------|-------------|-----------|-------------------------------------------|
| Primærfarge        | Teal/blågrønn | `#005F73` | Overskrifter, seksjonstopper, slide-bakgrunner |
| Sekundærfarge      | Lys teal    | `#0A9396` | Underoverskrifter, aksenter               |
| Aksentfarge        | Gull/amber  | `#E9C46A` | Oppgavenummer-bokser, fremhevinger        |
| Bakgrunn lys       | Kremhvit    | `#F8F9FA` | Sidebakgrunn i Word, lyse slides          |
| Bakgrunn grå       | Lys grå     | `#E9ECEF` | Alternerende tabellrader, oppgavebokser   |
| Tekstfarge mørk    | Nesten svart | `#1B1B1B` | Brødtekst                                |
| Tekstfarge medium  | Mørkgrå     | `#495057` | Bildetekst, instruksjoner                 |
| Hvit               | Hvit        | `#FFFFFF` | Tekst på mørke bakgrunner                 |

**Kontrastsjekk (WCAG AA):**
- `#FFFFFF` på `#005F73` → ratio 7.2:1 ✅
- `#1B1B1B` på `#F8F9FA` → ratio 18.1:1 ✅
- `#1B1B1B` på `#E9ECEF` → ratio 14.2:1 ✅
- `#1B1B1B` på `#E9C46A` → ratio 5.8:1 ✅

---

## Typografi

| Element              | Font          | Størrelse | Stil        |
|----------------------|---------------|-----------|-------------|
| Dokumenttittel       | Calibri       | 28 pt     | Bold        |
| H1 (seksjoner)       | Calibri       | 20 pt     | Bold        |
| H2 (underseksjoner)  | Calibri       | 16 pt     | Bold        |
| Brødtekst            | Calibri       | 12 pt     | Regular     |
| Oppgavenummer        | Calibri       | 13 pt     | Bold        |
| Instruksjon          | Calibri       | 12 pt     | Italic      |
| Ordlistetabell       | Calibri       | 11 pt     | Regular     |
| Fotnote/bildetekst   | Calibri       | 10 pt     | Regular     |

**Linjeavstand**: 1.15 for brødtekst, 1.0 for tabeller
**Avsnittavstand**: 6 pt etter hvert avsnitt, 12 pt etter overskrifter

---

## Word-dokument (DOCX)

### Sideformat
- Papirstørrelse: A4 (11906 × 16838 DXA)
- Marger: 2 cm alle kanter (1134 DXA)
- Topp-/bunntekst: Ja

### Topptekst
- Venstre: Yrke + nivå (f.eks. «Sykepleier – A2»)
- Høyre: «Molde voksenopplæringssenter»
- Skrift: Calibri 9 pt, farge `#495057`
- Skille: Tynn linje under (BorderStyle.SINGLE, farge `#005F73`, tykkelse 6)

### Bunntekst
- Venstre: «© MBO – Molde voksenopplæringssenter»
- Høyre: Sidetall (f.eks. «Side 1»)
- Skrift: Calibri 9 pt, farge `#495057`

### Forsideblokk (øverst i dokumentet, ingen egen side)
- Blå boks (farge `#005F73`): Tittel i hvit Calibri 28 pt bold
- Under tittelen: Yrke, nivå og dato i Calibri 12 pt, farge `#0A9396`

### Seksjoner i dokumentet (rekkefølge)
1. **Introduksjon** – kort tekst om yrket
2. **Yrkestekster** (3 stk.) – hver med egen H2-overskrift, tydelig adskilt
3. **Viktige ord og uttrykk** – tabell med 2 eller 3 kolonner (norsk | forklaring | evt. hjelpespråk)
4. **Oppgaver** – nummerert 1–N, delopgaver a–e

### Oppgaveboks-design
- Oppgavenummer («Oppgave 1») i liten boks med bakgrunn `#005F73`, hvit tekst
- Tittel på oppgaven i Calibri 13 pt bold under boksen
- Instruksjon i kursiv, farge `#495057`
- Delopgaver (a–e) med innrykk, Calibri 12 pt
- Svarlinjer: 3 tomme linjer med understrekning (`___________`) etter hvert delspørsmål der det passer
- Alternerende bakgrunn på delopgavene: a/c/e = hvit, b/d = `#E9ECEF`

### Ordlistetabell
- Kolonner: «Norsk» | «Forklaring» | (evt.) «[Hjelpespråk]»
- Toppraden: Bakgrunn `#005F73`, hvit bold tekst
- Alternerende rader: hvit / `#E9ECEF`
- Alle celleborder: `#CCCCCC`, 1 pt

### Fargebruk i tekster
- Nye ord som er i ordlisten: Uthev med farge `#0A9396` (teal) og kursiv første gang de brukes
- Viktige setninger: Kan ha lysgrå boks (`#E9ECEF`) rundt seg

---

## PowerPoint (PPTX)

### Oppsett
- Layout: `LAYOUT_16x9` (10" × 5.625")
- Font-par: Tittel = **Calibri Bold**, brødtekst = **Calibri**

### Slide-struktur (fast rekkefølge)

#### Slide 1 – Tittelslide (mørk)
- Bakgrunn: `#005F73` (hel farge)
- Stor yrkestittel: hvit, 44 pt bold, sentrert
- Undertittel: «Norsknivå [X] – Molde voksenopplæringssenter», hvit 20 pt
- Dekorativt gull-rektangel (`#E9C46A`) nederst som "banner" (0.15" høy, full bredde)

#### Slide 2 – Om yrket (lys)
- Bakgrunn: `#F8F9FA`
- Venstre kolonne (60%): 4–6 stikkord om hva yrket innebærer
- Høyre kolonne (40%): Ikon/symbolboks i `#005F73`
- Tittel: «Om yrket», teal bold

#### Slide 3 – Viktige ord og uttrykk (lys)
- Bakgrunn: `#F8F9FA`
- 8 nøkkelord i to kolonner, hvert ord i en liten boks
- Boks: hvit bakgrunn, `#005F73` venstrekant (0.06"), lett skygge
- Tittel: «Viktige ord», teal bold

#### Slide 4 – HMS (lys med gull-aksent)
- Bakgrunn: `#F8F9FA`
- 4 HMS-punkter som ikonliste (med fargerik prikk foran hvert punkt)
- Stor «HMS»-tekst til høyre som bakgrunnsdekor (60 pt, `#E9ECEF`, transparent)
- Tittel: «Helse, miljø og sikkerhet», teal bold

#### Slide 5 – Personlige egenskaper (lys)
- Bakgrunn: `#F8F9FA`
- 5 egenskaper i buede kort (RECTANGLE med avrunding) i to-kolonne grid
- Kortbakgrunn: `#E9ECEF`, venstrekant `#0A9396`
- Tittel: «Personlige egenskaper», teal bold

#### Slide 6 – Arbeidsoppgaver (lys)
- Bakgrunn: `#F8F9FA`
- 4 arbeidsoppgaver som nummerert liste med store tallikoner (sirkler i `#005F73`)
- Tittel: «Arbeidsoppgaver», teal bold

#### Slide 7 – Utdanning og veien videre (mørk)
- Bakgrunn: `#005F73`
- 3 utdanningsveier/karrieremuligheter som hvite kort på mørk bakgrunn
- Tittel: «Utdanning og karriere», hvit bold
- Dekorativt gull-rektangel øverst (`#E9C46A`, 0.15" høy)

#### Slide 8 – Avslutning (mørk)
- Bakgrunn: `#005F73`
- Stor tekst: «Lykke til!», hvit, sentrert
- Undertekst: «Molde voksenopplæringssenter – MBO», hvit 16 pt
- Gull-banner (`#E9C46A`) øverst og nederst

### Generelle PowerPoint-regler
- ALDRI unicode-bullets – bruk `bullet: true`
- ALDRI `#` foran hex-farger
- ALDRI dele option-objekter mellom shapes (PptxGenJS muterer dem)
- Minimum 0.5" marger fra slidekantene
- Titler: alltid venstrejustert unntatt tittelslide (sentrert)
- Brødtekst: alltid venstrejustert
- Tekstbokser over fargebakgrunner: sett `margin: 0` for presis justering

---

## Tilgjengelighet (WCAG)

- Alle bilder/ikoner skal ha `altText`
- Kontrast for all synlig tekst: minimum 4.5:1 (se fargepaletten over)
- Ikke bare farge for å formidle informasjon – bruk også form/tekst
- Skriftstørrelse aldri under 10 pt i trykte dokumenter, 14 pt i presentasjoner
- Unngå rene dekorative elementer som forstyrrer lesbarhet

---

## Hjelpespråk

Dersom hjelpespråk er valgt:

**Tosidig (side om side):**
- I Word-tekster: oversettelse i parentes etter norsk, teal kursiv – f.eks. *lege (doctor)*
- I ordlisten: egen kolonne til høyre

**Ordliste til slutt:**
- Eget avsnitt «Ordliste – [språk]» bakerst i dokumentet
- Tabell: Norsk | [Språk] (2 kolonner)
- I PPTX: egen slide etter slide 3 med oversettelsestabellen

---

*Designguide versjon 1.0 – Yrkesappen, Molde voksenopplæringssenter*
