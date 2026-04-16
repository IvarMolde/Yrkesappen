# Yrkesappen 🏢

**Molde voksenopplæringssenter – MBO**

Genererer skreddersydde Word-arbeidshefte og PowerPoint-presentasjoner for yrkesrettet norskopplæring (CEFR A1–B2). 

---

## Hva appen lager

For hvert yrke og norsknivå genereres en ZIP med:

| Fil | Innhold |
|-----|---------|
| `[yrke]-arbeidshefte-[nivå].docx` | 3 yrkestekster · ordliste · 10–13 oppgaver (a–e) |
| `[yrke]-presentasjon-[nivå].pptx` | 7 slides: tittel · om yrket · ord · HMS · egenskaper · utdanning · avslutning |

---

## Teknologi

| Komponent | Teknologi |
|-----------|-----------|
| Backend | Node.js + Express |
| AI-motor | Google Gemini 1.5 Flash (balansert kraft/kostnad) |
| Word | `docx` npm-pakke |
| PowerPoint | `pptxgenjs` npm-pakke |
| Nedlasting | `archiver` (ZIP) |
| Deploy | Vercel |

---

## Lokal utvikling

```bash
# Klon/kopier prosjektet
cd yrkesappen

# Installer avhengigheter
npm install

# Sett miljøvariabel
export GEMINI_API_KEY=din-nøkkel-her

# Start server
npm start
# → Åpne http://localhost:3000
```

---

## Deploy til Vercel

### 1. Push til GitHub

```bash
git init
git add .
git commit -m "Yrkesappen v1"
git remote add origin https://github.com/DITT-BRUKERNAVN/yrkesappen.git
git push -u origin main
```

### 2. Importer i Vercel

1. Gå til [vercel.com](https://vercel.com) → **Add New Project**
2. Velg GitHub-repoet `yrkesappen`
3. Under **Environment Variables**, legg til:
   - **Name:** `GEMINI_API_KEY`
   - **Value:** din Google Cloud Gemini API-nøkkel
4. Klikk **Deploy**

### 3. Ferdig!

Appen er live på `https://yrkesappen.vercel.app` (eller ditt valgte domene).

---

## Miljøvariabler

| Variabel | Påkrevd | Beskrivelse |
|----------|---------|-------------|
| `GEMINI_API_KEY` | ✅ | Google Cloud Gemini API-nøkkel |
| `PORT` | ❌ | Port for lokal kjøring (standard: 3000) |

---

## Designguide

Se [`DESIGN.md`](./DESIGN.md) for fullstendig dokumentasjon av fargepalett, typografi og layoutregler for Word og PowerPoint.

**Fargepalett:**
- Primær: `#005F73` (teal)
- Sekundær: `#0A9396`
- Aksent: `#E9C46A` (gull)
- Alle farger er WCAG 2.1 AA-godkjente

---

## Filstruktur

```
yrkesappen/
├── server.js          # Backend: API, DOCX- og PPTX-bygging
├── DESIGN.md          # Fullstendig designguide for dokumenter
├── package.json
├── vercel.json
├── README.md
└── public/
    └── index.html     # Frontend
```

---

*Yrkesappen v1.0 – Molde voksenopplæringssenter*
