# jobvault-generation-service

A lightweight Node.js/TypeScript HTTP service that generates tailored `.docx` CV and cover letter documents for the JobVault pipeline. Called internally by the JobVault Worker Service (.NET).

---

## How It Works

The Worker Service sends a JSON payload describing the tailored content for a specific job application. The generation service merges that content into pre-built `.docx` templates (using [docxtemplater](https://docxtemplater.com/)) and returns the rendered binary file directly in the HTTP response.

**What is fixed in every document (hardcoded in the template):**
- Name, contact details, LinkedIn, GitHub
- All 4 role headers: title, company, location, dates
- Education, Projects, Certifications, Languages

**What varies per job application (comes from the payload):**
- Headline (subtitle line under the name)
- Professional summary
- Skills table (rows and order adjusted per JD)
- Bullet points under each role
- Cover letter paragraphs

---

## Endpoints

### `GET /health`
Returns `{"status":"ok"}`. Used to verify the service is running.

---

### `POST /api/generate-cv`
Generates a tailored CV and returns it as a `.docx` binary.

### `POST /api/generate-cover-letter`
Generates a tailored cover letter and returns it as a `.docx` binary.

Both endpoints accept the same JSON payload shape and return:
- `Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- Body: raw `.docx` binary

---

## Payload Contract

```json
{
  "company": "string",
  "role": "string",
  "jdSource": "string",
  "headline": "string",
  "summary": "string",
  "skills": [
    { "label": "string", "value": "string" }
  ],
  "roles": [
    {
      "id": "calvergy | senior_baris | developer_baris | junior_baris",
      "bullets": ["string"]
    }
  ],
  "recipient": "string",
  "coverLetterParagraphs": ["string"],
  "compatibilityScore": 0,
  "strengths": ["string"],
  "gaps": ["string"],
  "tailoringNotes": "string"
}
```

### Field notes

| Field | Used by | Notes |
|---|---|---|
| `headline` | Both | Subtitle under the name — e.g. `"Full Stack Developer  \|  React.js  \|  TypeScript"` |
| `summary` | CV only | Professional summary paragraph |
| `skills` | CV only | Ordered array — put most JD-relevant categories first |
| `roles[].id` | CV only | Must be one of the four fixed IDs below |
| `roles[].bullets` | CV only | Supports `**bold**` markdown — rendered as Word bold |
| `recipient` | CL only | Salutation line — e.g. `"Dear Hiring Team,"` |
| `coverLetterParagraphs` | CL only | One string per paragraph |

### Valid role IDs

| ID | Role |
|---|---|
| `calvergy` | Software Engineer — Calvergy UG |
| `senior_baris` | Senior Software Developer — Bari's Technology Solutions |
| `developer_baris` | Software Developer — Bari's Technology Solutions |
| `junior_baris` | Junior Software Developer — Bari's Technology Solutions |

---

## Running Locally

```bash
npm install
npm run build
npm start
```

Or in dev mode (no build step):
```bash
npm run dev
```

Service listens on port `3000` by default. Override with the `PORT` environment variable.

---

## Kept Running with PM2

The service is registered with PM2 to survive reboots:

```bash
# First-time setup (already done)
npm run build
pm2 start dist/index.js --name jobvault-generation-service
pm2 save

# After a code change
npm run build
pm2 restart jobvault-generation-service
pm2 save

# Check status
pm2 status
pm2 logs jobvault-generation-service
```

---

## Updating the .docx Templates

The visual design (fonts, colours, margins, all static sections) lives in the builder scripts, not in the generation code. If you need to change the layout, edit the relevant script and rebuild:

```bash
# Rebuild both templates
npm run build:templates

# Rebuild one at a time
node scripts/build-cv-template.js
node scripts/build-cl-template.js
```

Templates are committed to `templates/` as build artifacts.

### What is dynamic in each template

**CV template (`cv_template.docx`)**
- `{headline}`, `{summary}`
- `{#skills}{label}{value}{/skills}` — table row loop
- `{#calvergy_bullets}{text}{/calvergy_bullets}` — and equivalent for each of the other 3 roles

**Cover letter template (`cl_template.docx`)**
- `{headline}`, `{date}`, `{company}`, `{role}`, `{recipient}`
- `{#paragraphs}{text}{/paragraphs}` — body paragraph loop

---

## Project Structure

```
src/
  index.ts                    # Entry point — starts Express
  types/
    GenerationPayload.ts      # Shared DTO interfaces
  core/
    generateCv.ts             # CV generation logic
    generateCoverLetter.ts    # Cover letter generation logic
    boldMarkdown.ts           # Post-processes **bold** markers into Word runs
  rest/
    router.ts                 # Route handlers + payload validation
    server.ts                 # Express app factory
scripts/
  build-cv-template.js        # Generates templates/cv_template.docx
  build-cl-template.js        # Generates templates/cl_template.docx
templates/
  cv_template.docx            # Compiled CV template (committed)
  cl_template.docx            # Compiled CL template (committed)
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `express` | HTTP server |
| `docxtemplater` | Tag substitution and loops in `.docx` templates |
| `pizzip` | ZIP manipulation (required by docxtemplater) |
| `docx` | Used by builder scripts to generate the template files |
| `adm-zip` | Used by builder scripts to patch table border XML |
