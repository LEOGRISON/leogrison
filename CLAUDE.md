# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

Not a software project. It holds **legal drafting for the Grison family holding** — a Brazilian
succession/asset-restructuring case (Grison e Cia Ltda., Palmas/TO). The only "code" is a document
generator. Everything here is Brazilian Portuguese and stays that way.

The case background lives **outside the repo**, in a PDF the user uploads
(`Descritivo_do_Caso___Grupo_Grison.pdf`). Read it before making substantive changes to any clause —
it carries the debt map, the family structure, the usufruct design, and the ITCMD analysis that the
drafting depends on. Nothing in the repo restates it.

## The one-document, three-formats rule

`docs/` currently holds a single deliverable — the profit-distribution clause of the shareholders'
agreement — in three representations that **must be kept in sync on every substantive edit**:

| File | Role |
|---|---|
| `anexo-i-destinacao-resultados.md` | Source text of the annex |
| `minuta-distribuicao-metas.html` | Published Artifact — same text plus the metas cards and the formula panel |
| `gerar-anexo-i-docx.js` → `Anexo-I-Destinacao-de-Resultados.docx` | Signable Word version, contract typography |

The `.docx` is **generated, never hand-edited** — edit the script and re-run it. A change to a rule, a
value, or a definition has to be applied three times.

The document is written as **Anexo I of the Acordo de Sócios** — pure contract language, numbered
items, no rationale beyond a short preliminary note. The user asked explicitly for it to stop reading
like a memo. Within the Acordo, the Mapa de Dívidas is Anexo II and the pró-labore Anexo III.

The HTML is published as an Artifact at
`https://claude.ai/code/artifact/98f48a8b-67af-4452-a1de-b9e6934790a3`. Republish the same file path
to update that URL rather than creating a new one, and bump the version stamp in all three files
(`versão 3.0` in the docx footer, in the HTML masthead, and `**Versão:**` in the markdown).

## Regenerating and verifying the .docx

`docx` (npm) is **not** preinstalled in this environment, and the script writes to an absolute path:

```bash
cd "$SCRATCHPAD"                                  # anywhere with node_modules
npm install docx
node /home/user/leogrison/docs/gerar-anexo-i-docx.js   # writes docs/Anexo-I-Destinacao-de-Resultados.docx
```

**Always verify the page count and the rendering visually** — LibreOffice reflows differently from the
source, and the annex is meant to stay near three pages:

```bash
apt-get update -qq && apt-get install -y libreoffice-writer poppler-utils   # neither is preinstalled
soffice --headless -env:UserInstallation=file:///tmp/lo --convert-to pdf --outdir . m.docx
pdfinfo m.pdf | grep Pages
pdftoppm -jpeg -r 100 m.pdf pg      # then Read the page images
```

Base LibreOffice ships without the Writer filter here — without `libreoffice-writer`, *every* `.docx`
fails with "source file could not be loaded", which looks like a corrupt document but is not.

When the layout overflows, trim prose first (the user wants substance kept); the disclaimer already
lives in a real page footer so it never consumes body flow.

## Substantive parameters of the current draft

These recur across all three files and are the things most likely to change:

- **Meta I** — dívida zero of the *Perímetro de Endividamento do Grupo*, defined by **economic
  origin** (business led by Wilson Grison) rather than by CNPJ, so it reaches debts registered under
  children, spouse, or third parties. Proven objectively by certidões + no protest + no Serasa/SPC
  entry + discharge of gravames. Parcelamento in good standing is explicitly **not** quitação.
- **Meta II** — R$ 2.000.000,00 caixa mínimo.
- **Meta III** — the **gross** amount of the apartment operation, R$ 2.144.197,27 for a furnished
  apartment of R$ 1.600.000,00. The gross exists because only Elman's biological children (Leonardo
  and Bárbara, 37,31% each) may bear the apartment's cost, while Wellington and Marcelo (12,44% each)
  and Leonardo's original 0,5% quota are paid out in cash on the same base. Hence
  `bruto = valor do imóvel ÷ 74,62%`. The contract fixes 74,62%, not the exact 74,625%, so the
  document matches the family's spreadsheet — do not "correct" it. These percentages govern this
  operation only, not other distributions.
- **Excluded from the perimeter**: the Santa Helena advance (self-liquidating — the loteador withholds
  repayments), and each partner's strictly personal debts.
- **Who the lock binds**: children as nu-proprietários, cotistas, and future partners (including
  employees entering partnership). The **usufructuaries Wilson and Elman are entirely outside it** —
  the fruits are theirs by right and are paid in full, without cap or condition.
- No fixed term: the lock is released only by the metas, never by elapsed time; a 60-month review
  does not relax them.

Open for the family's lawyer (Dr. Roger Lippi) and the tax specialist: the effect of a 120-month PGFN
transaction on Meta I (parcelamento is not quitação, so it postpones the meta); the statutory
mechanism for retaining some partners' dividends while paying the usufructuaries; and the tax
treatment of the company buying the apartment and registering it in Elman's name.

## Working conventions

- Branch: `claude/family-property-debt-structure-gvet5m`. Commit messages in Portuguese, matching the
  existing log.
- Drafting voice: clause text is formal legal Portuguese with bracketed placeholders (`[•]`, `[90]`,
  `[IPCA]`) for what the lawyer fills in. Cover notes and drafting notes are plain-spoken.
- Every document carries the disclaimer that it is a working draft, is not a legal opinion, and needs
  review by the family lawyer and a tax specialist before signature. Keep it.
