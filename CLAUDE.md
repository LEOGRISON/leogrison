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
| `minuta-clausula-metas-distribuicao.md` | Long form: cover note, full clause, drafting notes (Parte I / II / III) |
| `minuta-distribuicao-metas.html` | Same content as a published Artifact, for presenting to the father |
| `gerar-minuta-docx.js` → `Minuta-Distribuicao-por-Metas.docx` | Condensed two-page print version |

The `.docx` is **generated, never hand-edited** — edit the script and re-run it. The three files
carry different levels of detail on purpose (the `.docx` drops the drafting notes), so a change to a
rule, a value, or a definition has to be applied three times; a change to prose depth does not.

The HTML is published as an Artifact at
`https://claude.ai/code/artifact/98f48a8b-67af-4452-a1de-b9e6934790a3`. Republish the same file path
to update that URL rather than creating a new one, and bump the version stamp in all three files
(`VERSÃO 2.0` in the docx header, `versão 2.0` in the HTML masthead, `**Versão:**` in the markdown).

## Regenerating and verifying the .docx

`docx` (npm) is **not** preinstalled in this environment, and the script writes to an absolute path:

```bash
cd "$SCRATCHPAD"                                  # anywhere with node_modules
npm install docx
node /home/user/leogrison/docs/gerar-minuta-docx.js   # writes docs/Minuta-Distribuicao-por-Metas.docx
```

**The two-page limit is a hard requirement** the user set. Always verify visually — LibreOffice
reflows differently from the source:

```bash
apt-get update -qq && apt-get install -y libreoffice-writer poppler-utils   # neither is preinstalled
soffice --headless -env:UserInstallation=file:///tmp/lo --convert-to pdf --outdir . m.docx
pdfinfo m.pdf | grep Pages          # must say 2
pdftoppm -jpeg -r 100 m.pdf pg      # then Read pg-1.jpg / pg-2.jpg
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
- **Meta II** — R$ 2.000.000,00 caixa mínimo. **Meta III** — R$ 1.600.000,00 for Elman's apartment.
- **Excluded from the perimeter**: the Santa Helena advance (self-liquidating — the loteador withholds
  repayments), and each partner's strictly personal debts.
- **Who the lock binds**: children as nu-proprietários, cotistas, and future partners (including
  employees entering partnership). The **usufructuaries Wilson and Elman are entirely outside it** —
  the fruits are theirs by right and are paid in full, without cap or condition.
- No fixed term: the lock is released only by the metas, never by elapsed time; a 60-month review
  does not relax them.

Two open questions are flagged in the documents for the family's lawyer (Dr. Roger Lippi) and should
stay flagged until answered: the effect of a 120-month PGFN transaction on Meta I, and the statutory
mechanism for retaining some partners' dividends while paying the usufructuaries.

## Working conventions

- Branch: `claude/family-property-debt-structure-gvet5m`. Commit messages in Portuguese, matching the
  existing log.
- Drafting voice: clause text is formal legal Portuguese with bracketed placeholders (`[•]`, `[90]`,
  `[IPCA]`) for what the lawyer fills in. Cover notes and drafting notes are plain-spoken.
- Every document carries the disclaimer that it is a working draft, is not a legal opinion, and needs
  review by the family lawyer and a tax specialist before signature. Keep it.
