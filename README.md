# ICOTS paper template for Quarto

A [Quarto](https://quarto.org) extension for formatting papers submitted to [ICOTS (International Conference on Teaching Statistics)](https://icots12.oa-event.com/). The format produces a `.docx` output styled to ICOTS paper requirements.

> [!NOTE]
> This template was created with the assistance of [Claude](https://claude.ai) (Anthropic).

## Installing

```bash
quarto use template mitchelloharawild/quarto-icots
```

This will install the extension and create an example qmd file that you can use as a starting place for your article.

## Using

Once installed, set the document format to `quarto-icots-docx` in your YAML front matter:

```yaml
format:
  quarto-icots-docx: default
```

Then author your paper using standard Quarto markdown. Render with:

```bash
quarto render your-paper.qmd
```

## Format Options

The format is configured via standard Quarto YAML fields:

| Field | Description |
|---|---|
| `title` | Paper title |
| `author` | List of authors with `name`, `email`, and `affiliations` |
| `abstract` | Abstract text (max 125 words; do **not** begin with the word "Abstract") |
| `date` | Submission date (e.g. `last-modified`) |

Authors support the following sub-fields:

```yaml
author:
  - name: "Author Name"
    email: "author@example.com"
    affiliations:
      - name: "Institution, Country"
```

## Known Issues

- **Figure and table captions are not centered**: Despite my best efforts, captions are not automatically centered in the output `.docx`. Before submission, you will need to manually center these captions in Word. See [#1](https://github.com/mitchelloharawild/quarto-icots/issues/1) for details.

## Example

Here is the source code for a minimal sample document: [template.qmd](template.qmd).

````quarto
---
title: Title
format:
  quarto-icots-docx: default
author:
  - name: "Author1"
    email: "Contact email"
    affiliations:
      - name: "Address1"
  - name: "Author2"
    affiliations:
      - name: "Address2"
abstract: >
  Put your abstract here: this text is not to exceed 125 words. Do **not** start
  the abstract with the word "Abstract".
date: last-modified
---

## My First Heading

Text of paper...
````