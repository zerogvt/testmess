<!-- feature: exam-variants -->
# exam_variants

Turn one multiple-choice test in Word into any number of shuffled variants, each
written twice: a student copy (questions + options) and a professor copy (same
paper, plus a "Variant N" banner and the answer key on its own last page).

**The student copy adds nothing.** Apart from the reordering and the
renumbering it is paragraph-for-paragraph the source test — no banner, no
heading, no blank line — so nothing on the page tells a student which copy they
are holding. Which variant a paper is can be read off its file name, or off the
professor copy that matches it.

Questions are reordered and the options inside each question are reshuffled, so
neighbours cannot copy each other. The key is recomputed to follow the shuffle —
it always points at the same *answer content*, never at the old marker.

Option markers are read from the document rather than assumed: `(A)`, `a.`,
`α)` and `(iii)` all work, and a variant re-emits that question's own markers in
their original order, so a Greek-lettered test stays Greek-lettered and keeps
agreeing with its own instructions line.

Equations written with the Word equation editor survive untouched. See
[Why raw XML](#why-raw-xml-and-not-python-docx) for how, and why that drove the
design.

## Requirements

Python 3.8+. Nothing else — standard library only (`zipfile`, `xml.etree`,
`random`, `argparse`). No virtualenv, no `pip install`.

## Running it

```bash
python3 exam_variants.py samples/calculus_practice_test_2.docx
```

That reads the test and writes 6 files into the current directory:

```
student_1.docx  professor_1.docx
student_2.docx  professor_2.docx
student_3.docx  professor_3.docx
```

Options:

```
python3 exam_variants.py DOCX [-n VARIANTS] [--seed SEED] [--out-dir DIR] [-f]

  DOCX                 source .docx test, answer key on the last page
  -n, --variants N     how many variants to create      (default: 3)
  --seed SEED          random seed, for reproducible variants (default: random)
  --out-dir DIR        where to write the documents      (default: .)
  -f, --force          overwrite existing papers without asking
```

Examples:

```bash
# five variants
python3 exam_variants.py samples/calculus_practice_test_2.docx -n 5

# reproducible: the same seed always produces exactly the same papers,
# which is what you want if you have to reprint one of them later
python3 exam_variants.py samples/calculus_practice_test_2.docx -n 5 --seed 2026

# somewhere other than the current folder
python3 exam_variants.py samples/calculus_practice_test_2.docx --out-dir ./exams
```

It prints what it did, including each variant's key, so you can check a paper
without opening it:

```
Read 10 questions from samples/calculus_practice_test_2.docx
Variant 1 -> student_1.docx, professor_1.docx   key: 1C, 2C, 3A, 4A, 5D, ...
```

### Overwriting existing papers

If any `student_N.docx` or `professor_N.docx` is already in the output folder,
the run stops and asks before writing anything:

```
These files already exist in .:
  student_1.docx
  professor_1.docx
Overwrite 2 file(s)? [y/N]
```

Anything other than `y`/`yes` aborts and **nothing** is written. The question is
asked once for the whole run, not per file: a folder where `student_1.docx` is
from this run and `student_2.docx` from the previous one is the one outcome that
would actually put the wrong paper in front of a student, so a run either
replaces the whole set or touches nothing.

With no input available at all — cron, a closed terminal, `< /dev/null` — it
refuses rather than guessing, and says so. Two ways past it:

```bash
python3 exam_variants.py samples/test.docx --force        # yes, replace them
python3 exam_variants.py samples/test.docx --out-dir v2   # or just write elsewhere
```

Exit codes: `0` success, `1` the source could not be read or parsed, `2` bad
arguments, `3` declined the overwrite (nothing written).

## What the source document must look like

The parser follows the layout of `samples/calculus_practice_test_2.docx`:

- Any front matter (title, instructions) above the first question — copied to
  every variant as-is.
- A question starts a paragraph with `1.` or `1)`.
- Each option is its own paragraph starting with a marker: `(A)`, `A)`, `A.`,
  `a)`, `[A]`, `α)`, `(iii)` — any single letter in any script (Latin, Greek,
  Cyrillic, either case) or a short roman numeral.
- A last page headed **Answer Key**, one entry per paragraph: `1.  A`, `1)  α`,
  `1.  (c)`. The entries may be in any order; they are matched by question
  number. Key and options need not agree on case (`Α` finds `α`).

Markers are stored exactly as written and compared case-insensitively — see
`same_marker()`. Uppercasing them would turn an `α)` paper into an `Α)` one on
the way out. Digits are deliberately *not* accepted as option markers, since
`1)` cannot be told apart from a question number.

Stems and options may run over several paragraphs; the extra paragraphs stay
attached to whatever they continue. Anything that does not fit the test is
rejected loudly rather than guessed at: a question with fewer than two options,
duplicate option letters, a missing key entry, or a key pointing at a letter
that does not exist all raise `ValueError` (see `validate_exam`).

## Architecture

Three stages, one dictionary between each — the dictionaries are the contract,
which is what makes the whole thing testable without ever opening Word.

```
samples/calculus_practice_test_2.docx
        |
        |  parse_exam()            read the .docx, recognise questions,
        v                          options and the key
   exam dict  ────────────────────────────────────────────────┐
        |                                                     |
        |  make_variants(exam, n, seed)                       |  (paragraph XML
        v      make_variant() once per variant                |   is shared, and
   variant dict  x n                                          |   never mutated)
        |                                                     |
        |  write_variant(exam, variant, out_dir)              |
        |      render_document_xml(include_key=False) ────────┤
        |      render_document_xml(include_key=True)  ────────┘
        v
   student_N.docx + professor_N.docx
```

### 1. `parse_exam(path) -> dict`

Reads `word/document.xml` out of the archive and walks the body paragraphs,
classifying each one by its opening label. Every stem and every option keeps
**its own paragraph XML**, plus a plain-text rendering used only for tests,
logging and label matching.

```python
{
  'source': 'samples/calculus_practice_test_2.docx',
  'title': 'Calculus Practice Test',
  'preamble_xml': [<xml>, ...],            # everything above question 1
  'key_templates': {'break': <xml>,        # page break, key heading and one key
                    'heading': <xml>,      # line, reused so the professor copy
                    'entry': <xml>},       # keeps the original formatting
  'questions': [
    {'number': 1,
     'stem_text': 'Evaluate the limit: ...',
     'stem_xml': [<xml>, ...],
     'options': [{'letter': 'A',      # marker, verbatim
                  'text': '4', 'xml': [<xml>, ...]}, ...],
     'answer': 'A'},
    ...
  ],
}
```

### 2. `make_variants(exam, count, seed) -> [dict]`

Pure data, no I/O: shuffles the question order and each question's options with
a seeded `random.Random`, renumbers the questions from 1, and recomputes the
answer marker. Markers are not generated but reused: slot 1 of a variant carries
whatever marker slot 1 of that question carried in the source, which is what
keeps a Greek paper Greek. The XML payloads are shared with the exam dict and
copied only at render time, so building a variant cannot corrupt the source.

```python
{
  'index': 1,
  'source': 'samples/calculus_practice_test_2.docx',
  'questions': [
    {'number': 1,            # position in this variant
     'source_number': 7,     # which original question it is
     'stem_text': ..., 'stem_xml': [...],
     'options': [{'letter': 'A',          # marker in this variant
                  'source_letter': 'C',   # marker in the original
                  'text': ..., 'xml': [...]}, ...],
     'answer': 'B'},         # marker of the correct option, in this variant
    ...
  ],
}
```

`source_number` and `source_letter` are the provenance of every shuffle. They
are what lets the tests assert that the new key still points at the same answer
*content* as the original, rather than merely at some letter.

### 3. `write_variant(exam, variant, out_dir)`

Called twice per variant, with and without the key. Before any of this runs,
`main()` asks `confirm_overwrite()` about the full list of targets from
`variant_paths()` — see [Overwriting existing papers](#overwriting-existing-papers).
The prompt function is injected rather than calling `input()` directly, which is
what lets the tests answer it.

`render_document_xml()` re-parses the source document, empties the body, and
refills it: the source's own front matter, then each question's paragraphs in
their new order, then — professor copy only — a "Variant N" banner, the page
break, the **Answer Key** heading and one line per question. `sectPr` (page setup) is
put back last, where Word expects it.

`write_docx()` then copies the source archive part for part, swapping in only
the rebuilt `word/document.xml`. Styles, fonts, numbering and the Cambria Math
font table are therefore literally the originals, so a variant renders exactly
like the source test.

### Why raw XML, and not `python-docx`

The equations are OMML (`<m:oMath>` markup from the Word equation editor), and
OMML has no faithful text form. Flattening question 1 gives `x2 – 4x – 2`:
the fraction bar and the exponent are simply gone, and that string is wrong
maths. Rebuilding equations from text would mean re-deriving structure that was
never in the text.

So nothing is ever re-rendered. Each stem and option travels as the raw
paragraph XML lifted from the source, and the **only** thing rewritten is the
leading label — `3.` becomes `7.`, `(B)` becomes `(A)`. `relabel()` locates that
label across the concatenated `<w:t>` runs (Word routinely splits `(A)  4` into
`(A` + `)  4`) and edits only the runs it actually covers, leaving bold labels
and adjacent formatting intact. Because the replacement is a string, a marker
may change width (`iii` → `i`) without disturbing anything around it. Equation
runs are `<m:t>`, never `<w:t>`, so a label rewrite cannot reach inside an
equation.

This also removes the dependency: the standard library is enough, and the
result is more faithful than a re-render would have been. Verified on the
sample test — all 31 `<m:oMath>` nodes are canonically identical in every
generated document.

## Tests

58 tests, `unittest` from the standard library (no pytest needed):

```bash
python3 -m unittest test_exam_variants        # quiet
python3 -m unittest test_exam_variants -v     # per-test
python3 test_exam_variants.py                 # same, verbose
```

Most of them run against the dictionaries — parsing, the key page not being read
as an eleventh question, permutation and renumbering invariants, the correct
answer following the shuffle by content, seed determinism, the exam dict not
being mutated, and label rewriting split across runs. Others are end-to-end over
the written `.docx` files: the student copy has no key, the professor key matches
its variant, the equation count is preserved, every part of the package is still
there and paragraph ids stay unique.

`FunctionalTest` runs the whole pipeline — `.docx` in, `.docx` out, through the
command line — over **both** sample tests, the Latin-lettered
`samples/calculus_practice_test_2.docx` and the Greek-lettered
`samples/calculus_practice_test_3.docx`, asserting the same things of each via
`subTest` — including that every equation in a variant is identical to the
source's, markup for markup, compared independently of how the namespace
prefixes happen to be spelled. Nothing about the source documents is
hard-coded that the documents themselves can answer: the expected equation
count is read from the source, so re-exporting a test does not break the
suite. That is what stops anything quietly depending on the markers being
`A`–`D`. `MarkerRecognitionTest` pins the label regexes down directly, including
the negative cases: prose such as `Does not exist`, `Note: assume x > 0` or
`Eq. 4 applies here` must *not* be read as an option marker. `OverwriteTest`
covers the confirmation: which answers count as yes, that declining leaves every
existing file byte-for-byte untouched, that `--force` skips the question, and
that no input at all is treated as no. `test_student_copy_introduces_nothing_new`
compares the student paper paragraph for paragraph against the source, counts
included, so no banner or stray blank line can creep back in.

## Files

| File | |
|---|---|
| `exam_variants.py` | the program: parse, shuffle, write |
| `test_exam_variants.py` | the test suite |
| `samples/calculus_practice_test_2.docx` | sample source test, Latin markers `(A)`–`(D)` |
| `samples/calculus_practice_test_3.docx` | same test, Greek markers `α)`–`δ)` |
| `student_N.docx`, `professor_N.docx` | generated output, written to `--out-dir` |

Every `.docx` lives under `samples/`, sources and generated papers alike. The
output folder is still `--out-dir`, which defaults to the current directory, so
add `--out-dir samples` if you want new runs to land there too.

## Known limitation

The generated documents have been checked structurally (valid archive, all parts
present, well-formed XML, original namespace declarations preserved, equations
byte-identical), but neither Word nor LibreOffice was available here, so they
have not been opened in a word processor. Worth one look before printing.
