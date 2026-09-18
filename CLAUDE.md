# CLAUDE.md

Context for working on **testmess**, beyond what the code and README already
show. Read `README.md` first for how the program is put together; this file is
only the things that are easy to break without knowing them.

Feature tag for new files here: `feature: exam-variants` (the tag predates the
rename to testmess — keep it, do not "fix" it to `feature: testmess`).

## The invariants

These are the promises the tests exist to defend. Breaking one silently puts a
wrong exam paper in front of a student, so treat a failure here as a bug in the
change, not in the test.

- **Equations never go through plain text.** The tests are written with the Word
  equation editor (OMML, `<m:oMath>`), which has no faithful text rendering:
  `x²−4 / x−2` flattens to the wrong maths `x2 – 4x – 2`. Stems and options are
  carried as the *raw paragraph XML* copied from the source and re-emitted
  unchanged. The only thing ever rewritten in a paragraph is its leading label.
- **The student copy adds nothing.** Apart from the reordering and the
  renumbering it is paragraph-for-paragraph the source test — no banner, no
  heading, no blank line. Nothing on the page may reveal which variant a student
  holds. (The professor copy is where additions go.)
- **Option markers come from the document.** `(A)`, `a.`, `α)`, `(iii)` — never
  assume A–D, never generate a fresh sequence, never uppercase them. A variant
  reuses that question's own markers in their original order, so a Greek paper
  stays Greek and keeps agreeing with its own instructions line. Compare markers
  with `same_marker()`, which case-folds.
- **The key follows the shuffle by content**, not by letter: after shuffling it
  must point at the same answer text it pointed at in the source.
- **Refuse rather than guess.** A missing key entry, duplicate markers, a key
  naming an option that does not exist: raise, do not improvise. Half-correct
  exam papers are worse than no papers.
- **Overwrites are confirmed once, before anything is written** — a run replaces
  the whole set or touches nothing. Never default to yes when there is nobody to
  ask.

## Conventions

- **Standard library only.** No `pip install`, no virtualenv, no `python-docx`.
  Teachers run this on a Windows laptop straight after installing Python; a
  dependency is a support call. The raw-XML approach is also strictly more
  faithful than any re-render, so this is not a sacrifice.
- 2-space indentation, in the Python too.
- Tests are stdlib `unittest` (no pytest on the target machines). Run them with
  `./run_tests.sh`. Drive the CLI through the tests' `run_cli()` helper so its
  output stays captured and a passing run prints one clear `OK`.
- Do not hard-code facts about `samples/*.docx` into tests — the sample
  documents get re-exported and the numbers move (the equation count already
  changed 30 → 31 once). Derive them from the source document at test time.
- Generated papers (`student_N.docx`, `professor_N.docx`) are disposable output
  and are gitignored anywhere in the tree. The sample sources are not.

## Releases and docs

- Releases are cut by hand on GitHub, with no uploaded assets, so users click
  GitHub's own **Source code (zip)**.
- Keep the README's TL;DR **version-neutral** ("the latest one", "a folder named
  after the version"). It named 0.1 once and was stale within a day.
- The TL;DR is aimed at a teacher who may not have Python installed. Keep it
  short, keep it Windows-flavoured (`python`, backslashes), and only link
  official sources (python.org, Microsoft Learn) for installing Python.

## Verification

Generated documents were confirmed to open correctly in Word on 2026-09-18.
That confirmation is manual and does not repeat itself: there is no Word or
LibreOffice on the development machine, so everything the tests check is
structural — valid archive, every source part present, well-formed XML,
original namespace declarations preserved, equations identical to the source
markup for markup. If a change touches how the package is written, those checks
are necessary but not sufficient; open one of the results in Word before
claiming it renders.
