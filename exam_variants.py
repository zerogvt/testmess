#!/usr/bin/env python3
# feature: exam-variants
"""Build shuffled variants of a multiple-choice test stored as a Word document.

The source .docx holds the questions, the lettered options, and -- on its last
page -- the answer key.  The document is read into one intermediate dictionary
(``parse_exam``), from which one dictionary per variant is derived
(``make_variants``), and only then are the Word documents written.

Equations are the reason this module does not go through plain text.  They are
written with the Word equation editor, i.e. OMML (``<m:oMath>``) markup, which
has no faithful text rendering: "x squared minus 4 over x minus 2" would come
out as "x2 - 4x - 2".  So every stem and every option is carried around as the
*raw paragraph XML* lifted from the source document and re-emitted unchanged;
the only thing rewritten is the leading label ("3." or "(B)").  Everything else
in the package -- styles, fonts, numbering, the math font table -- is copied
from the source archive as-is, so the variants render exactly like the original.

The plain-text fields kept next to the XML (``stem_text``, ``text``) are for
tests, logging and debugging only; they are never used to produce output.

Option markers are whatever the source uses -- "(A)", "a.", "α)" -- and each
variant re-emits that question's own markers in their original order, so a
Greek-lettered test stays Greek-lettered and keeps matching its own rubric.

Stdlib only: zipfile + xml.etree cover the whole job.
"""

from __future__ import annotations

import argparse
import random
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS = {
  'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
DOCUMENT_PART = 'word/document.xml'
XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
# A stem opens with "7." or "7)"; an option with a marker such as "(B)", "b.",
# "α)" or "(iii)"; a key line is a question number and a marker on their own.
#
# A marker is any single letter in any script -- [^\W\d_] is "letter", not
# "A-Za-z", which is what lets Greek (α β γ δ) and Cyrillic work -- or a short
# roman numeral.  Roman numerals are whitelisted rather than allowed as "any
# two or three letters", so an option or stem that merely starts with a short
# word ("No.", "Eq.") is not mistaken for a marker.
MARKER = r'(?:[^\W\d_]|[ivxIVX]{2,4})'
QUESTION_LABEL = re.compile(r'^\s*(\d+)\s*[.)]')
OPTION_LABEL = re.compile(r'^\s*[(\[]?\s*(' + MARKER + r')\s*[.)\]]')
KEY_ENTRY = re.compile(r'^\s*(\d+)\s*[.)]\s*[(\[]?\s*(' + MARKER
                       + r')\s*[)\]]?\s*$')
KEY_HEADING = re.compile(r'^\s*answer\s*key\b', re.IGNORECASE)
ROOT_OPEN_TAG = re.compile(r'<w:document\b[^>]*>')
XMLNS_DECL = re.compile(r'xmlns:([A-Za-z0-9_.-]+)\s*=\s*"([^"]+)"')


def same_marker(left, right):
  """Compare two option markers, ignoring case ("A" == "a", "Α" == "α").

  Markers are stored exactly as the document writes them -- uppercasing them
  would turn a Greek "α)" test into an "Α)" one on the way out -- so any
  comparison between them has to be case-folded here.
  """
  return left.casefold() == right.casefold()


def qn(prefix, tag):
  """Qualified tag name, e.g. qn('w', 'p') -> '{...wordprocessingml...}p'."""
  return '{%s}%s' % (NS[prefix], tag)


# --------------------------------------------------------------------------
# XML helpers
# --------------------------------------------------------------------------

def paragraph_text(paragraph):
  """Flatten a paragraph to plain text, equation runs (<m:t>) included.

  Lossy by nature -- used for recognising labels and for test assertions,
  never for writing documents.
  """
  parts = []
  for node in paragraph.iter():
    if node.tag in (qn('w', 't'), qn('m', 't')):
      parts.append(node.text or '')
    elif node.tag == qn('w', 'tab'):
      parts.append('\t')
  return ''.join(parts)


def has_page_break(paragraph):
  for node in paragraph.iter(qn('w', 'br')):
    if node.get(qn('w', 'type')) == 'page':
      return True
  return False


def set_run_text(node, value):
  """Set a <w:t> value, keeping significant leading/trailing whitespace."""
  node.text = value
  if value != value.strip():
    node.set(XML_SPACE, 'preserve')


def replace_span(nodes, start, end, replacement):
  """Replace characters [start, end) of the concatenated <w:t> texts.

  A label can be split over several runs ("(A" + ")  4"), so the span is
  mapped back onto whichever runs it covers.  Only those runs are touched,
  which keeps the surrounding character formatting intact.
  """
  position = 0
  inserted = False
  for node in nodes:
    text = node.text or ''
    low, high = position, position + len(text)
    position = high
    if high <= start or low >= end:
      continue
    cut_from = max(start, low) - low
    cut_to = min(end, high) - low
    head = text[:cut_from]
    tail = text[cut_to:]
    set_run_text(node, head + ('' if inserted else replacement) + tail)
    inserted = True


def relabel(paragraph, pattern, new_label):
  """Rewrite the label opening a paragraph ("3." -> "7.", "(B)" -> "(A)")."""
  nodes = list(paragraph.iter(qn('w', 't')))
  joined = ''.join(node.text or '' for node in nodes)
  match = pattern.match(joined)
  if match is None:
    return False
  replace_span(nodes, match.start(1), match.end(1), new_label)
  return True


def strip_paragraph_ids(paragraph):
  """Drop w14:paraId / w14:textId from a cloned paragraph.

  Those ids are meant to be unique per paragraph, and the answer key lines are
  all stamped out of the same template paragraph.
  """
  for name in ('paraId', 'textId'):
    paragraph.attrib.pop(qn('w14', name), None)
  return paragraph


def paragraph_xml(paragraph):
  """Serialise a paragraph, namespace declarations included, for the dict."""
  return ET.tostring(paragraph, encoding='unicode')


def parse_paragraph(xml_text):
  return ET.fromstring(xml_text)


def read_document_xml(docx_path):
  with zipfile.ZipFile(docx_path) as archive:
    return archive.read(DOCUMENT_PART).decode('utf-8')


def register_namespaces(xml_text):
  """Teach ElementTree the document's own prefixes (w:, m:, w14: ...).

  Without this the serialiser invents ns0/ns1 prefixes, which breaks the
  prefix list in the root's mc:Ignorable attribute.
  """
  match = ROOT_OPEN_TAG.search(xml_text)
  head = match.group(0) if match else xml_text[:4000]
  for prefix, uri in XMLNS_DECL.findall(head):
    ET.register_namespace(prefix, uri)
  return match.group(0) if match else None


# --------------------------------------------------------------------------
# Reading the source test
# --------------------------------------------------------------------------

def parse_exam(docx_path):
  """Read the test into the intermediate dictionary.

  Returns::

    {
      'source': str,
      'title': str,
      'preamble_xml': [paragraph xml, ...],     # everything above question 1
      'key_templates': {'break': xml|None, 'heading': xml|None,
                        'entry': xml|None},
      'questions': [
        {'number': int, 'stem_text': str, 'stem_xml': [xml, ...],
         'answer': 'A',
         'options': [{'letter': 'A', 'text': str, 'xml': [xml, ...]}, ...]},
        ...
      ],
    }
  """
  docx_path = Path(docx_path)
  xml_text = read_document_xml(docx_path)
  register_namespaces(xml_text)
  body = ET.fromstring(xml_text).find(qn('w', 'body'))

  preamble = []
  questions = []
  key_templates = {'break': None, 'heading': None, 'entry': None}
  answers = {}
  section = 'preamble'
  current = None
  previous = None

  for child in body:
    if child.tag != qn('w', 'p'):
      continue
    text = paragraph_text(child).strip()

    if KEY_HEADING.match(text):
      section = 'key'
      key_templates['heading'] = paragraph_xml(child)
      if previous is not None and has_page_break(previous):
        key_templates['break'] = paragraph_xml(previous)
      previous = child
      continue

    if section == 'key':
      match = KEY_ENTRY.match(text)
      if match:
        answers[int(match.group(1))] = match.group(2)
        if key_templates['entry'] is None:
          key_templates['entry'] = paragraph_xml(child)
      previous = child
      continue

    question_match = QUESTION_LABEL.match(text)
    option_match = OPTION_LABEL.match(text) if current is not None else None

    if question_match:
      section = 'questions'
      current = {
        'number': int(question_match.group(1)),
        'stem_text': text[question_match.end():].strip(),
        'stem_xml': [paragraph_xml(child)],
        'options': [],
        'answer': None,
      }
      questions.append(current)
    elif option_match:
      current['options'].append({
        'letter': option_match.group(1),
        'text': text[option_match.end():].strip(),
        'xml': [paragraph_xml(child)],
      })
    elif section == 'preamble':
      preamble.append(paragraph_xml(child))
    elif text and current is not None:
      # A stem or an option that runs over several paragraphs: keep it with
      # whatever it continues, so nothing is dropped.
      target = current['options'][-1] if current['options'] else current
      if 'stem_xml' in target:
        target['stem_xml'].append(paragraph_xml(child))
        target['stem_text'] = (target['stem_text'] + ' ' + text).strip()
      else:
        target['xml'].append(paragraph_xml(child))
        target['text'] = (target['text'] + ' ' + text).strip()

    previous = child

  for question in questions:
    question['answer'] = answers.get(question['number'])

  exam = {
    'source': str(docx_path),
    'title': _first_text(preamble),
    'preamble_xml': preamble,
    'key_templates': key_templates,
    'questions': questions,
  }
  validate_exam(exam)
  return exam


def _first_text(preamble_xml):
  if not preamble_xml:
    return ''
  return paragraph_text(parse_paragraph(preamble_xml[0])).strip()


def validate_exam(exam):
  """Fail loudly rather than emit a test with a wrong or missing key."""
  if not exam['questions']:
    raise ValueError('no questions found in %s' % exam['source'])
  for question in exam['questions']:
    number = question['number']
    if len(question['options']) < 2:
      raise ValueError('question %d has %d option(s)'
                       % (number, len(question['options'])))
    markers = [option['letter'] for option in question['options']]
    folded = [marker.casefold() for marker in markers]
    if len(set(folded)) != len(folded):
      raise ValueError('question %d has duplicate option markers: %s'
                       % (number, ', '.join(markers)))
    if question['answer'] is None:
      raise ValueError('no answer key entry for question %d' % number)
    if not any(same_marker(question['answer'], marker) for marker in markers):
      raise ValueError('key for question %d is %s, which is not one of %s'
                       % (number, question['answer'], ', '.join(markers)))


# --------------------------------------------------------------------------
# Building the variants
# --------------------------------------------------------------------------

def make_variant(exam, index, rng):
  """Derive one variant dictionary: questions reordered, options reshuffled.

  The XML payloads are shared with ``exam`` -- they are never mutated, only
  copied at render time -- while numbers, markers and the answer are new.

  Markers are not generated, they are reused: slot 1 of a variant carries
  whatever marker slot 1 of that question carried in the source.  A test
  lettered (A)-(D) stays Latin, one lettered α)-δ) stays Greek, and the
  document keeps agreeing with its own instructions line.
  """
  order = list(range(len(exam['questions'])))
  rng.shuffle(order)

  questions = []
  for position, source_index in enumerate(order, start=1):
    source = exam['questions'][source_index]
    markers = [option['letter'] for option in source['options']]
    shuffled = list(source['options'])
    rng.shuffle(shuffled)

    options = []
    answer = None
    for slot, option in enumerate(shuffled):
      marker = markers[slot]
      options.append({
        'letter': marker,
        'source_letter': option['letter'],
        'text': option['text'],
        'xml': option['xml'],
      })
      if same_marker(option['letter'], source['answer']):
        answer = marker

    questions.append({
      'number': position,
      'source_number': source['number'],
      'stem_text': source['stem_text'],
      'stem_xml': source['stem_xml'],
      'options': options,
      'answer': answer,
    })

  return {'index': index, 'source': exam['source'], 'questions': questions}


def make_variants(exam, count, seed=None):
  rng = random.Random(seed)
  variants = []
  for index in range(1, count + 1):
    variants.append(make_variant(exam, index, rng))
  return variants


# --------------------------------------------------------------------------
# Writing the variant documents
# --------------------------------------------------------------------------

def _element(tag, **attributes):
  node = ET.Element(qn('w', tag))
  for name, value in attributes.items():
    node.set(qn('w', name), value)
  return node


def _text_paragraph(text, bold=False, centered=False, size=None, after='60'):
  paragraph = _element('p')
  properties = _element('pPr')
  if centered:
    properties.append(_element('jc', val='center'))
  properties.append(_element('spacing', after=after))
  paragraph.append(properties)
  run = _element('r')
  run_properties = _element('rPr')
  if bold:
    run_properties.append(_element('b'))
    run_properties.append(_element('bCs'))
  if size:
    run_properties.append(_element('sz', val=size))
    run_properties.append(_element('szCs', val=size))
  if len(run_properties):
    run.append(run_properties)
  node = _element('t')
  set_run_text(node, text)
  run.append(node)
  paragraph.append(run)
  return paragraph


def _page_break_paragraph():
  paragraph = _element('p')
  run = _element('r')
  run.append(_element('br', type='page'))
  paragraph.append(run)
  return paragraph


def _banner_paragraph(variant):
  """A "Variant N" line for the professor copy only.

  The student copy gets no banner, and nothing else the source did not have:
  a variant has to read as the same test its class was given, not as one
  visibly stamped copy out of several.  Which variant a student paper is can
  be read off its file name, or off the professor copy it matches.
  """
  label = 'Variant %d - Professor copy (with answer key)' % variant['index']
  return _text_paragraph(label, bold=True, centered=True, size='22', after='160')


def _key_entry_paragraph(exam, number, letter):
  """Build one key line, reusing the source key line's formatting."""
  template = exam['key_templates']['entry']
  if template is not None:
    paragraph = strip_paragraph_ids(parse_paragraph(template))
    nodes = list(paragraph.iter(qn('w', 't')))
    joined = ''.join(node.text or '' for node in nodes)
    match = KEY_ENTRY.match(joined)
    if match is not None:
      # Right to left: rewriting the letter first keeps the number's offsets.
      replace_span(nodes, match.start(2), match.end(2), letter)
      replace_span(nodes, match.start(1), match.end(1), str(number))
      return paragraph
  return _text_paragraph('%d.  %s' % (number, letter), after='40')


def _key_paragraphs(exam, variant):
  templates = exam['key_templates']
  paragraphs = []
  if templates['break'] is not None:
    paragraphs.append(parse_paragraph(templates['break']))
  else:
    paragraphs.append(_page_break_paragraph())
  if templates['heading'] is not None:
    paragraphs.append(parse_paragraph(templates['heading']))
  else:
    paragraphs.append(_text_paragraph('Answer Key', bold=True, size='28',
                                      after='160'))
  for question in variant['questions']:
    paragraphs.append(_key_entry_paragraph(exam, question['number'],
                                           question['answer']))
  return paragraphs


def render_document_xml(exam, variant, include_key):
  """Rebuild word/document.xml for one variant, reusing the source markup."""
  xml_text = read_document_xml(exam['source'])
  open_tag = register_namespaces(xml_text)
  root = ET.fromstring(xml_text)
  body = root.find(qn('w', 'body'))
  section_properties = body.find(qn('w', 'sectPr'))
  for child in list(body):
    body.remove(child)

  for item in exam['preamble_xml']:
    body.append(parse_paragraph(item))
  if include_key:
    body.append(_banner_paragraph(variant))

  for question in variant['questions']:
    stem = [parse_paragraph(item) for item in question['stem_xml']]
    relabel(stem[0], QUESTION_LABEL, str(question['number']))
    body.extend(stem)
    for option in question['options']:
      paragraphs = [parse_paragraph(item) for item in option['xml']]
      relabel(paragraphs[0], OPTION_LABEL, option['letter'])
      body.extend(paragraphs)

  if include_key:
    body.extend(_key_paragraphs(exam, variant))
  if section_properties is not None:
    body.append(section_properties)

  serialised = ET.tostring(root, encoding='unicode')
  if open_tag is not None:
    # Restore the original root tag so every prefix listed in mc:Ignorable
    # stays declared, even where this variant happens not to use it.
    serialised = ROOT_OPEN_TAG.sub(lambda _match: open_tag, serialised, count=1)
  return XML_DECLARATION + serialised.encode('utf-8')


def write_docx(source_path, destination, document_xml):
  """Copy the source package, swapping in the rebuilt document part."""
  destination = Path(destination)
  with zipfile.ZipFile(source_path) as source_zip:
    with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as out_zip:
      for item in source_zip.infolist():
        if item.filename == DOCUMENT_PART:
          out_zip.writestr(item, document_xml)
        else:
          out_zip.writestr(item, source_zip.read(item.filename))
  return destination


def variant_paths(out_dir, count):
  """Every file a run would write, in the order it would write them."""
  out_dir = Path(out_dir)
  paths = []
  for index in range(1, count + 1):
    for name in ('student', 'professor'):
      paths.append(out_dir / ('%s_%d.docx' % (name, index)))
  return paths


def confirm_overwrite(paths, force=False, ask=None, stream=None):
  """Ask once, before anything is written, if any target already exists.

  Asking up front rather than per file means a run either replaces the whole
  set or touches nothing -- a half-overwritten folder, where student_1 is from
  this run and student_2 from the last one, is the one outcome that would
  actually put the wrong paper in front of a student.

  ``ask`` is the prompt function (``input``), injected so the tests can answer.
  """
  existing = [path for path in paths if path.exists()]
  if not existing or force:
    return True

  # The listing belongs with the question, so it goes where input() writes
  # its prompt; only the failure messages go to stderr.
  stream = stream or sys.stdout
  shown = existing[:6]
  print('These files already exist in %s:' % (existing[0].parent or '.'),
        file=stream)
  for path in shown:
    print('  %s' % path.name, file=stream)
  if len(existing) > len(shown):
    print('  ... and %d more' % (len(existing) - len(shown)), file=stream)

  if ask is None:
    ask = input
  try:
    answer = ask('Overwrite %d file(s)? [y/N] ' % len(existing))
  except EOFError:
    # Nothing on stdin: a cron job, a pipe that ran dry, a closed terminal.
    # Never guess "yes" on somebody's behalf here.
    print('error: nobody to ask (no input); re-run with --force to overwrite, '
          'or --out-dir to write elsewhere', file=sys.stderr)
    return False
  except KeyboardInterrupt:
    print(file=stream)
    return False
  return answer.strip().lower() in ('y', 'yes')


def write_variant(exam, variant, out_dir='.'):
  """Write student_N.docx and professor_N.docx; return both paths."""
  out_dir = Path(out_dir)
  out_dir.mkdir(parents=True, exist_ok=True)
  written = []
  for name, include_key in (('student', False), ('professor', True)):
    target = out_dir / ('%s_%d.docx' % (name, variant['index']))
    written.append(write_docx(exam['source'], target,
                              render_document_xml(exam, variant, include_key)))
  return tuple(written)


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------

def parse_args(argv=None):
  parser = argparse.ArgumentParser(
    description='Create shuffled student/professor variants of a '
                'multiple-choice Word test.')
  parser.add_argument('docx', help='source .docx test, answer key on the last page')
  parser.add_argument('-n', '--variants', type=int, default=3,
                      help='number of variants to create (default: 3)')
  parser.add_argument('--seed', type=int, default=None,
                      help='random seed, for reproducible variants')
  parser.add_argument('--out-dir', default='.',
                      help='where to write the documents (default: current directory)')
  parser.add_argument('-f', '--force', action='store_true',
                      help='overwrite existing student_N.docx / professor_N.docx '
                           'without asking')
  return parser.parse_args(argv)


def main(argv=None, ask=None):
  args = parse_args(argv)
  if args.variants < 1:
    print('error: --variants must be at least 1', file=sys.stderr)
    return 2
  try:
    exam = parse_exam(args.docx)
  except (OSError, KeyError, ET.ParseError, ValueError) as error:
    print('error: could not read %s: %s' % (args.docx, error), file=sys.stderr)
    return 1

  targets = variant_paths(args.out_dir, args.variants)
  if not confirm_overwrite(targets, force=args.force, ask=ask):
    print('aborted: nothing was written', file=sys.stderr)
    return 3

  print('Read %d questions from %s' % (len(exam['questions']), args.docx))
  for variant in make_variants(exam, args.variants, args.seed):
    student, professor = write_variant(exam, variant, args.out_dir)
    key = ', '.join('%d%s' % (q['number'], q['answer'])
                    for q in variant['questions'])
    print('Variant %d -> %s, %s   key: %s'
          % (variant['index'], student.name, professor.name, key))
  return 0


if __name__ == '__main__':
  sys.exit(main())
