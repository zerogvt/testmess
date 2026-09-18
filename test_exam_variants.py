#!/usr/bin/env python3
# feature: exam-variants
"""Unit tests for exam_variants (stdlib unittest; no pytest in this env).

Most assertions run against the intermediate dictionaries -- that is what they
are there for -- with a smaller set of end-to-end checks on the written .docx
packages, including a count of the equation nodes so a lost <m:oMath> fails the
build instead of quietly reaching a student.
"""

import os
import random
import re
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path

import exam_variants as ev

SAMPLES = Path(__file__).parent / 'samples'
SOURCE = SAMPLES / 'calculus_practice_test_2.docx'
EXPECTED_KEY = {1: 'A', 2: 'D', 3: 'C', 4: 'A', 5: 'C',
                6: 'B', 7: 'D', 8: 'A', 9: 'B', 10: 'C'}

# The same test in two marker alphabets.  Test 3 marks its options with Greek
# letters and no opening bracket ("α)  4"), and its key reads "1.  α".
FIXTURES = (
  {'name': 'latin',
   'path': SOURCE,
   'markers': ['A', 'B', 'C', 'D'],
   'key': EXPECTED_KEY},
  {'name': 'greek',
   'path': SAMPLES / 'calculus_practice_test_3.docx',
   'markers': ['α', 'β', 'γ', 'δ'],
   'key': {1: 'α', 2: 'δ', 3: 'γ', 4: 'α', 5: 'γ',
           6: 'β', 7: 'δ', 8: 'α', 9: 'β', 10: 'γ'}},
)


def document_root(docx_path):
  return ET.fromstring(ev.read_document_xml(docx_path))


def count_math(xml_root):
  return len(list(xml_root.iter(ev.qn('m', 'oMath'))))


def strip_label(text):
  """Drop a leading "7." or "(B)" so a paragraph can be compared by content."""
  for pattern in (ev.QUESTION_LABEL, ev.OPTION_LABEL):
    match = pattern.match(text)
    if match:
      return text[match.end():].strip()
  return text.strip()


def math_signature(xml_root):
  """Every equation's full structure, as a multiset.

  Tags and attribute names come back in Clark notation ({uri}local), so this
  compares the markup itself and not how a serialiser happened to spell the
  namespace prefixes -- two documents can be identical and still write m: vs
  ns0: for the same namespace.
  """
  def shape(element):
    children = tuple(shape(child) for child in element)
    return (element.tag, tuple(sorted(element.attrib.items())),
            (element.text or '').strip(), children)

  return Counter(shape(node) for node in xml_root.iter(ev.qn('m', 'oMath')))


def body_paragraphs(xml_root):
  body = xml_root.find(ev.qn('w', 'body'))
  return [p for p in body if p.tag == ev.qn('w', 'p')]


def paragraph_texts(xml_root):
  return [ev.paragraph_text(p).strip() for p in body_paragraphs(xml_root)]


class ParseExamTest(unittest.TestCase):

  @classmethod
  def setUpClass(cls):
    cls.exam = ev.parse_exam(SOURCE)

  def test_reads_every_question(self):
    numbers = [q['number'] for q in self.exam['questions']]
    self.assertEqual(numbers, list(range(1, 11)))

  def test_every_question_has_four_lettered_options(self):
    for question in self.exam['questions']:
      letters = [option['letter'] for option in question['options']]
      self.assertEqual(letters, ['A', 'B', 'C', 'D'],
                       'question %d' % question['number'])

  def test_reads_the_answer_key(self):
    found = {q['number']: q['answer'] for q in self.exam['questions']}
    self.assertEqual(found, EXPECTED_KEY)

  def test_key_page_is_not_mistaken_for_questions(self):
    # "1.  A" on the key page must not become an eleventh question.
    self.assertEqual(len(self.exam['questions']), 10)
    self.assertIsNotNone(self.exam['key_templates']['heading'])
    self.assertIsNotNone(self.exam['key_templates']['entry'])

  def test_labels_are_stripped_from_the_stored_text(self):
    first = self.exam['questions'][0]
    self.assertTrue(first['stem_text'].startswith('Evaluate the limit'))
    self.assertEqual(first['options'][0]['text'], '4')
    self.assertEqual(first['options'][3]['text'], 'Does not exist')

  def test_equation_markup_is_kept_verbatim(self):
    # Question 1's stem holds the limit; question 4's options are equations.
    stem = ET.fromstring(self.exam['questions'][0]['stem_xml'][0])
    self.assertEqual(count_math(stem), 1)
    for option in self.exam['questions'][3]['options']:
      node = ET.fromstring(option['xml'][0])
      self.assertEqual(count_math(node), 1, option['text'])
    # The raw OMML, not a text rendering: the fraction is still a fraction.
    self.assertIn('oMath', self.exam['questions'][0]['stem_xml'][0])
    self.assertIn('<m:f>', self.exam['questions'][0]['stem_xml'][0]
                  .replace('ns0:', 'm:'))

  def test_preamble_is_captured(self):
    self.assertEqual(self.exam['title'], 'Calculus Practice Test')
    self.assertEqual(len(self.exam['preamble_xml']), 3)


class ValidationTest(unittest.TestCase):

  def setUp(self):
    self.exam = ev.parse_exam(SOURCE)

  def test_missing_key_entry_is_rejected(self):
    self.exam['questions'][2]['answer'] = None
    with self.assertRaises(ValueError):
      ev.validate_exam(self.exam)

  def test_key_pointing_at_a_missing_option_is_rejected(self):
    self.exam['questions'][2]['answer'] = 'Z'
    with self.assertRaises(ValueError):
      ev.validate_exam(self.exam)

  def test_question_without_options_is_rejected(self):
    self.exam['questions'][0]['options'] = []
    with self.assertRaises(ValueError):
      ev.validate_exam(self.exam)


class MakeVariantTest(unittest.TestCase):

  @classmethod
  def setUpClass(cls):
    cls.exam = ev.parse_exam(SOURCE)

  def variants(self, count=5, seed=7):
    return ev.make_variants(self.exam, count, seed)

  def test_default_count_is_three(self):
    # The CLI default, checked where it is declared.
    self.assertEqual(ev.parse_args([str(SOURCE)]).variants, 3)

  def test_questions_are_a_permutation_renumbered_from_one(self):
    for variant in self.variants():
      numbers = [q['number'] for q in variant['questions']]
      sources = sorted(q['source_number'] for q in variant['questions'])
      self.assertEqual(numbers, list(range(1, 11)))
      self.assertEqual(sources, list(range(1, 11)))

  def test_options_are_a_permutation_relettered_from_a(self):
    for variant in self.variants():
      for question in variant['questions']:
        letters = [o['letter'] for o in question['options']]
        sources = sorted(o['source_letter'] for o in question['options'])
        self.assertEqual(letters, ['A', 'B', 'C', 'D'])
        self.assertEqual(sources, ['A', 'B', 'C', 'D'])

  def test_option_contents_are_preserved_per_question(self):
    by_number = {q['number']: q for q in self.exam['questions']}
    for variant in self.variants():
      for question in variant['questions']:
        source = by_number[question['source_number']]
        self.assertEqual(sorted(o['text'] for o in question['options']),
                         sorted(o['text'] for o in source['options']))
        self.assertEqual(question['stem_text'], source['stem_text'])

  def test_the_correct_answer_follows_the_shuffle(self):
    by_number = {q['number']: q for q in self.exam['questions']}
    for variant in self.variants():
      for question in variant['questions']:
        source = by_number[question['source_number']]
        expected = next(o['text'] for o in source['options']
                        if o['letter'] == source['answer'])
        actual = next(o['text'] for o in question['options']
                      if o['letter'] == question['answer'])
        self.assertEqual(actual, expected,
                         'question %d' % question['source_number'])

  def test_shuffling_actually_happens(self):
    # Not a property of any single variant, but over a handful the original
    # order must not survive everywhere.
    variants = self.variants(count=8, seed=3)
    orders = [tuple(q['source_number'] for q in v['questions'])
              for v in variants]
    self.assertTrue(any(order != tuple(range(1, 11)) for order in orders))
    letters = []
    for variant in variants:
      for question in variant['questions']:
        letters.append(tuple(o['source_letter'] for o in question['options']))
    self.assertTrue(any(item != ('A', 'B', 'C', 'D') for item in letters))

  def test_same_seed_gives_the_same_variants(self):
    first = ev.make_variants(self.exam, 3, seed=42)
    second = ev.make_variants(self.exam, 3, seed=42)
    self.assertEqual(_signature(first), _signature(second))

  def test_different_seeds_give_different_variants(self):
    self.assertNotEqual(_signature(ev.make_variants(self.exam, 3, seed=1)),
                        _signature(ev.make_variants(self.exam, 3, seed=2)))

  def test_variant_does_not_mutate_the_exam(self):
    before = _exam_signature(self.exam)
    ev.make_variants(self.exam, 3, seed=11)
    self.assertEqual(_exam_signature(self.exam), before)


def _exam_signature(exam):
  rows = []
  for question in exam['questions']:
    letters = ''.join(o['letter'] for o in question['options'])
    rows.append((question['number'], letters, question['answer'],
                 tuple(o['text'] for o in question['options'])))
  return rows


def _signature(variants):
  rows = []
  for variant in variants:
    for question in variant['questions']:
      letters = ''.join(o['source_letter'] for o in question['options'])
      rows.append((variant['index'], question['number'],
                   question['source_number'], letters, question['answer']))
  return rows


class LabelRewritingTest(unittest.TestCase):

  def paragraph(self, *runs):
    paragraph = ET.Element(ev.qn('w', 'p'))
    for text in runs:
      run = ET.SubElement(paragraph, ev.qn('w', 'r'))
      node = ET.SubElement(run, ev.qn('w', 't'))
      node.text = text
    return paragraph

  def test_number_split_across_runs(self):
    paragraph = self.paragraph('10', '.  ', 'Find the area')
    self.assertTrue(ev.relabel(paragraph, ev.QUESTION_LABEL, '2'))
    self.assertEqual(ev.paragraph_text(paragraph), '2.  Find the area')

  def test_letter_split_across_runs(self):
    paragraph = self.paragraph('(D', ')  Does not exist')
    self.assertTrue(ev.relabel(paragraph, ev.OPTION_LABEL, 'B'))
    self.assertEqual(ev.paragraph_text(paragraph), '(B)  Does not exist')

  def test_label_inside_a_single_run(self):
    paragraph = self.paragraph('(C)  12')
    self.assertTrue(ev.relabel(paragraph, ev.OPTION_LABEL, 'A'))
    self.assertEqual(ev.paragraph_text(paragraph), '(A)  12')

  def test_only_the_label_run_is_touched(self):
    paragraph = self.paragraph('7', '.  ', 'Evaluate:  ', ' dx')
    ev.relabel(paragraph, ev.QUESTION_LABEL, '3')
    runs = [node.text for node in paragraph.iter(ev.qn('w', 't'))]
    self.assertEqual(runs, ['3', '.  ', 'Evaluate:  ', ' dx'])

  def test_paragraph_without_a_label_is_left_alone(self):
    paragraph = self.paragraph('Choose the best answer.')
    self.assertFalse(ev.relabel(paragraph, ev.QUESTION_LABEL, '1'))
    self.assertEqual(ev.paragraph_text(paragraph), 'Choose the best answer.')


class WrittenDocumentTest(unittest.TestCase):

  @classmethod
  def setUpClass(cls):
    cls.exam = ev.parse_exam(SOURCE)
    cls.variant = ev.make_variants(cls.exam, 1, seed=5)[0]
    cls.tmp = tempfile.TemporaryDirectory()
    cls.student, cls.professor = ev.write_variant(cls.exam, cls.variant,
                                                  cls.tmp.name)

  @classmethod
  def tearDownClass(cls):
    cls.tmp.cleanup()

  def test_files_are_named_per_variant(self):
    self.assertEqual(self.student.name, 'student_1.docx')
    self.assertEqual(self.professor.name, 'professor_1.docx')

  def test_packages_keep_every_source_part(self):
    with zipfile.ZipFile(SOURCE) as source:
      expected = sorted(source.namelist())
    for path in (self.student, self.professor):
      with zipfile.ZipFile(path) as written:
        self.assertEqual(sorted(written.namelist()), expected)
        self.assertIsNone(written.testzip())

  def test_documents_are_well_formed_and_keep_the_w_prefix(self):
    for path in (self.student, self.professor):
      xml_text = ev.read_document_xml(path)
      ET.fromstring(xml_text)
      self.assertTrue(xml_text.startswith('<?xml version="1.0" '
                                          'encoding="UTF-8" standalone="yes"?>'))
      self.assertIn('<w:body>', xml_text)
      self.assertNotIn('ns0:', xml_text)

  def test_no_equation_is_lost(self):
    expected = count_math(document_root(SOURCE))
    self.assertGreater(expected, 0)
    for path in (self.student, self.professor):
      self.assertEqual(count_math(document_root(path)), expected, path.name)

  def test_student_document_has_no_key(self):
    texts = paragraph_texts(document_root(self.student))
    self.assertFalse(any(ev.KEY_HEADING.match(text) for text in texts))
    self.assertEqual(sum(1 for t in texts if ev.KEY_ENTRY.match(t)), 0)

  def test_key_lines_do_not_share_paragraph_ids(self):
    para_id = ev.qn('w14', 'paraId')
    ids = []
    for paragraph in body_paragraphs(document_root(self.professor)):
      value = paragraph.get(para_id)
      if value is not None:
        ids.append(value)
    self.assertEqual(len(ids), len(set(ids)))

  def test_professor_key_matches_the_variant(self):
    texts = paragraph_texts(document_root(self.professor))
    entries = []
    for text in texts:
      match = ev.KEY_ENTRY.match(text)
      if match:
        entries.append((int(match.group(1)), match.group(2)))
    expected = [(q['number'], q['answer']) for q in self.variant['questions']]
    self.assertEqual(entries, expected)
    self.assertTrue(any(ev.KEY_HEADING.match(text) for text in texts))

  def test_questions_are_renumbered_in_order_in_both_documents(self):
    for path in (self.student, self.professor):
      texts = paragraph_texts(document_root(path))
      stop = len(texts)
      for position, text in enumerate(texts):
        if ev.KEY_HEADING.match(text):
          stop = position
          break
      numbers = []
      letters = []
      for text in texts[:stop]:
        question = ev.QUESTION_LABEL.match(text)
        option = ev.OPTION_LABEL.match(text)
        if question:
          numbers.append(int(question.group(1)))
          letters.append([])
        elif option and letters:
          letters[-1].append(option.group(1))
      self.assertEqual(numbers, list(range(1, 11)), path.name)
      for item in letters:
        self.assertEqual(item, ['A', 'B', 'C', 'D'], path.name)

  def test_stem_and_option_text_survives_the_round_trip(self):
    texts = paragraph_texts(document_root(self.student))
    for question in self.variant['questions']:
      stem = next(t for t in texts
                  if t.startswith('%d.' % question['number']))
      self.assertIn(question['stem_text'][:20], stem)
      for option in question['options']:
        if option['text']:
          self.assertTrue(any(t.startswith('(%s)' % option['letter'])
                              and option['text'] in t for t in texts),
                          '%s missing' % option['text'])

  def test_both_documents_carry_the_same_questions(self):
    student = [t for t in paragraph_texts(document_root(self.student))]
    professor = paragraph_texts(document_root(self.professor))
    for text in student:
      if ev.QUESTION_LABEL.match(text) or ev.OPTION_LABEL.match(text):
        self.assertIn(text, professor)


class MarkerRecognitionTest(unittest.TestCase):
  """The label regexes must not assume the Latin alphabet."""

  def test_option_markers_in_any_script(self):
    samples = [
      ('(A)  4', 'A'), ('A)  4', 'A'), ('A.  4', 'A'), ('a)  4', 'a'),
      ('α)  4', 'α'), ('δ)  Does not exist', 'δ'), ('(β)  0', 'β'),
      ('б)  0', 'б'), ('[C]  2', 'C'), ('  (D)  2', 'D'),
      ('i)  4', 'i'), ('(iii)  4', 'iii'), ('iv.  4', 'iv'),
    ]
    for text, expected in samples:
      with self.subTest(text=text):
        match = ev.OPTION_LABEL.match(text)
        self.assertIsNotNone(match, text)
        self.assertEqual(match.group(1), expected)

  def test_prose_is_not_read_as_a_marker(self):
    for text in ('Does not exist', 'Any value', 'Evaluate the limit:',
                 'Note: assume x > 0', 'Choose the best answer.',
                 'Eq. 4 applies here', '12x3 - 4x'):
      with self.subTest(text=text):
        self.assertIsNone(ev.OPTION_LABEL.match(text))

  def test_key_entries_in_any_script(self):
    samples = [('1.  A', ('1', 'A')), ('10)  δ', ('10', 'δ')),
               ('3.  (c)', ('3', 'c')), ('7.  iii', ('7', 'iii'))]
    for text, expected in samples:
      with self.subTest(text=text):
        match = ev.KEY_ENTRY.match(text)
        self.assertIsNotNone(match, text)
        self.assertEqual(match.groups(), expected)

  def test_key_entry_does_not_swallow_a_question(self):
    self.assertIsNone(ev.KEY_ENTRY.match('1.  Evaluate the limit'))

  def test_markers_compare_case_insensitively(self):
    self.assertTrue(ev.same_marker('A', 'a'))
    self.assertTrue(ev.same_marker('α', 'Α'))
    self.assertFalse(ev.same_marker('α', 'a'))
    self.assertFalse(ev.same_marker('A', 'B'))


class FunctionalTest(unittest.TestCase):
  """Full pipeline over both sample tests: .docx in, .docx out.

  Everything here is asserted for each fixture in turn, so the Latin and the
  Greek paper are held to exactly the same standard and nothing may quietly
  depend on the markers being A-D.
  """

  def test_source_documents_are_present(self):
    for fixture in FIXTURES:
      with self.subTest(fixture['name']):
        self.assertTrue(fixture['path'].exists(), fixture['path'])

  def test_parses_both_marker_alphabets(self):
    for fixture in FIXTURES:
      with self.subTest(fixture['name']):
        exam = ev.parse_exam(fixture['path'])
        self.assertEqual(len(exam['questions']), 10)
        found = {}
        for question in exam['questions']:
          markers = [option['letter'] for option in question['options']]
          self.assertEqual(markers, fixture['markers'],
                           'question %d' % question['number'])
          found[question['number']] = question['answer']
        self.assertEqual(found, fixture['key'])

  def test_variants_keep_the_source_markers(self):
    for fixture in FIXTURES:
      with self.subTest(fixture['name']):
        exam = ev.parse_exam(fixture['path'])
        for variant in ev.make_variants(exam, 3, seed=4):
          for question in variant['questions']:
            markers = [option['letter'] for option in question['options']]
            self.assertEqual(markers, fixture['markers'])
            self.assertIn(question['answer'], fixture['markers'])

  def test_student_copy_introduces_nothing_new(self):
    """A student paper is the source test, shuffled -- and nothing else.

    No variant banner, no added heading, no blank line: apart from the
    reordering and the renumbering it must be paragraph-for-paragraph what the
    class would have been given, so nothing on the page says which copy it is.
    """
    for fixture in FIXTURES:
      with self.subTest(fixture['name']):
        exam = ev.parse_exam(fixture['path'])
        expected = Counter()
        for item in exam['preamble_xml']:
          expected[ev.paragraph_text(ET.fromstring(item)).strip()] += 1
        for question in exam['questions']:
          expected[question['stem_text']] += 1
          for option in question['options']:
            expected[option['text']] += 1

        with tempfile.TemporaryDirectory() as folder:
          self.assertEqual(ev.main([str(fixture['path']), '-n', '1',
                                    '--seed', '3', '--out-dir', folder]), 0)
          student = document_root(Path(folder) / 'student_1.docx')
          professor = document_root(Path(folder) / 'professor_1.docx')

        found = Counter()
        for text in paragraph_texts(student):
          found[strip_label(text)] += 1
        self.assertEqual(found, expected)
        # Same count too, so an added blank paragraph cannot slip through.
        self.assertEqual(len(body_paragraphs(student)), sum(expected.values()))

        # The professor copy is that plus exactly the key apparatus:
        # a variant banner, a page break, the heading, one line per question.
        self.assertEqual(len(body_paragraphs(professor)),
                         len(body_paragraphs(student)) + 3
                         + len(exam['questions']))

  def test_end_to_end_documents(self):
    for fixture in FIXTURES:
      with self.subTest(fixture['name']):
        self._check_written_documents(fixture)

  def _check_written_documents(self, fixture):
    exam = ev.parse_exam(fixture['path'])
    by_number = {q['number']: q for q in exam['questions']}
    # Counted from the source rather than hard-coded: these documents get
    # re-exported, and the invariant is "the variants hold every equation the
    # source holds", not "the source holds 31 of them".
    expected_math = count_math(document_root(fixture['path']))
    self.assertGreater(expected_math, 0, fixture['name'])
    expected_shapes = math_signature(document_root(fixture['path']))
    with tempfile.TemporaryDirectory() as folder:
      self.assertEqual(ev.main([str(fixture['path']), '-n', '2', '--seed', '13',
                                '--out-dir', folder]), 0)
      written = sorted(path.name for path in Path(folder).glob('*.docx'))
      self.assertEqual(written, ['professor_1.docx', 'professor_2.docx',
                                 'student_1.docx', 'student_2.docx'])

      # main() reshuffles internally, so rebuild the same variants to compare
      # the papers against: same source, same seed, same count.
      variants = ev.make_variants(exam, 2, seed=13)
      for variant in variants:
        student = Path(folder) / ('student_%d.docx' % variant['index'])
        professor = Path(folder) / ('professor_%d.docx' % variant['index'])

        for path in (student, professor):
          root = document_root(path)
          self.assertEqual(count_math(root), expected_math,
                           '%s lost an equation' % path.name)
          # Not just as many equations -- the same ones, markup for markup.
          self.assertEqual(math_signature(root), expected_shapes,
                           '%s altered an equation' % path.name)
          self._check_question_layout(root, fixture, path.name)

        student_texts = paragraph_texts(document_root(student))
        self.assertFalse(any(ev.KEY_HEADING.match(t) for t in student_texts))
        self.assertFalse(any(ev.KEY_ENTRY.match(t) for t in student_texts))

        professor_texts = paragraph_texts(document_root(professor))
        self.assertTrue(any(ev.KEY_HEADING.match(t) for t in professor_texts))
        entries = []
        for text in professor_texts:
          match = ev.KEY_ENTRY.match(text)
          if match:
            entries.append((int(match.group(1)), match.group(2)))
        self.assertEqual(entries,
                         [(q['number'], q['answer'])
                          for q in variant['questions']])

        # The printed key must point at the same answer *content* as the
        # source key did, not merely at some marker.
        for question in variant['questions']:
          source = by_number[question['source_number']]
          expected = next(o['text'] for o in source['options']
                          if ev.same_marker(o['letter'], source['answer']))
          actual = next(o['text'] for o in question['options']
                        if ev.same_marker(o['letter'], question['answer']))
          self.assertEqual(actual, expected,
                           'variant %d question %d'
                           % (variant['index'], question['number']))

  def _check_question_layout(self, root, fixture, name):
    numbers = []
    markers = []
    for text in paragraph_texts(root):
      if ev.KEY_HEADING.match(text):
        break
      question = ev.QUESTION_LABEL.match(text)
      option = ev.OPTION_LABEL.match(text)
      if question:
        numbers.append(int(question.group(1)))
        markers.append([])
      elif option and markers:
        markers[-1].append(option.group(1))
    self.assertEqual(numbers, list(range(1, 11)), name)
    for item in markers:
      self.assertEqual(item, fixture['markers'], name)


class OverwriteTest(unittest.TestCase):
  """Existing papers are never replaced without the user saying so."""

  def setUp(self):
    self.folder = tempfile.TemporaryDirectory()
    self.addCleanup(self.folder.cleanup)
    self.dir = Path(self.folder.name)
    self.asked = []

  def answer(self, reply):
    def ask(prompt):
      self.asked.append(prompt)
      return reply
    return ask

  def existing(self, *names):
    paths = []
    for name in names:
      path = self.dir / name
      path.write_bytes(b'old')
      paths.append(path)
    return paths

  def quiet(self):
    return open(os.devnull, 'w')

  def test_variant_paths_lists_both_documents_per_variant(self):
    names = [path.name for path in ev.variant_paths(self.dir, 2)]
    self.assertEqual(names, ['student_1.docx', 'professor_1.docx',
                             'student_2.docx', 'professor_2.docx'])

  def test_no_question_when_nothing_exists(self):
    ok = ev.confirm_overwrite(ev.variant_paths(self.dir, 3),
                              ask=self.answer('n'))
    self.assertTrue(ok)
    self.assertEqual(self.asked, [])

  def test_no_question_with_force(self):
    self.existing('student_1.docx')
    with self.quiet() as devnull:
      ok = ev.confirm_overwrite(ev.variant_paths(self.dir, 1), force=True,
                                ask=self.answer('n'), stream=devnull)
    self.assertTrue(ok)
    self.assertEqual(self.asked, [])

  def test_asks_once_for_the_whole_run(self):
    self.existing('student_1.docx', 'professor_2.docx')
    with self.quiet() as devnull:
      ok = ev.confirm_overwrite(ev.variant_paths(self.dir, 2),
                                ask=self.answer('y'), stream=devnull)
    self.assertTrue(ok)
    self.assertEqual(len(self.asked), 1)

  def test_accepted_answers(self):
    self.existing('student_1.docx')
    for reply, expected in (('y', True), ('Y', True), ('yes', True),
                            (' YES ', True), ('n', False), ('', False),
                            ('no', False), ('later', False)):
      with self.subTest(reply=reply):
        with self.quiet() as devnull:
          self.assertEqual(
            ev.confirm_overwrite(ev.variant_paths(self.dir, 1),
                                 ask=self.answer(reply), stream=devnull),
            expected)

  def test_no_input_at_all_means_no(self):
    self.existing('student_1.docx')

    def ask(prompt):
      raise EOFError

    with self.quiet() as devnull:
      self.assertFalse(ev.confirm_overwrite(ev.variant_paths(self.dir, 1),
                                            ask=ask, stream=devnull))

  def test_interrupt_means_no(self):
    self.existing('student_1.docx')

    def ask(prompt):
      raise KeyboardInterrupt

    with self.quiet() as devnull:
      self.assertFalse(ev.confirm_overwrite(ev.variant_paths(self.dir, 1),
                                            ask=ask, stream=devnull))

  def test_declining_leaves_every_file_untouched(self):
    self.assertEqual(ev.main([str(SOURCE), '-n', '2', '--seed', '1',
                              '--out-dir', str(self.dir)]), 0)
    before = {}
    for path in self.dir.glob('*.docx'):
      before[path.name] = path.read_bytes()
    self.assertEqual(len(before), 4)

    code = ev.main([str(SOURCE), '-n', '2', '--seed', '99',
                    '--out-dir', str(self.dir)], ask=self.answer('n'))
    self.assertEqual(code, 3)
    self.assertEqual(len(self.asked), 1)
    after = {}
    for path in self.dir.glob('*.docx'):
      after[path.name] = path.read_bytes()
    self.assertEqual(after, before)

  def test_agreeing_replaces_them(self):
    self.assertEqual(ev.main([str(SOURCE), '-n', '1', '--seed', '1',
                              '--out-dir', str(self.dir)]), 0)
    before = (self.dir / 'student_1.docx').read_bytes()
    code = ev.main([str(SOURCE), '-n', '1', '--seed', '99',
                    '--out-dir', str(self.dir)], ask=self.answer('y'))
    self.assertEqual(code, 0)
    self.assertNotEqual((self.dir / 'student_1.docx').read_bytes(), before)

  def test_force_replaces_without_asking(self):
    self.assertEqual(ev.main([str(SOURCE), '-n', '1', '--seed', '1',
                              '--out-dir', str(self.dir)]), 0)
    code = ev.main([str(SOURCE), '-n', '1', '--seed', '99', '--force',
                    '--out-dir', str(self.dir)], ask=self.answer('n'))
    self.assertEqual(code, 0)
    self.assertEqual(self.asked, [])

  def test_a_partial_collision_still_asks(self):
    # Only one of the four targets exists: still a question, never a
    # silently half-replaced folder.
    self.existing('professor_2.docx')
    with self.quiet() as devnull:
      ok = ev.confirm_overwrite(ev.variant_paths(self.dir, 2),
                                ask=self.answer('n'), stream=devnull)
    self.assertFalse(ok)
    self.assertEqual(len(self.asked), 1)


class CommandLineTest(unittest.TestCase):

  def test_end_to_end_writes_two_files_per_variant(self):
    with tempfile.TemporaryDirectory() as folder:
      code = ev.main([str(SOURCE), '-n', '4', '--seed', '9',
                      '--out-dir', folder])
      self.assertEqual(code, 0)
      written = sorted(p.name for p in Path(folder).glob('*.docx'))
      expected = []
      for index in range(1, 5):
        expected.append('professor_%d.docx' % index)
        expected.append('student_%d.docx' % index)
      self.assertEqual(written, sorted(expected))

  def test_missing_file_exits_non_zero(self):
    self.assertEqual(ev.main(['no_such_file.docx']), 1)

  def test_zero_variants_is_rejected(self):
    self.assertEqual(ev.main([str(SOURCE), '-n', '0']), 2)


if __name__ == '__main__':
  unittest.main(verbosity=2)
