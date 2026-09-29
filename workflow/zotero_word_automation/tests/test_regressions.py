"""Focused offline regressions; no Word or Zotero library access."""
import contextlib
import csv
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
from xml.sax.saxutils import escape
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import build_payload as build
import validate_docx_zotero as check

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def document(body):
    return f'<w:document xmlns:w="{W}"><w:body><w:p>{body}</w:p></w:body></w:document>'


def field(code, display):
    midpoint = len(code) // 2
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText>{escape(code[:midpoint])}</w:instrText></w:r>'
            f'<w:r><w:instrText>{escape(code[midpoint:])}</w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:t>{escape(display)}</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def citation(title='Hidden title'):
    return 'ADDIN ZOTERO_ITEM CSL_CITATION ' + json.dumps({
        'citationID': 'test', 'citationItems': [{'uris': ['http://zotero.org/users/local/X/items/ABC12345'],
                                              'itemData': {'title': title}}]})


class Regressions(unittest.TestCase):
    def run_validator(self, xml, *args):
        with tempfile.TemporaryDirectory() as folder:
            doc = Path(folder) / 'test.docx'
            with zipfile.ZipFile(doc, 'w') as archive:
                archive.writestr('word/document.xml', xml)
            with patch.object(sys, 'argv', ['validate', '--docx', str(doc), *args]), contextlib.redirect_stdout(io.StringIO()) as output:
                status = check.main()
            return status, json.loads(output.getvalue())

    def test_hidden_metadata_does_not_count_as_visible_text(self):
        status, result = self.run_validator(document(field(citation(), '1')), '--must-contain', 'Hidden title')
        self.assertEqual(status, 2)
        self.assertEqual(result['zotero_item_field_count'], 1)

    def test_visible_reference_and_split_codes(self):
        xml = document(field(citation(), '1') + field('ADDIN ZOTERO_BIBL {} CSL_BIBLIOGRAPHY', 'A & B Visible title'))
        status, result = self.run_validator(xml, '--bibliography-contains', 'A & B Visible title', '--expected-bibs', '1')
        self.assertEqual(status, 0)
        self.assertEqual(result['refresh_completion'], 'not_determined_by_static_check')

    def test_body_title_does_not_prove_reference_exists(self):
        xml = document('<w:r><w:t>Expected title</w:t></w:r>' + field('ZOTERO_BIBL {}', 'Other title'))
        self.assertEqual(self.run_validator(xml, '--bibliography-contains', 'Expected title')[0], 2)

    def test_placeholders_rejected(self):
        for text in ('[TG01.original]', '[CL01_G1]', '[Bibliography placeholder]'):
            with self.subTest(text=text):
                self.assertEqual(self.run_validator(document(field(citation(), text)))[0], 2)

    def test_deleted_and_hidden_runs_excluded(self):
        xml = document('<w:del><w:r><w:t>deleted</w:t></w:r></w:del><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r><w:r><w:t>visible</w:t></w:r>')
        self.assertEqual(check.xml_to_text(xml), 'visible')

    def test_simple_field_and_broken_field(self):
        xml = document('<w:fldSimple w:instr="ZOTERO_BIBL {}"><w:r><w:t>Reference</w:t></w:r></w:fldSimple>')
        self.assertEqual(len(check.inspect_xml(xml)[1]), 1)
        self.assertTrue(check.inspect_xml(document('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'))[2])

    def test_duplicate_citation_id_rejected(self):
        self.assertEqual(self.run_validator(document(field(citation(), '1') * 2))[0], 2)

    def test_identity_normalization_and_conflicts(self):
        row = {'citation_id': 'EXAMPLE', 'zotero_key': 'ABC12345', 'title': 'MoS2: A & B', 'doi': 'https://doi.org/10.1234/Abc'}
        item = {'id': 'http://zotero.org/users/local/X/items/ABC12345', 'title': 'MoS<sub>2</sub> A &amp; B', 'DOI': '10.1234/abc'}
        build.verify_identity(row, item)
        for key, value in [('title', 'Different paper'), ('DOI', '10.9999/wrong'), ('id', 'http://zotero.org/users/local/X/items/OTHER')]:
            with self.subTest(key=key), self.assertRaises(SystemExit):
                build.verify_identity(row, {**item, key: value})

    def test_style_required_before_api(self):
        with patch.object(sys, 'argv', ['build', '--docx', 'x', '--task-table', 'x', '--crosswalk', 'x', '--output', 'x']), contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            build.main()

    def run_builder(self, **changes):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            doc = folder / 'fixture.docx'
            doc.write_bytes(b'fixture; no Word invocation')
            row = {'insert_as_group': 'G1', 'anchor_text': 'Exact text', 'position': 'after',
                   'citation_id': 'ID1', 'citation_order': '1', 'status': 'main',
                   'default_insert': 'yes', 'action': 'add', **changes}
            cw = {'citation_id': 'ID1', 'title': 'Test title', 'doi': '10.1234/test',
                  'zotero_key': 'ABC12345', 'zotero_item_id': '1'}
            for name, data in [('tasks.csv', row), ('crosswalk.csv', cw)]:
                with (folder / name).open('w', newline='', encoding='utf-8') as stream:
                    writer = csv.DictWriter(stream, fieldnames=data.keys())
                    writer.writeheader()
                    writer.writerow(data)
            argv = ['build', '--docx', str(doc), '--task-table', str(folder / 'tasks.csv'),
                    '--crosswalk', str(folder / 'crosswalk.csv'), '--output', str(folder / 'payload.json'),
                    '--style', 'http://www.zotero.org/styles/nano-letters']
            item = {'id': 'http://zotero.org/users/local/X/items/ABC12345', 'title': 'Test title', 'DOI': '10.1234/test'}
            with patch.object(sys, 'argv', argv), patch.object(build, 'fetch_csl_json', return_value=item), contextlib.redirect_stdout(io.StringIO()):
                build.main()
            return json.loads((folder / 'payload.json').read_text())

    def test_payload_has_fingerprint_and_unique_default(self):
        payload = self.run_builder()
        self.assertEqual(len(payload['doc_sha256']), 64)
        self.assertEqual(payload['groups'][0]['anchor_occurrence'], 'unique')

    def test_incremental_actions_rejected_by_add_only_builder(self):
        for action in ('keep', 'replace', 'delete'):
            with self.subTest(action=action), self.assertRaises(SystemExit):
                self.run_builder(action=action)

    def test_invalid_order_and_position_rejected(self):
        for change in ({'citation_order': 'unknown'}, {'position': 'middle'}, {'anchor_occurrence': 'maybe'}):
            with self.subTest(change=change), self.assertRaises(SystemExit):
                self.run_builder(**change)


if __name__ == '__main__':
    unittest.main(verbosity=2)
