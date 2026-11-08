"""Fixture-limited Office compatibility check; execute only on a disposable runner."""
import hashlib
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(exe, profile, source, extension, destination, expected=True):
    command = [str(exe), f'-env:UserInstallation={profile.as_uri()}', '--headless',
               '--convert-to', extension, '--outdir', str(destination), str(source)]
    result = subprocess.run(command, capture_output=True, text=True, timeout=180)
    output = destination / (source.stem + '.' + extension)
    print('command:', command, 'exit:', result.returncode, 'stdout:', result.stdout,
          'stderr:', result.stderr, flush=True)
    if expected:
        assert result.returncode == 0 and output.is_file() and output.stat().st_size > 0
        print('output sha256:', digest(output), flush=True)
        return output
    assert not output.exists(), f'unexpected output: {output}'


def main():
    exe = Path(sys.argv[1]).resolve(strict=True)
    subprocess.run([str(exe), '--version'], check=True, timeout=30)
    with tempfile.TemporaryDirectory(prefix='cfh-office-') as root:
        root = Path(root)
        inputs, middle, outputs = [root / p for p in ('inputs', 'middle', 'outputs')]
        for directory in (inputs, middle, outputs):
            directory.mkdir()
        profile = root / 'profile'
        doc = Document()
        paragraph = doc.add_paragraph('CFH Writer roundtrip ')
        paragraph.add_run('bold sentinel').bold = True
        doc.save(inputs / 'writer.docx')
        book = Workbook()
        sheet = book.active
        sheet.title = 'Data'
        sheet['A1'], sheet['A2'], sheet['A3'] = 21, 2, '=A1*A2'
        book.save(inputs / 'calc.xlsx')
        slides = Presentation()
        for title in ('CFH slide one', 'CFH slide two'):
            slides.slides.add_slide(slides.slide_layouts[0]).shapes.title.text = title
        slides.save(inputs / 'impress.pptx')
        original = {p: digest(p) for p in inputs.iterdir()}
        print('input sha256:', {p.name: h for p, h in original.items()}, flush=True)
        for name, inter, final in (('writer.docx', 'odt', 'docx'),
                                   ('calc.xlsx', 'ods', 'xlsx'),
                                   ('impress.pptx', 'odp', 'pptx')):
            first = run(exe, profile, inputs / name, inter, middle)
            run(exe, profile, first, final, outputs)
        paragraphs = Document(outputs / 'writer.docx').paragraphs
        assert any('CFH Writer roundtrip' in p.text for p in paragraphs)
        assert any('bold sentinel' in r.text and r.bold for p in paragraphs for r in p.runs)
        sheet = load_workbook(outputs / 'calc.xlsx', data_only=False)['Data']
        assert (sheet['A1'].value, sheet['A2'].value, sheet['A3'].value) == (21, 2, '=A1*A2')
        slides = Presentation(outputs / 'impress.pptx')
        assert len(slides.slides) == 2
        for slide, title in zip(slides.slides, ('CFH slide one', 'CFH slide two')):
            assert any(title in shape.text for shape in slide.shapes if shape.has_text_frame)
        assert all(digest(p) == h for p, h in original.items())
        bad = inputs / 'malformed.docx'
        bad.write_bytes(b'not an office document')
        run(exe, profile, bad, 'odt', middle, expected=False)
        # A regular file in place of a directory is reliably unusable, even for admins.
        blocked = root / 'not-a-directory'
        blocked.write_text('occupied', encoding='utf-8')
        run(exe, profile, inputs / 'writer.docx', 'odt', blocked, expected=False)
        print('FIXTURE-LIMITED OFFICE ROUNDTRIP PASSED', flush=True)


if __name__ == '__main__':
    main()
