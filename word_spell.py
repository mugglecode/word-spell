import shutil
import tempfile
import zipfile
import re
import os
from zipfile import ZIP_DEFLATED

class Document():
    def __init__(self, path: str):
        self.path = path
        self.xml = self._get_word_xml()
        if isinstance(self.xml, bytes):
            self.xml = self._get_word_xml().decode()


    def _get_word_xml(self):
        with open(self.path, 'rb') as f:
            zip = zipfile.ZipFile(f, compression=ZIP_DEFLATED)
            xml_content = zip.read('word/document.xml')
        return xml_content

    def save(self, out: str, xml: str):
        tmp_dir = tempfile.mkdtemp()
        f = open(self.path, 'rb')
        zip = zipfile.ZipFile(f, compression=ZIP_DEFLATED)
        zip.extractall(tmp_dir)

        with open(os.path.join(tmp_dir, 'word/document.xml'), 'w') as f:
            f.write(xml)

        filenames = zip.namelist()

        with zipfile.ZipFile(out, "w") as docx:
            for filename in filenames:
                docx.write(os.path.join(tmp_dir, filename), filename)

        shutil.rmtree(tmp_dir)

    def get_vars(self):
        exp = re.compile(
            r'<w:t[^>]*>([^<>]*)</w:t>')
        matches = []
        for m in exp.finditer(self.xml):
            matches.append(m)

        result = []
        current = ''
        opening_found = 0
        closing_found = 0
        full_pattern = re.compile(r'{{([\S\s]*)}}')
        for m in matches:
            content = m.group(1)
            full_match = full_pattern.search(m.group(1))
            if full_match:
                result.append(full_match.group(1))
                continue

            if content.startswith('{{'):
                opening_found = 2
                if content.endswith('}'):
                    closing_found = 1
                    current = content.replace('{{', '').replace('}', '')
                else:
                    current = content.replace('{{', '')
                continue

            if '{{' in content:
                pos = content.find('{{')
                opening_found = 2
                current = content[pos+2]
                continue
            elif content.endswith('{'):
                opening_found += 1
                continue

            if '}}' in content:
                pos = content.find('}}')
                closing_found = 0
                opening_found = 0
                current += content[:pos-2]
                result.append(current)
                current = ''
            elif content.endswith('}'):
                closing_found += 1

            if content == '{':
                opening_found += 1
                if opening_found == 2:
                    continue
            elif content == '{{':
                opening_found += 2
                continue

            if opening_found == 2:
                if content != '}' and content != '}}':
                    current += content
                elif content == '}':
                    closing_found += 1
                elif content == '}}':
                    closing_found += 2

            if closing_found == 2:
                result.append(current)
                current = ''
                opening_found = 0
                closing_found = 0

        return result

    def debug_args(self, **kwargs):
        vars = self.get_vars()
        for v in kwargs:
            try:
                vars.remove(v)
            except:
                print(v)
        print(vars)

    def render_from_template(self, out: str, **kwargs):
        vars = self.get_vars()
        for v in vars:
            if v not in kwargs:
                raise ValueError(f'expected {len(vars)} arguments got {kwargs.__len__()}, check typo')
        xml = self.xml
        for v in vars:
            r = r'{(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?{[ ]*(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?' + v +'(<[A-Za-z<>=/0-9- \":\u4e00-\u9fa5]*>)?[ ]*}(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?}'
            exp = re.compile(r)
            xml = exp.sub(str(kwargs[v]), xml)
        self.save(out, xml)
