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

    def save(self, out: str):
        tmp_dir = tempfile.mkdtemp()
        f = open(self.path, 'rb')
        zip = zipfile.ZipFile(f, compression=ZIP_DEFLATED)
        zip.extractall(tmp_dir)

        with open(os.path.join(tmp_dir, 'word/document.xml'), 'w') as f:
            f.write(self.xml)

        filenames = zip.namelist()

        with zipfile.ZipFile(out, "w") as docx:
            for filename in filenames:
                docx.write(os.path.join(tmp_dir, filename), filename)

        shutil.rmtree(tmp_dir)

    def get_vars(self):
        exp = re.compile(
            r'{(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?{[ ]*(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?([A-Za-z0-9_\u4e00-\u9fa5]*)(<[A-Za-z<>=/0-9- \":\u4e00-\u9fa5]*>)?[ ]*}(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?}')
        vars = exp.findall(self.xml)
        result = []
        item: tuple
        for item in vars:
            if item.__len__() == 1:
                result.append(item[0])
            else:
                for m in item:
                    if '<' in m or '>' in m:
                        continue
                    elif m != '':
                        result.append(m)
                        break
        return result

    def debug_args(self, **kwargs):
        vars = self.get_vars()
        for v in kwargs:
            vars.remove(v)
        return vars

    def render_from_template(self, out: str, **kwargs):
        vars = self.get_vars()
        for v in vars:
            if v not in kwargs:
                raise ValueError(f'expected {len(vars)} arguments got {kwargs.__len__()}, check typo')
        for v in vars:
            r = r'{(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?{[ ]*(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?' + v +'(<[A-Za-z<>=/0-9- \":\u4e00-\u9fa5]*>)?[ ]*}(<[A-Za-z\" -=0-9<>/:\u4e00-\u9fa5]*>)?}'
            exp = re.compile(r)
            self.xml = exp.sub(str(kwargs[v]), self.xml)
        self.save(out)
