import cgi
import json
import os.path
import re
import sys
import uuid
from tkinter import messagebox as mdialog

from docxtpl import DocxTemplate


_illegal_unichrs = [(0x00, 0x08), (0x0B, 0x0C), (0x0E, 0x1F),
                    (0x7F, 0x84), (0x86, 0x9F),
                    (0xFDD0, 0xFDDF), (0xFFFE, 0xFFFF)]
if sys.maxunicode >= 0x10000:  # not narrow build
    _illegal_unichrs.extend([(0x1FFFE, 0x1FFFF), (0x2FFFE, 0x2FFFF),
                             (0x3FFFE, 0x3FFFF), (0x4FFFE, 0x4FFFF),
                             (0x5FFFE, 0x5FFFF), (0x6FFFE, 0x6FFFF),
                             (0x7FFFE, 0x7FFFF), (0x8FFFE, 0x8FFFF),
                             (0x9FFFE, 0x9FFFF), (0xAFFFE, 0xAFFFF),
                             (0xBFFFE, 0xBFFFF), (0xCFFFE, 0xCFFFF),
                             (0xDFFFE, 0xDFFFF), (0xEFFFE, 0xEFFFF),
                             (0xFFFFE, 0xFFFFF), (0x10FFFE, 0x10FFFF)])
_illegal_ranges = ['%s-%s' % (chr(low), chr(high))
                   for (low, high) in _illegal_unichrs]
_illegal_xml_chars_RE = re.compile(u'[%s]' % u''.join(_illegal_ranges))


class Docx(object):
    def __init__(self, json_url, template_url):
        self.json_url = json_url
        self.template_url = template_url

    def read_data(self):
        with open(self.json_url, 'r') as f:
            load_data = json.load(f)

        json_data = json.dumps(load_data)
        json_data = cgi.escape(json_data)
        json_data = json_data.replace('\n', '\\n')
        dict_data = json.loads(json_data)

        for d in dict_data:
            for k in d.keys():
                try:
                    d[k] = _illegal_xml_chars_RE.sub('', d[k])
                except TypeError:
                    pass

        return dict_data

    def render(self):
        dict_data = self.read_data()

        docx = DocxTemplate(self.template_url)
        docx.render({'applications': dict_data})

        file_name = '{}.{}'.format(str(uuid.uuid4()), 'docx')

        save_dir = os.path.join(os.path.curdir, 'output')
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        docx.save(os.path.join(save_dir, file_name))

        mdialog.showinfo('Successful', 'Please check the output folder.')
