# Word(docx)の自動作成

- [Word(docx)の自動生成](#worddocx%E3%81%AE%E8%87%AA%E5%8B%95%E7%94%9F%E6%88%90)
  - [はじめに](#%E3%81%AF%E3%81%98%E3%82%81%E3%81%AB)
    - [目的](#%E7%9B%AE%E7%9A%84)
    - [関連する記事](#%E9%96%A2%E9%80%A3%E3%81%99%E3%82%8B%E8%A8%98%E4%BA%8B)
    - [実行環境](#%E5%AE%9F%E8%A1%8C%E7%92%B0%E5%A2%83)
    - [ソースコード](#%E3%82%BD%E3%83%BC%E3%82%B9%E3%82%B3%E3%83%BC%E3%83%89)
  - [UIの実装](#ui%E3%81%AE%E5%AE%9F%E8%A3%85)
  - [docx作成の実装](#docx%E4%BD%9C%E6%88%90%E3%81%AE%E5%AE%9F%E8%A3%85)

## はじめに

Mac環境の記事ですが、Windows環境も同じ手順になります。環境依存の部分は読み替えてお試しください。

### 目的

この記事を最後まで読むと、次のことができるようになります。

- デスクトップアプリを実装する
- Word(docx)の自動作成を実装する

`アプリ`

JSONデータとWord(docx)テンプレートをレンダリングして出力する。複数件のデータはページに分けて出力する。

<img width="400" alt="スクリーンショット 2019-06-07 14.16.32.png" src="https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/326996/3be3366a-095d-4971-7358-2976dc8e0874.png">

`JSONデータ`

```sample_data.json
[
    {
        "create_date": "令和元年 5月 1日",
        "to_company_name": "カレンダー株式会社",
        "to_company_department": "イヤホン本部",
        "relocation_date": "令和元年 7月 28日",
        "post_code": "123-4567",
        "new_address": "東京都港区サンプル1-2-3 ビルディング 45F",
        "new_phone_number": "1234-56-7890",
        "new_fax_number": "1234-56-7890"
    },
    {
        "create_date": "2019年 5月 1日",
        "to_company_name": "冷暖房リモコン株式会社",
        "to_company_department": "ティッシュケース本部",
        "relocation_date": "2019年 7月 28日",
        "post_code": "987-6543",
        "new_address": "東京都港区サンプル9-8-7 ビルディング 65F",
        "new_phone_number": "0987-65-4321",
        "new_fax_number": "0987-65-4321"
    }
]
```

`Word(docx)テンプレート`

<img width="500" alt="スクリーンショット 2019-06-07 17.12.37.png" src="https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/326996/67720e91-ac45-aa4d-fe2a-f8d324748c71.png">

`実行結果`

page 1/2

<img width="500" alt="スクリーンショット 2019-06-07 17.26.42.png" src="https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/326996/06feaa58-41ca-52dc-058f-9cd186468ef4.png">

page 2/2

<img width="500" alt="スクリーンショット 2019-06-07 17.26.53.png" src="https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/326996/1b38aead-0584-c7b1-98c9-8c2b32beb0c4.png">

### 関連する記事

- [tkinter - Python interface to Tcl/Tk](https://docs.python.org/3/library/tkinter.html)
- [docxtpl - PyPI](https://pypi.org/project/docxtpl/)

### 実行環境

| 環境         | Ver.    |
| ------------ | ------- |
| macOS Mojave | 10.14.5 |
| Python       | 3.7.3   |
| tkinter      | 8.5     |
| docxtpl      | 0.6.1   |

### ソースコード

実際に実装内容やソースコードを追いながら読むとより理解が深まるかと思います。是非ご活用ください。

[GitHub](https://github.com/nsuhara/python-CreateDocx.git)

## UIの実装

UIは**Tkinter**で実装します。

Tkinterとは、Windows、MacOS、Linuxに対応するクロスプラットフォームなGUIライブラリです。

```app.py
import os
import tkinter as tk
from tkinter import filedialog as fdialog
from tkinter import messagebox as mdialog

from model import Docx


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.create_widgets()

    def set_title(self):
        self.master.title('Create Docx')

    def set_menu_bar(self):
        self.menu_bar = tk.Menu(self.master)
        self.master.config(menu=self.menu_bar)
        file_menu = tk.Menu(self.menu_bar)
        file_menu.add_command(label='Exit', command=self.master.quit)
        self.menu_bar.add_cascade(label='File', menu=file_menu)

    def select_file(self, entry):
        entry.delete(0, tk.END)
        entry.insert(0, fdialog.askopenfilename(initialdir=os.getcwd()))

    def create_docx(self, json_url, template_url):
        if not os.path.exists(json_url) or not os.path.exists(template_url):
            mdialog.showerror('Error', 'Please select JSON and Template.')
            return

        docx = Docx(json_url=json_url, template_url=template_url)
        docx.render()

    def set_body(self):
        tk.Label(self.master, text='JSON:').grid(row=0, column=0)
        entry_json = tk.Entry(self.master)
        entry_json.grid(row=0, column=1, pady=5)
        tk.Button(self.master, text='Select...',
                  command=lambda: self.select_file(entry_json)).grid(row=0, column=2)

        tk.Label(self.master, text='Template:').grid(row=1, column=0)
        entry_template = tk.Entry(self.master)
        entry_template.grid(row=1, column=1, pady=5)
        tk.Button(self.master, text='Select...',
                  command=lambda: self.select_file(entry_template)).grid(row=1, column=2)

        tk.Button(self.master, text='Create', width=30,
                  command=lambda: self.create_docx(entry_json.get(), entry_template.get())).grid(row=2, column=0, columnspan=3)

    def create_widgets(self):
        self.master.geometry()
        self.entry = tk.Entry(self.master)

        self.set_title()
        self.set_menu_bar()
        self.set_body()


# fix tkinter bug start
def fix_bug():
    width_height = root.winfo_geometry().split('+')[0].split('x')
    width = int(width_height[0])
    height = int(width_height[1])
    root.geometry('{}x{}'.format(width+1, height+1))
# fix tkinter bug end


if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    # fix tkinter bug start
    root.update()
    root.after(0, fix_bug)
    # fix tkinter bug end
    app.mainloop()
```

`※ tkinterで白画面になる不具合がありサイズを1プラスすることで改修`

## docx作成の実装

docxの作成は**docxtpl**で実装します。

docxtplとは、JSONデータとWord(docx)テンプレートをレンダリングするライブラリです。

```model.py
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
```
