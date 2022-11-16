#!/usr/bin/env python3.8

"""

https://docs.python.org/ja/3.9/library/dialog.html#module-tkinter.simpledialog

class tkinter.simpledialog.Dialog(parent, title=None)
  The base class for custom dialogs.

  body(master)
    Override to construct the dialog's interface and return the widget that should have initial focus.

  buttonbox()
    Default behaviour adds OK and Cancel buttons. Override for custom button layouts.

"""

import tkinter as tk
from tkinter import ttk
from tkinter import Entry
from tkinter.simpledialog import Dialog
from functools import partial

#YES/NOダイアログ
class AskYesNoDialog(Dialog):

    MyFrame = partial(ttk.Frame, style="MyDialog.TFrame")
    MyLabel = partial(ttk.Label, style="MyDialog.TLabel")
    MyButton = partial(ttk.Button, style="MyDialog.TButton")

    def __init__(self, parent=None, title="", message=""):

        if not parent:
            parent = tk._default_root

        self._message = message

        Dialog.__init__(self, parent, title)


    def body(self, master):
        self.MyLabel(master, text=self._message).pack()

    def buttonbox(self):
        frame = self.MyFrame(self)
        frame.pack()

        w = self.MyButton(frame, text="OK", command=self.ok, default=tk.ACTIVE)
        w.pack(side=tk.LEFT, padx=5, pady=5)

        w = self.MyButton(frame, text="Cancel", command=self.cancel)
        w.pack(side=tk.LEFT, padx=5, pady=5)

    def validate(self):
        # XXX:
        self.result = True
        return 1


# see also tkinter.messagebox.askyesno
def askyesno(title, message=None, parent=None):
    dialog = AskYesNoDialog(parent, title, message)
    return dialog.result

#テキスト入力ダイアログ
class TextInputDialog(Dialog):

    MyFrame = partial(ttk.Frame, style="MyDialog.TFrame")
    MyLabel = partial(ttk.Label, style="MyDialog.TLabel")
    MyButton = partial(ttk.Button, style="MyDialog.TButton")
    MyTextBox = partial(ttk.Entry, style="MyDialog.TEntry")

    def __init__(self, parent=None, title="", message=""):

        if not parent:
            parent = tk._default_root

        self._message = message
        
        #固定設定設定
        font = ("", 14)
        style = ttk.Style()
        style.configure("MyDialog.TLabel", font=font)
        style.configure("MyDialog.TButton", font=font)
        
        Dialog.__init__(self, parent, title)

    def body(self, master):
        self.MyLabel(master, text=self._message).pack()
        self.textString = tk.StringVar()
        self.MyTextBox(master,textvariable=self.textString).pack()

    def buttonbox(self):
        frame = self.MyFrame(self)
        frame.pack()

        w = self.MyButton(frame, text="送信", command=self.ok, default=tk.ACTIVE)
        w.pack(side=tk.LEFT, padx=5, pady=5)

    def validate(self):
        # XXX:
        self.result = True
        return 1
    
    def getText(self):
        return self.textString.get()

# see also tkinter.messagebox.askyesno
def askText(title, message=None, parent=None):
    dialog = TextInputDialog(parent, title, message)
    return dialog.getText()

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    #dialog = TextInputDialog(root, "不明なカテゴリ", "〇〇の正しいカテゴリを「SSM・GM・DC・HF」から選んで入力してください。")
    result = askText("不明なカテゴリ", "〇〇の正しいカテゴリを「SSM・GM・DC・HF」から選んで入力してください。", parent=root)
    #result = dialog.getText()
    if result == "":
        print("RESULT>", False)
    else:
        print("RESULT>", result)