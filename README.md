WordAutomationSample
====================

Word 側にブックマークを指定しておくと、それをたよりにデータを差し込むことができる。

```vb
        Dim word As Word.Application = CreateObject("Word.Application")
        word.Visible = True

        Dim document As Word.Document = word.Documents.Add(String.Format("{0}\template.dotx", path))
        document.Bookmarks.Item("title").Range.Text = "タイトル"
        document.Bookmarks.Item("caption").Range.Text = "キャプション"

```

RTF を別途作っておくと、これを Word 文書の途中に差し込むこともできる。

```vb
            table.Cell(r, 2).Range.InsertFile(FileName:=String.Format("{0}\{1}.rtf", path, r - 1),
                                              Range:="",ConfirmConversions:=False,
                                              Link:=False,Attachment:=False)

```

こんな感じになる。
![](http://cdn-ak.f.st-hatena.com/images/fotolife/d/dechnostick/20140818/20140818014458.png)
