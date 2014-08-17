Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Word
Imports System.Reflection
Imports System.IO

Class Class1

    Public Shared Sub Main(args() As String)

        Dim path = Assembly.GetExecutingAssembly().Location
        path = Directory.GetParent(path).FullName
        path = Directory.GetParent(path).FullName
        path = Directory.GetParent(path).FullName

        Dim word As Word.Application = CreateObject("Word.Application")
        word.Visible = True

        Dim document As Word.Document = word.Documents.Add(String.Format("{0}\template.dotx", path))
        document.Bookmarks.Item("title").Range.Text = "タイトル"
        document.Bookmarks.Item("caption").Range.Text = "キャプション"

        Dim table As Word.Table =
            document.Tables.Add(Range:=document.Bookmarks.Item("list").Range,
                                NumRows:=4, NumColumns:=2,
                                DefaultTableBehavior:=WdDefaultTableBehavior.wdWord9TableBehavior,
                                AutoFitBehavior:=WdAutoFitBehavior.wdAutoFitContent)

        table.Cell(1, 2).Range.Text = "天気"
        For r = 2 To 4
            table.Cell(r, 1).Range.Text = String.Format("{0}日目", r - 1)
            table.Cell(r, 2).Range.InsertFile(FileName:=String.Format("{0}\{1}.rtf", path, r - 1),
                                              Range:="",ConfirmConversions:=False,
                                              Link:=False,Attachment:=False)
        Next

        Dim pos As Double = word.InchesToPoints(1)
        document.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        Dim range As Word.Range = document.Bookmarks.Item("\endofdoc").Range
        range.InsertParagraphAfter()
        range.InsertAfter("以上")
    End Sub
End Class
