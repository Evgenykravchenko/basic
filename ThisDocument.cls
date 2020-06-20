VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub FormatRange()
 Dim rngFormat As Range
 Set rngFormat = ActiveDocument.Range( _
 Start:=ActiveDocument.Paragraphs(1).Range.Start, _
 End:=ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.End)
  
' With rngFormat
' .Font.Name = "Times New Roman"
' .ParagraphFormat.Alignment = wdAlignParagraphJustify
' .Font.Size = 10
' .Italic = False
' End With
 Dim s As String
 s = StrConv("", vbUnicode)
 For i = 1 To rngFormat.Characters.Count
    c = rngFormat.Characters.Item(i)
    a = AscW(rngFormat.Characters.Item(i))
    '61543
    If rngFormat.Characters.Item(i) = ChrW(61543) Then
        c = ChrW(947)
    End If
    If rngFormat.Characters.Item(i).Bold Then
        c = "<b>" + c + "</b>"
    End If
    If rngFormat.Characters.Item(i).Font.Superscript = -1 Then
        c = "<sup>" + c + "</sup>"
    End If
    s = s + c
 Next

Const strFN As String = "R:\1.txt"
    Open strFN For Output As #1
    Print #1, s
    Close #1

CreateTextFile:
 FileName = "R:\2.txt"
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set out = fso.CreateTextFile(FileName, True, True)
 out.WriteLine (s)
 out.Close
End Sub







