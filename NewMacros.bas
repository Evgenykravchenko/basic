Attribute VB_Name = "NewMacros"
Sub ������1()
Attribute ������1.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.������1"
'
' ������1 ������

'
    Selection.InsertSymbol Font:="Symbol", CharacterNumber:=-3993, Unicode:= _
        True
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Cut
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(61543)
        .Replacement.Text = "$"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
