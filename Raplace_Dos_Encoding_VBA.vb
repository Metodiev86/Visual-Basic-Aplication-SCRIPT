Sub ReplaceSymbolsFromFile()
    Dim replacements As Object
    Set replacements = CreateObject("Scripting.Dictionary")
    
    ' Добавяне на символи и техните замени
    replacements.Add ChrW(9619), "т"
    replacements.Add ChrW(9617), "р"
    replacements.Add ChrW(9618), "с"
    replacements.Add ChrW(9558), "ч"
    replacements.Add ChrW(9474), "у"
    replacements.Add ChrW(9563), "ю"
    replacements.Add ChrW(9553), "ъ"
    replacements.Add ChrW(9569), "х"
    replacements.Add ChrW(9571), "щ"
    replacements.Add ChrW(9557), "ш"
    replacements.Add ChrW(9570), "ц"
    replacements.Add ChrW(9488), "я"
    replacements.Add ChrW(9508), "ф"
    
    Dim key As Variant
    For Each key In replacements.Keys
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = key
            .Replacement.Text = replacements(key)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next key
    
    ActiveDocument.Save
End Sub
