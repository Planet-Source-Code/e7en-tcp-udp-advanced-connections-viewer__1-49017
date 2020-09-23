Attribute VB_Name = "modMisc"
Sub LvResize(lv As ListView)
For X = 1 To lv.ColumnHeaders.Count
    lv.ColumnHeaders(X).Width = lv.Width / lv.ColumnHeaders.Count - 20
Next

End Sub
