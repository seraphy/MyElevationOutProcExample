Private WithEvents obj As MyElevationOutProcSrv.MyElevationOutProcSrv

Public Sub TestCOM()
    Set obj = GetObject("Elevation:Administrator!new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
    'Set obj = GetObject("new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
    obj.name = "FooBar"
    Call obj.ShowHello
    Set obj = Nothing
End Sub

Private Sub obj_NamePropertyChanging(ByVal name As String, ByRef cancel As Boolean)
    If (MsgBox("Changing? " & name, vbYesNo, "Confirm") <> vbYes) Then
        cancel = True
    End If
End Sub

Private Sub obj_NamePropertyChanged(ByVal name As String)
    MsgBox "changed: " & name
End Sub
