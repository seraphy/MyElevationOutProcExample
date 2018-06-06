Option Explicit
Dim obj

'Set obj = CreateObject("MyElevationOutProcSrv")
'Set obj = GetObject("new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
Set obj = GetObject("Elevation:Administrator!new:7AEFA37C-4494-4AE8-9378-0157A0B919AE")
WScript.ConnectObject obj, "obj_"

obj.Name = "FooBar"
obj.ShowHello()

Sub obj_NamePropertyChanging(ByVal name, ByRef cancel)
	WScript.Echo("Changing: " & name)
	' VBSÇ©ÇÁÇÕbyrefÇÃcancelílÇÕï‘ãpÇ≈Ç´Ç»Ç¢ÅB
End Sub

Sub obj_NamePropertyChanged(ByVal name)
	WScript.Echo("Changed: " & name)
End Sub
