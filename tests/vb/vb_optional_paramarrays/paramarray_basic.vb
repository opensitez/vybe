' vybe-test: vb/vb_optional_paramarrays/paramarray_basic
' origin: languages/vb/tests/vb/test_vb_optional_paramarrays.rs

Module M
Function Sum(ParamArray args() As Integer) As Integer
Dim s = 0
For Each v In args
s += v
Next
Return s
End Function
Sub Main()
Console.WriteLine(Sum(1, 2, 3))
End Sub
End Module
