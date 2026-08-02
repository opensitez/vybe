' vybe-test: vb/vb_optional_paramarrays/paramarray_passing_array
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
Dim arr() As Integer = {1, 2, 3}
Console.WriteLine(Sum(arr))
End Sub
End Module
