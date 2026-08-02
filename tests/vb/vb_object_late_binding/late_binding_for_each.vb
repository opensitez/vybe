' vybe-test: vb/vb_object_late_binding/late_binding_for_each
' origin: languages/vb/tests/vb/test_vb_object_late_binding.rs

Option Strict Off
Module M
Sub Main()
Dim arr() As Integer = {10}
Dim obj As Object = arr
For Each v In obj
Console.WriteLine(v)
Next
End Sub
End Module
