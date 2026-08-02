' vybe-test: vb/vb_for_next_loops/for_each_type_coercion
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Option Strict Off
Module M
Sub Main()
Dim arr() As Object = {1, 2, 3}
Dim sum As Integer = 0
For Each v As Integer In arr
sum += v
Next
Console.WriteLine(sum)
End Sub
End Module
