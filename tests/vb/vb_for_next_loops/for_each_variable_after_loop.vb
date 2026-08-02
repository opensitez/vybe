' vybe-test: vb/vb_for_next_loops/for_each_variable_after_loop
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim arr() As Integer = {1}
Dim v As Integer
For Each v In arr
Next
Console.WriteLine(v)
End Sub
End Module
