' vybe-test: vb/vb_for_next_loops/for_each_inline_declaration
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim arr() As Integer = {1, 2, 3}
Dim sum = 0
For Each v As Integer In arr
sum += v
Next
Console.WriteLine(sum)
End Sub
End Module
