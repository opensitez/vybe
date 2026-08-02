' vybe-test: vb/vb_for_next_loops/for_each_continue_for
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim arr() As Integer = {1, 2, 3, 4}
Dim sum = 0
For Each v In arr
If v = 2 Then Continue For
sum += v
Next
Console.WriteLine(sum)
End Sub
End Module
