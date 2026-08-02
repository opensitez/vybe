' vybe-test: vb/vb_for_next_loops/for_each_empty_array
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim arr() As Integer = {}
Dim count = 0
For Each v In arr
count += 1
Next
Console.WriteLine(count)
End Sub
End Module
