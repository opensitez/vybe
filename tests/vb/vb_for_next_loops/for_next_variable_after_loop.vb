' vybe-test: vb/vb_for_next_loops/for_next_variable_after_loop
' origin: languages/vb/tests/vb/test_vb_for_next_loops.rs

Module M
Sub Main()
Dim i As Integer
For i = 1 To 3
Next
Console.WriteLine(i)
End Sub
End Module
