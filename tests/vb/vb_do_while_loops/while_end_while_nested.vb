' vybe-test: vb/vb_do_while_loops/while_end_while_nested
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Sub Main()
Dim i = 0, count = 0
While i < 2
Dim j = 0
While j < 3
count += 1
j += 1
End While
i += 1
End While
Console.WriteLine(count)
End Sub
End Module
