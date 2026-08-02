' vybe-test: vb/vb_do_while_loops/while_condition_function_call
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Dim calls As Integer = 0
Function Check() As Boolean
calls += 1
Return calls < 3
End Function
Sub Main()
While Check()
End While
Console.WriteLine(calls)
End Sub
End Module
