' vybe-test: vb/vb_do_while_loops/do_loop_condition_function_call
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Module M
Dim calls As Integer = 0
Function Check() As Boolean
calls += 1
Return calls < 3
End Function
Sub Main()
Do While Check()
Loop
Console.WriteLine(calls)
End Sub
End Module
