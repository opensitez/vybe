' vybe-test: vb/vb_do_while_loops/do_loop_boolean_conversion
' origin: languages/vb/tests/vb/test_vb_do_while_loops.rs

Option Strict Off
Module M
Sub Main()
Dim x = 0
Do While "True"
x += 1
Exit Do
Loop
Console.WriteLine(x)
End Sub
End Module
