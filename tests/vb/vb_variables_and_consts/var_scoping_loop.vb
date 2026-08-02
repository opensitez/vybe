' vybe-test: vb/vb_variables_and_consts/var_scoping_loop
' origin: languages/vb/tests/vb/test_vb_variables_and_consts.rs

Module M
Sub Main()
For i = 1 To 2
Dim x = i
Next
' x is accessible outside loop in older VB, but VB.NET block scopes Dim in loops depending on strictness. Actually, Dim inside For is block-scoped.
Console.WriteLine("Parsed")
End Sub
End Module
