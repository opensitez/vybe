' vybe-test: vb/vb_select_case_statements/select_case_exit_select
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Dim x = 1
Select Case x
Case 1
Console.WriteLine("A")
Exit Select
Console.WriteLine("B")
End Select
End Sub
End Module
