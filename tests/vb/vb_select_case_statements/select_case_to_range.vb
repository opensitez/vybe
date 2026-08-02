' vybe-test: vb/vb_select_case_statements/select_case_to_range
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Dim x = 15
Select Case x
Case 1 To 10
Console.WriteLine("A")
Case 11 To 20
Console.WriteLine("B")
End Select
End Sub
End Module
