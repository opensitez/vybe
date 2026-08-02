' vybe-test: vb/vb_select_case_statements/select_case_multiple_values
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Dim x = 3
Select Case x
Case 1, 2, 3
Console.WriteLine("A")
Case Else
Console.WriteLine("B")
End Select
End Sub
End Module
