' vybe-test: vb/vb_select_case_statements/select_case_is_operator
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Dim x = 25
Select Case x
Case Is > 20
Console.WriteLine("A")
Case Else
Console.WriteLine("B")
End Select
End Sub
End Module
