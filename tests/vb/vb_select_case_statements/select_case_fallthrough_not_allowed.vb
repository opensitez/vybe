' vybe-test: vb/vb_select_case_statements/select_case_fallthrough_not_allowed
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Dim x = 1
Select Case x
Case 1
Console.WriteLine("A")
' No fallthrough to Case 2 implicitly
Case 2
Console.WriteLine("B")
End Select
End Sub
End Module
