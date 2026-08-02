' vybe-test: vb/vb_select_case_statements/select_case_boolean_false
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Select Case False
Case True
Console.WriteLine("T")
Case False
Console.WriteLine("F")
End Select
End Sub
End Module
