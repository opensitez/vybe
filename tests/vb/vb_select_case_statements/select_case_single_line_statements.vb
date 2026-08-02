' vybe-test: vb/vb_select_case_statements/select_case_single_line_statements
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Select Case 1: Case 1: Console.WriteLine("A"): Case 2: Console.WriteLine("B"): End Select
End Sub
End Module
