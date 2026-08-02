' vybe-test: vb/vb_select_case_statements/select_case_evaluation_order
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Function F(v As Integer) As Integer
Console.WriteLine(v)
Return v
End Function
Sub Main()
Select Case F(5)
Case F(1), F(5), F(10)
Console.WriteLine("Match")
End Select
End Sub
End Module
