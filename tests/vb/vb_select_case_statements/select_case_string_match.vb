' vybe-test: vb/vb_select_case_statements/select_case_string_match
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Module M
Sub Main()
Dim s = "Hello"
Select Case s
Case "Hi"
Console.WriteLine("A")
Case "Hello"
Console.WriteLine("B")
End Select
End Sub
End Module
