' vybe-test: vb/vb_select_case_statements/select_case_type_mismatch_throws
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Option Strict Off
Module M
Sub Main()
Dim x = "ABC"
Try
Select Case x
Case 1 To 10
Console.WriteLine("A")
End Select
Catch
Console.WriteLine("Caught")
End Try
End Sub
End Module
