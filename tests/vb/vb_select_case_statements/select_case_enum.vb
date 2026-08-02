' vybe-test: vb/vb_select_case_statements/select_case_enum
' origin: languages/vb/tests/vb/test_vb_select_case_statements.rs

Enum E
A
B
End Enum
Module M
Sub Main()
Dim x = E.B
Select Case x
Case E.A
Console.WriteLine("A")
Case E.B
Console.WriteLine("B")
End Select
End Sub
End Module
