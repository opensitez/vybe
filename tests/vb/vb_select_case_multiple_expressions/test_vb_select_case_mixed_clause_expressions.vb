' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_mixed_clause_expressions
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim x = 5
        Select Case x
            Case 1, 3, 7 To 10, Is > 100
                Console.WriteLine("Matched Group A")
            Case 4 To 6, Is < 0
                Console.WriteLine("Matched Group B")
            Case Else
                Console.WriteLine("Default")
        End Select
    End Sub
End Module
