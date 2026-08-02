' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_is_relational_operator
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim val = 150
        Select Case val
            Case Is < 0
                Console.WriteLine("Negative")
            Case Is >= 100
                Console.WriteLine("High")
            Case Else
                Console.WriteLine("Normal")
        End Select
    End Sub
End Module
