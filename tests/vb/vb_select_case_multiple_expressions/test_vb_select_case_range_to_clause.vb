' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_range_to_clause
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim score = 85
        Select Case score
            Case 90 To 100
                Console.WriteLine("A")
            Case 80 To 89
                Console.WriteLine("B")
            Case 70 To 79
                Console.WriteLine("C")
            Case Else
                Console.WriteLine("F")
        End Select
    End Sub
End Module
