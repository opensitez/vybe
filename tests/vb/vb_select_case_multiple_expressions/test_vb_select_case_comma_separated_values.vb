' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_comma_separated_values
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim day = 6
        Select Case day
            Case 1, 7
                Console.WriteLine("Weekend")
            Case 2, 3, 4, 5, 6
                Console.WriteLine("Weekday")
            Case Else
                Console.WriteLine("Invalid")
        End Select
    End Sub
End Module
