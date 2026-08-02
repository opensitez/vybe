' vybe-test: vb/vb_select_case_multiple_expressions/test_vb_select_case_decimal_values
' origin: languages/vb/tests/vb/test_vb_select_case_multiple_expressions.rs

Module Program
    Sub Main()
        Dim price As Decimal = 19.99D
        Select Case price
            Case 0.0D To 9.99D
                Console.WriteLine("Low")
            Case 10.0D To 49.99D
                Console.WriteLine("Medium")
            Case Else
                Console.WriteLine("High")
        End Select
    End Sub
End Module
