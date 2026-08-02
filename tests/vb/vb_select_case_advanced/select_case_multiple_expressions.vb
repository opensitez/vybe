' vybe-test: vb/vb_select_case_advanced/select_case_multiple_expressions
' origin: languages/vb/tests/vb/test_vb_select_case_advanced.rs

Module M
    Sub Main()
        Dim dayOfWeek As Integer = 6
        Select Case dayOfWeek
            Case 1, 7
                Console.WriteLine("Weekend")
            Case 2, 3, 4, 5, 6
                Console.WriteLine("Weekday")
            Case Else
                Console.WriteLine("Invalid")
        End Select
    End Sub
End Module
