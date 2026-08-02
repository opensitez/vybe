' vybe-test: vb/vb_comprehensive/select_case_multiple_values
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Dim day As Integer = 6
        Select Case day
            Case 1, 2, 3, 4, 5
                Console.WriteLine("weekday")
            Case 6, 7
                Console.WriteLine("weekend")
            Case Else
                Console.WriteLine("unknown")
        End Select
    End Sub
End Module
