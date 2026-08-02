' vybe-test: vb/vb_select_case_advanced/select_case_to_range
' origin: languages/vb/tests/vb/test_vb_select_case_advanced.rs

Module M
    Sub Main()
        Dim score As Integer = 85
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
