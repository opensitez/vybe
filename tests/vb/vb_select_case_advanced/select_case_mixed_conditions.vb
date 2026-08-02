' vybe-test: vb/vb_select_case_advanced/select_case_mixed_conditions
' origin: languages/vb/tests/vb/test_vb_select_case_advanced.rs

Module M
    Sub Main()
        Dim value As Integer = 50
        Select Case value
            Case 1 To 10, 20 To 30, Is >= 100
                Console.WriteLine("Group 1")
            Case 40, 50, 60
                Console.WriteLine("Group 2")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
