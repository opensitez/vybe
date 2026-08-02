' vybe-test: vb/vb_control_flow/select_case_else
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim x As Integer = 99
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case Else
                Console.WriteLine("default")
        End Select
    End Sub
End Module
