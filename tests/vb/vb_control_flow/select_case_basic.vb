' vybe-test: vb/vb_control_flow/select_case_basic
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim x As Integer = 2
        Select Case x
            Case 1
                Console.WriteLine("one")
            Case 2
                Console.WriteLine("two")
            Case 3
                Console.WriteLine("three")
            Case Else
                Console.WriteLine("other")
        End Select
    End Sub
End Module
