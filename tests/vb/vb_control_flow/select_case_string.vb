' vybe-test: vb/vb_control_flow/select_case_string
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim color As String = "red"
        Select Case color
            Case "red"
                Console.WriteLine("R")
            Case "green"
                Console.WriteLine("G")
            Case "blue"
                Console.WriteLine("B")
        End Select
    End Sub
End Module
