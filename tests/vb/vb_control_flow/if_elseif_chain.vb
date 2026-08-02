' vybe-test: vb/vb_control_flow/if_elseif_chain
' origin: languages/vb/tests/vb/test_vb_control_flow.rs

Module M
    Sub Main()
        Dim score As Integer = 75
        If score >= 90 Then
            Console.WriteLine("A")
        ElseIf score >= 80 Then
            Console.WriteLine("B")
        ElseIf score >= 70 Then
            Console.WriteLine("C")
        Else
            Console.WriteLine("F")
        End If
    End Sub
End Module
