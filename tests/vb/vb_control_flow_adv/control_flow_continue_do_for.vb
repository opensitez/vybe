' vybe-test: vb/vb_control_flow_adv/control_flow_continue_do_for
' origin: languages/vb/tests/vb/test_vb_control_flow_adv.rs

Module M
    Sub Main()
        For i As Integer = 1 To 3
            If i = 2 Then Continue For
            Console.WriteLine("For " & i)
        Next
        
        Dim j = 0
        Do While j < 3
            j += 1
            If j = 2 Then Continue Do
            Console.WriteLine("Do " & j)
        Loop
    End Sub
End Module
