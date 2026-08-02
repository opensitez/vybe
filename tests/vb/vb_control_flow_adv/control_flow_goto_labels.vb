' vybe-test: vb/vb_control_flow_adv/control_flow_goto_labels
' origin: languages/vb/tests/vb/test_vb_control_flow_adv.rs

Module M
    Sub Main()
        Dim i = 0
    StartLabel:
        If i = 3 Then
            GoTo EndLabel
        End If
        Console.WriteLine(i)
        i += 1
        GoTo StartLabel
        
    EndLabel:
        Console.WriteLine("Done")
    End Sub
End Module
