' vybe-test: vb/vb_control_flow_edges/do_until_loop_until
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Main()
        Dim i = 0
        ' Technically Do Until and Loop Until together is a syntax edge case
        Do Until i > 5
            i += 1
        Loop Until i = 3
        Console.WriteLine(i)
    End Sub
End Module
