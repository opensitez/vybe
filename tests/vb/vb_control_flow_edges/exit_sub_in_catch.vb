' vybe-test: vb/vb_control_flow_edges/exit_sub_in_catch
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Module M
    Sub Test()
        Try
            Throw New System.Exception()
        Catch
            Console.WriteLine("Catch")
            Exit Sub
        Finally
            Console.WriteLine("Finally")
        End Try
        Console.WriteLine("After")
    End Sub

    Sub Main()
        Test()
    End Sub
End Module
