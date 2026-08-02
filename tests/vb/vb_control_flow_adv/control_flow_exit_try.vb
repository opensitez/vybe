' vybe-test: vb/vb_control_flow_adv/control_flow_exit_try
' origin: languages/vb/tests/vb/test_vb_control_flow_adv.rs

Module M
    Sub Main()
        Try
            Console.WriteLine("Start")
            Exit Try
            Console.WriteLine("Middle")
        Catch
        Finally
            Console.WriteLine("Finally")
        End Try
        Console.WriteLine("End")
    End Sub
End Module
