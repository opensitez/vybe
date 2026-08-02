' vybe-test: vb/vb_optional_arguments/optional_arguments_can_control_loop_iterations
' origin: languages/vb/tests/vb/test_vb_optional_arguments.rs

Module M
    Function CountUp(Optional repeatCount As Integer = 3) As Integer
        Dim total As Integer = 0
        For i As Integer = 1 To repeatCount
            total = total + i
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(CountUp())
        Console.WriteLine(CountUp(4))
    End Sub
End Module
