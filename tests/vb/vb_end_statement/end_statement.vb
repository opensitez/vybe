' vybe-test: vb/vb_end_statement/end_statement
' origin: languages/vb/tests/vb/test_vb_end_statement.rs

Module M
    Sub DoSomething()
        Console.WriteLine("Start")
        End ' Terminates execution immediately
        Console.WriteLine("End") ' Unreachable
    End Sub

    Sub Main()
        DoSomething()
        Console.WriteLine("Main End") ' Unreachable
    End Sub
End Module
