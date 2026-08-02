' vybe-test: vb/vb_comprehensive/try_catch_no_error
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

Module M
    Sub Main()
        Try
            Console.WriteLine("no error")
        Catch ex As Exception
            Console.WriteLine("caught")
        End Try
        Console.WriteLine("done")
    End Sub
End Module
