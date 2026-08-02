' vybe-test: vb/vb_exception_filters/exception_filters_when
' origin: languages/vb/tests/vb/test_vb_exception_filters.rs

Module M
    Sub Main()
        Try
            Throw New System.InvalidOperationException("Test 1")
        Catch ex As Exception When ex.Message.Contains("2")
            Console.WriteLine("Caught 2")
        Catch ex As Exception When ex.Message.Contains("1")
            Console.WriteLine("Caught 1")
        Catch ex As Exception
            Console.WriteLine("Caught other")
        End Try
    End Sub
End Module
