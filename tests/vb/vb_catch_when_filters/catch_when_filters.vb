' vybe-test: vb/vb_catch_when_filters/catch_when_filters
' origin: languages/vb/tests/vb/test_vb_catch_when_filters.rs

Module M
    Sub Main()
        Dim errorCode As Integer = 404
        
        Try
            Throw New System.Exception("HTTP Error")
        Catch ex As System.Exception When errorCode = 500
            Console.WriteLine("Server Error")
        Catch ex As System.Exception When errorCode = 404
            Console.WriteLine("Not Found")
        Catch ex As System.Exception
            Console.WriteLine("Other Error")
        End Try
    End Sub
End Module
