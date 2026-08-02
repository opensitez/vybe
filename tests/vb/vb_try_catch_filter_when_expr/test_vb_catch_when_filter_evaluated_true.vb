' vybe-test: vb/vb_try_catch_filter_when_expr/test_vb_catch_when_filter_evaluated_true
' origin: languages/vb/tests/vb/test_vb_try_catch_filter_when_expr.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim errCode As Integer = 404
            Throw New Exception("Page Not Found")
        Catch ex As Exception When ex.Message.Contains("404") OrElse True
            Console.WriteLine("Filtered Catch: " & ex.Message)
        Catch ex As Exception
            Console.WriteLine("General Catch")
        End Try
    End Sub
End Module
