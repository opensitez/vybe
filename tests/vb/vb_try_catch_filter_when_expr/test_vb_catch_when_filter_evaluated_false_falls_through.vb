' vybe-test: vb/vb_try_catch_filter_when_expr/test_vb_catch_when_filter_evaluated_false_falls_through
' origin: languages/vb/tests/vb/test_vb_try_catch_filter_when_expr.rs

Imports System

Module Program
    Sub Main()
        Try
            Throw New InvalidOperationException("Operation Failed")
        Catch ex As Exception When ex.Message.Contains("Database")
            Console.WriteLine("Database Catch")
        Catch ex As Exception
            Console.WriteLine("Fallback Catch: " & ex.GetType().Name)
        End Try
    End Sub
End Module
