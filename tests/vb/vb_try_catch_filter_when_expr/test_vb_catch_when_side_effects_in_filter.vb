' vybe-test: vb/vb_try_catch_filter_when_expr/test_vb_catch_when_side_effects_in_filter
' origin: languages/vb/tests/vb/test_vb_try_catch_filter_when_expr.rs

Imports System

Module Program
    Public FilterCount As Integer = 0

    Public Function LogAndCheck(ex As Exception) As Boolean
        FilterCount += 1
        Return False
    End Function

    Sub Main()
        Try
            Throw New Exception("Test")
        Catch ex As Exception When LogAndCheck(ex)
            Console.WriteLine("Caught in filter")
        Catch ex As Exception
            Console.WriteLine("Caught in fallback, count=" & FilterCount)
        End Try
    End Sub
End Module
