' vybe-test: vb/vb_casting_patterns/direct_vs_try_cast_difference_with_incompatible_types
' origin: languages/vb/tests/vb/test_vb_casting_patterns.rs

Imports System

Module M
    Sub Main()
        Dim boxed As Object = 99
        Dim castResult As Object = TryCast(boxed, String)
        Console.WriteLine(castResult Is Nothing)

        Try
            Dim direct As String = DirectCast(boxed, String)
            Console.WriteLine(direct)
        Catch ex As InvalidCastException
            Console.WriteLine("DirectFailed")
        End Try
    End Sub
End Module
