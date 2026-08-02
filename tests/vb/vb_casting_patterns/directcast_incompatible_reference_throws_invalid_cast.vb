' vybe-test: vb/vb_casting_patterns/directcast_incompatible_reference_throws_invalid_cast
' origin: languages/vb/tests/vb/test_vb_casting_patterns.rs

Imports System

Class A
End Class

Class B
End Class

Module M
    Sub Main()
        Dim a As Object = New A()
        Try
            Dim b As B = DirectCast(a, B)
            Console.WriteLine("NoCast")
        Catch ex As InvalidCastException
            Console.WriteLine("CastFailed")
        End Try
    End Sub
End Module
