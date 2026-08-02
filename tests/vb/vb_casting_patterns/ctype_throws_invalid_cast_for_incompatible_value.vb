' vybe-test: vb/vb_casting_patterns/ctype_throws_invalid_cast_for_incompatible_value
' origin: languages/vb/tests/vb/test_vb_casting_patterns.rs

Imports System

Module M
    Sub Main()
        Dim o As Object = True
        Try
            Console.WriteLine(CType(o, Integer))
            Console.WriteLine("NoCast")
        Catch ex As InvalidCastException
            Console.WriteLine("CastFailed")
        End Try
    End Sub
End Module
