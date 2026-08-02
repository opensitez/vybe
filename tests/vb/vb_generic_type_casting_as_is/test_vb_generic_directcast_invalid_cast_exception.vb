' vybe-test: vb/vb_generic_type_casting_as_is/test_vb_generic_directcast_invalid_cast_exception
' origin: languages/vb/tests/vb/test_vb_generic_type_casting_as_is.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim boxed As Object = "NotAnInt"
            Dim num As Integer = DirectCast(boxed, Integer)
            Console.WriteLine(num)
        Catch ex As InvalidCastException
            Console.WriteLine("DirectCast InvalidCastException Caught")
        End Try
    End Sub
End Module
