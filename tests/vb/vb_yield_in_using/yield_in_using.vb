' vybe-test: vb/vb_yield_in_using/yield_in_using
' origin: languages/vb/tests/vb/test_vb_yield_in_using.rs

Imports System
Imports System.Collections.Generic

Class Res
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed")
    End Sub
End Class

Module M
    Iterator Function Generate() As IEnumerable(Of Integer)
        Using r As New Res()
            Yield 1
            Yield 2
        End Using
    End Function

    Sub Main()
        For Each x In Generate()
            Console.WriteLine(x)
        Next
    End Sub
End Module
