' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_disposes_disposable_enumerator
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System
Imports System.Collections
Imports System.Collections.Generic

Class DisposableCollection
    Implements IEnumerable(Of String)

    Private Class CustomDispEnum
        Implements IEnumerator(Of String)
        Public Property Current As String Implements IEnumerator(Of String).Current
        Private Property Current1 As Object Implements IEnumerator.Current
            Get
                Return Current
            End Get
        End Property

        Private readDone As Boolean = False
        Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
            If Not readDone Then
                Current = "SingleItem"
                readDone = True
                Return True
            End If
            Return False
        End Function

        Public Sub Reset() Implements IEnumerator.Reset
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Console.WriteLine("Enumerator Disposed")
        End Sub
    End Class

    Public Function GetEnumerator() As IEnumerator(Of String) Implements IEnumerable(Of String).GetEnumerator
        Return New CustomDispEnum()
    End Function

    Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function
End Class

Module Program
    Sub Main()
        Dim col As New DisposableCollection()
        For Each item In col
            Console.WriteLine(item)
        Next
    End Sub
End Module
