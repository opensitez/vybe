' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_disposes_disposable_enumerator
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Imports System
Imports System.Collections
Imports System.Collections.Generic
Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module


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
            __P(CStr("Enumerator Disposed"))
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
            __P(CStr(item))
        Next
        __Check("SingleItem
Enumerator Disposed")
    End Sub
End Module
