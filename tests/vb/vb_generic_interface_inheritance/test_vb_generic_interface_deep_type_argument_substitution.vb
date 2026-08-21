' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_deep_type_argument_substitution
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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


Interface IDataPipeline(Of TIn, TOut)
    Function Process(input As IEnumerable(Of TIn)) As List(Of TOut)
End Interface

Class StringLengthPipeline
    Implements IDataPipeline(Of String, Integer)
    Public Function Process(input As IEnumerable(Of String)) As List(Of Integer) Implements IDataPipeline(Of String, Integer).Process
        Dim res As New List(Of Integer)()
        For Each s In input
            res.Add(s.Length)
        Next
        Return res
    End Function
End Class

Module Program
    Sub Main()
        Dim p As IDataPipeline(Of String, Integer) = New StringLengthPipeline()
        Dim lengths = p.Process({"A", "BB", "CCC"})
        __P(CStr(String.Join(",", lengths)))
        __Check("1,2,3")
    End Sub
End Module
