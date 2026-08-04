' vybe-test: vb/vb_using_statement_advanced/using_statement_multiple_resources
' origin: languages/vb/tests/vb/test_vb_using_statement_advanced.rs

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

Imports System

Class Resource
    Implements IDisposable
    
    Public Name As String
    
    Public Sub New(n As String)
        Name = n
        __P(CStr("Acquired " & Name))
    End Sub
    
    Public Sub Dispose() Implements IDisposable.Dispose
        __P(CStr("Disposed " & Name))
    End Sub
End Class

Module M
    Sub Main()
        ' Multiple resources of the same type can be declared in one Using block
        Using r1 As New Resource("R1"), r2 As New Resource("R2")
            __P(CStr("Using " & r1.Name & " and " & r2.Name))
        End Using
        __Check("Acquired R1
Acquired R2
Using R1 and R2
Disposed R2
Disposed R1")
    End Sub
End Module
