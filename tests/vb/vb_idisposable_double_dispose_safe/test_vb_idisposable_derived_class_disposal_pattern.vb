' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_derived_class_disposal_pattern
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

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


Class BaseRes
    Implements IDisposable
    Protected BaseDisposed As Boolean = False

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not BaseDisposed Then
            If disposing Then __P(CStr("Base Managed Cleaned"))
            BaseDisposed = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Class DerivedRes
    Inherits BaseRes

    Private DerivedDisposed As Boolean = False

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not DerivedDisposed Then
            If disposing Then __P(CStr("Derived Managed Cleaned"))
            DerivedDisposed = True
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New DerivedRes()
        res.Dispose()
        __Check("Derived Managed Cleaned
Base Managed Cleaned")
    End Sub
End Module
