' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_array_of_handles
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

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
Imports System.Runtime.InteropServices

Class ArrayItemHandle
    Inherits SafeHandle
    Public Sub New(val As Long)
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(val))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim handles As ArrayItemHandle() = {New ArrayItemHandle(1), New ArrayItemHandle(2)}
        For Each h In handles
            h.Dispose()
        Next
        __P(CStr("All Array Handles Disposed"))
        __Check("All Array Handles Disposed")
    End Sub
End Module
