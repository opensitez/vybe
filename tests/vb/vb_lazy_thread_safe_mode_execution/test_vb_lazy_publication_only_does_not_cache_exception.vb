' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_publication_only_does_not_cache_exception
' origin: languages/vb/tests/vb/test_vb_lazy_thread_safe_mode_execution.rs

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
Imports System.Threading

Module Program
    Sub Main()
        Dim attempts = 0
        Dim lazyVal As New Lazy(Of String)(Function()
            attempts += 1
            If attempts = 1 Then Throw New InvalidOperationException("Fail 1")
            Return "Success"
            __Check("First Attempt Failed
Success|Attempts=2")
        End Function, LazyThreadSafetyMode.PublicationOnly)

        Try
            Dim v = lazyVal.Value
        Catch ex As InvalidOperationException
            __P(CStr("First Attempt Failed"))
        End Try

        Dim vSuccess = lazyVal.Value
        __P(CStr(vSuccess & "|Attempts=" & attempts))
    End Sub
End Module
