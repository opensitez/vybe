' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_exception_cached_on_failure
' origin: languages/vb/tests/vb/test_vb_lazy_thread_safe_mode_execution.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Module Program
    Sub Main()
        Dim attempts = 0
        Dim lazyVal As New Lazy(Of String)(Function()
            attempts += 1
            Throw New InvalidOperationException("Fail " & attempts)
        End Function)

        Try
            Dim v = lazyVal.Value
        Catch ex1 As InvalidOperationException
            __Check(CStr(ex1.Message), "Fail 1")
        End Try

        Try
            Dim v2 = lazyVal.Value
        Catch ex2 As InvalidOperationException
            __Check(CStr(ex2.Message & "|Attempts=" & attempts), "Fail 1|Attempts=1")
        End Try
    End Sub
End Module
