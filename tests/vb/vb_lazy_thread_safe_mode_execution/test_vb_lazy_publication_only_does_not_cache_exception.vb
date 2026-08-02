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

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
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
        End Function, LazyThreadSafetyMode.PublicationOnly)

        Try
            Dim v = lazyVal.Value
        Catch ex As InvalidOperationException
            __Check(CStr("First Attempt Failed"), "First Attempt Failed")
        End Try

        Dim vSuccess = lazyVal.Value
        __Check(CStr(vSuccess & "|Attempts=" & attempts), "Success|Attempts=2")
    End Sub
End Module
