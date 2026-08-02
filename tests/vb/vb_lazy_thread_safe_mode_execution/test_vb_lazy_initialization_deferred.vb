' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_initialization_deferred
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
        Dim initialized = False
        Dim lazyVal As New Lazy(Of Integer)(Function()
            initialized = True
            Return 42
        End Function)

        __Check(CStr("IsCreatedBefore: " & lazyVal.IsValueCreated), "IsCreatedBefore: False")
        Dim v = lazyVal.Value
        __Check(CStr("IsCreatedAfter: " & lazyVal.IsValueCreated & "|Val=" & v), "IsCreatedAfter: True|Val=42")
    End Sub
End Module
