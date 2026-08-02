' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_thread_safety_mode_execution_and_publication
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
        Dim lazyVal As New Lazy(Of Integer)(Function() 100, LazyThreadSafetyMode.ExecutionAndPublication)
        __Check(CStr(lazyVal.IsValueCreated & "|" & lazyVal.Value), "False|100")
    End Sub
End Module
