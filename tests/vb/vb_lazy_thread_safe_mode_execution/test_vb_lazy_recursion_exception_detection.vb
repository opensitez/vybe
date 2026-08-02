' vybe-test: vb/vb_lazy_thread_safe_mode_execution/test_vb_lazy_recursion_exception_detection
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
    Private recLazy As Lazy(Of Integer)

    Sub Main()
        recLazy = New Lazy(Of Integer)(Function() recLazy.Value + 1)
        Try
            Dim val = recLazy.Value
        Catch ex As InvalidOperationException
            __Check(CStr("Recursive Lazy Initialization Caught"), "Recursive Lazy Initialization Caught")
        End Try
    End Sub
End Module
