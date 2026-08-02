' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_target_site_method_reflection
' origin: languages/vb/tests/vb/test_vb_exception_stack_trace_preservation.rs

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
    Private Sub FailingMethod()
        Throw New InvalidOperationException("TargetSite")
    End Sub

    Sub Main()
        Try
            FailingMethod()
        Catch ex As Exception
            __Check(CStr(ex.TargetSite IsNot Nothing AndAlso ex.TargetSite.Name = "FailingMethod"), "True")
        End Try
    End Sub
End Module
