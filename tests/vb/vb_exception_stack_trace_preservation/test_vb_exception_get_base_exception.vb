' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_get_base_exception
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
    Sub Main()
        Try
            Dim root As New OverflowException("Root Cause")
            Dim mid As New InvalidOperationException("Mid Layer", root)
            Dim top As New Exception("Top Layer", mid)
            Throw top
        Catch ex As Exception
            __Check(CStr(ex.GetBaseException().Message), "Root Cause")
        End Try
    End Sub
End Module
