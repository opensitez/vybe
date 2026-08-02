' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_dispatch_info_capture_throw
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
Imports System.Runtime.ExceptionServices

Module Program
    Private captured As ExceptionDispatchInfo

    Private Sub CauseError()
        Try
            Throw New InvalidOperationException("Captured Exception")
        Catch ex As Exception
            captured = ExceptionDispatchInfo.Capture(ex)
        End Try
    End Sub

    Sub Main()
        CauseError()
        Try
            captured.Throw()
        Catch ex As Exception
            __Check(CStr("Re-thrown: " & ex.Message), "Re-thrown: Captured Exception")
        End Try
    End Sub
End Module
