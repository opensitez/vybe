' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_double_free_throws
' origin: languages/vb/tests/vb/test_vb_gchandle_alloc_free.rs

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
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim handle = GCHandle.Alloc(obj)
        handle.Free()
        Try
            handle.Free()
        Catch ex As InvalidOperationException
            __Check(CStr("InvalidOperationException Caught on Double Free"), "InvalidOperationException Caught on Double Free")
        End Try
    End Sub
End Module
