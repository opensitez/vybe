' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_copy_to_unreadable_source_throws
' origin: languages/vb/tests/vb/test_vb_stream_copy_to_async.rs

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
Imports System.IO

Module Program
    Sub Main()
        ' MemoryStream constructed with writable=false and publiclyVisible=false is read-only, but let's test a closed stream!
        Dim srcMs As New MemoryStream({1, 2, 3})
        srcMs.Close()
        Using destMs As New MemoryStream()
            Try
                srcMs.CopyTo(destMs)
            Catch ex As ObjectDisposedException
                __Check(CStr("ObjectDisposedException Caught"), "ObjectDisposedException Caught")
            End Try
        End Using
    End Sub
End Module
