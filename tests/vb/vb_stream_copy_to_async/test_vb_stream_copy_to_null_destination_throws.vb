' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_copy_to_null_destination_throws
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
        Using srcMs As New MemoryStream({1})
            Try
                srcMs.CopyTo(Nothing)
            Catch ex As ArgumentNullException
                __Check(CStr("ArgumentNullException Caught"), "ArgumentNullException Caught")
            End Try
        End Using
    End Sub
End Module
