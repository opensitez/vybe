' vybe-test: vb/vb_memory_stream_capacity_expansion/test_vb_memory_stream_closed_stream_throws_object_disposed
' origin: languages/vb/tests/vb/test_vb_memory_stream_capacity_expansion.rs

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
        Dim ms As New MemoryStream()
        ms.Close()
        Try
            ms.WriteByte(1)
        Catch ex As ObjectDisposedException
            __Check(CStr("ObjectDisposedException Caught"), "ObjectDisposedException Caught")
        End Try
    End Sub
End Module
