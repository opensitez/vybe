' vybe-test: vb/vb_memory_stream_capacity_expansion/test_vb_memory_stream_get_buffer_throws_when_not_publicly_visible
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
        Dim data As Byte() = {1, 2, 3}
        ' MemoryStream constructed with publiclyVisible=false
        Using ms As New MemoryStream(data, 0, 3, False, False)
            Try
                ms.GetBuffer()
            Catch ex As UnauthorizedAccessException
                __Check(CStr("UnauthorizedAccessException Caught"), "UnauthorizedAccessException Caught")
            End Try
        End Using
    End Sub
End Module
