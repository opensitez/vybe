' vybe-test: vb/vb_memory_stream_capacity_expansion/test_vb_memory_stream_try_get_buffer_segment
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
        Using ms As New MemoryStream()
            ms.WriteByte(50)
            Dim segment As ArraySegment(Of Byte)
            Dim ok = ms.TryGetBuffer(segment)
            __Check(CStr(ok & ":" & segment.Array(0)), "True:50")
        End Using
    End Sub
End Module
