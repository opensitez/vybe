' vybe-test: vb/vb_memory_stream_capacity_expansion/test_vb_memory_stream_seek_begin_current_end
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

Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.Write({10, 20, 30, 40, 50}, 0, 5)
            ms.Seek(1, SeekOrigin.Begin)
            Dim b1 = ms.ReadByte()
            ms.Seek(-1, SeekOrigin.End)
            Dim b2 = ms.ReadByte()
            __Check(CStr(b1 & "|" & b2), "20|50")
        End Using
    End Sub
End Module
