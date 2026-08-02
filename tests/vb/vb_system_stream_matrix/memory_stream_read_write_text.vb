' vybe-test: vb/vb_system_stream_matrix/memory_stream_read_write_text
' origin: languages/vb/tests/vb/test_vb_system_stream_matrix.rs

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

Module M
    Sub Main()
        Using ms As New MemoryStream()
            Dim writer As New StreamWriter(ms)
            writer.Write("abc")
            writer.Flush()
            __Check(CStr(ms.Length > 0), "True")
            ms.Position = 0
            Using reader As New StreamReader(ms)
                __Check(CStr(reader.ReadToEnd()), "abc")
            End Using
        End Using
    End Sub
End Module
