' vybe-test: vb/vb_stream_reader_writer_text/test_vb_memory_stream_reader_writer
' origin: languages/vb/tests/vb/test_vb_stream_reader_writer_text.rs

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
            Using writer As New StreamWriter(ms, System.Text.Encoding.UTF8, 1024, leaveOpen:=True)
                writer.WriteLine("Line1")
                writer.WriteLine("Line2")
                writer.Flush()
            End Using

            ms.Position = 0

            Using reader As New StreamReader(ms)
                __Check(CStr(reader.ReadLine()), "Line1")
                __Check(CStr(reader.ReadLine()), "Line2")
            End Using
        End Using
    End Sub
End Module
