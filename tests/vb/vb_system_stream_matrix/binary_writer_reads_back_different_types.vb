' vybe-test: vb/vb_system_stream_matrix/binary_writer_reads_back_different_types
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

Imports System.IO

Module M
    Sub Main()
        Using ms As New MemoryStream()
            Using writer As New BinaryWriter(ms)
                writer.Write(123)
                writer.Write(True)
                writer.Write("done")
            End Using
            ms.Position = 0
            Using reader As New BinaryReader(ms)
                __Check(CStr(reader.ReadInt32()), "123")
                __Check(CStr(reader.ReadBoolean()), "True")
                __Check(CStr(reader.ReadString()), "done")
            End Using
        End Using
    End Sub
End Module
