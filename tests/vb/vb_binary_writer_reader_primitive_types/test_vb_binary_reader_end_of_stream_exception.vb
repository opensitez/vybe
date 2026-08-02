' vybe-test: vb/vb_binary_writer_reader_primitive_types/test_vb_binary_reader_end_of_stream_exception
' origin: languages/vb/tests/vb/test_vb_binary_writer_reader_primitive_types.rs

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
            Using br As New BinaryReader(ms)
                Try
                    Dim n = br.ReadInt32()
                Catch ex As EndOfStreamException
                    __Check(CStr("EndOfStreamException Caught"), "EndOfStreamException Caught")
                End Try
            End Using
        End Using
    End Sub
End Module
