' vybe-test: vb/vb_binary_writer_reader_primitive_types/test_vb_binary_writer_reader_roundtrip_primitives
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
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(42)
                bw.Write("BinaryString")
                bw.Write(3.14159)
                bw.Write(True)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim num = br.ReadInt32()
                Dim str = br.ReadString()
                Dim dbl = br.ReadDouble()
                Dim flag = br.ReadBoolean()
                __Check(CStr(num & "|" & str & "|" & dbl & "|" & flag), "42|BinaryString|3.14159|True")
            End Using
        End Using
    End Sub
End Module
