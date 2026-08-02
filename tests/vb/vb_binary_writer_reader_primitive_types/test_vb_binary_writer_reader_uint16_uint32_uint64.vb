' vybe-test: vb/vb_binary_writer_reader_primitive_types/test_vb_binary_writer_reader_uint16_uint32_uint64
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
                bw.Write(65000US)
                bw.Write(4000000000UI)
                bw.Write(18000000000000000000UL)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim u16 = br.ReadUInt16()
                Dim u32 = br.ReadUInt32()
                Dim u64 = br.ReadUInt64()
                __Check(CStr(u16 & "|" & u32 & "|" & u64), "65000|4000000000|18000000000000000000")
            End Using
        End Using
    End Sub
End Module
