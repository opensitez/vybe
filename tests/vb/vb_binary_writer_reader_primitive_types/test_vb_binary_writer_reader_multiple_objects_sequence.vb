' vybe-test: vb/vb_binary_writer_reader_primitive_types/test_vb_binary_writer_reader_multiple_objects_sequence
' origin: languages/vb/tests/vb/test_vb_binary_writer_reader_primitive_types.rs

Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                For i As Integer = 1 To 5
                    bw.Write("Item_" & i)
                Next
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                For i As Integer = 1 To 5
                    Console.WriteLine(br.ReadString())
                Next
            End Using
        End Using
    End Sub
End Module
