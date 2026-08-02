' vybe-test: vb/vb_string_builder_capacity_ops/test_vb_sb_chunk_enumerator
' origin: languages/vb/tests/vb/test_vb_string_builder_capacity_ops.rs

Imports System.Text

Module Program
    Sub Main()
        Dim sb As New StringBuilder("Chunk Enumerator Test")
        Dim count As Integer = 0
        For Each chunk In sb.GetChunks()
            count += 1
        Next
        Console.WriteLine(count > 0)
    End Sub
End Module
