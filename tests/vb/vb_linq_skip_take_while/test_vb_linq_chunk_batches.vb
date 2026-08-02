' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_chunk_batches
' origin: languages/vb/tests/vb/test_vb_linq_skip_take_while.rs

Imports System.Linq

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        Dim chunks = numbers.Chunk(2)
        For Each chunk In chunks
            Console.WriteLine(String.Join("-", chunk))
        Next
    End Sub
End Module
