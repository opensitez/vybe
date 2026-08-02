' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_copy_to_with_custom_buffer_size
' origin: languages/vb/tests/vb/test_vb_stream_copy_to_async.rs

Imports System.IO

Module Program
    Sub Main()
        Dim data As Byte() = New Byte(99) {}
        For i As Integer = 0 To 99 : data(i) = CByte(i) : Next

        Using srcMs As New MemoryStream(data)
            Using destMs As New MemoryStream()
                srcMs.CopyTo(destMs, 16) ' 16-byte buffer size
                Console.WriteLine(destMs.Length)
            End Using
        End Using
    End Sub
End Module
