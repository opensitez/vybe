' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_chunked_stream_reading
' origin: languages/vb/tests/vb/test_vb_convert_to_base64_string.rs

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

Module Program
    Sub Main()
        Dim data As Byte() = Encoding.UTF8.GetBytes("Chunk1Chunk2Chunk3")
        Using ms As New MemoryStream(data)
            Dim buffer(5) As Byte
            Dim readCount = ms.Read(buffer, 0, 6)
            Dim chunkB64 = Convert.ToBase64String(buffer, 0, readCount)
            __Check(CStr(chunkB64), "Q2h1bmsx")
        End Using
    End Sub
End Module
