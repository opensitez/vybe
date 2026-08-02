' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_read_async_buffer
' origin: languages/vb/tests/vb/test_vb_stream_copy_to_async.rs

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
Imports System.Threading.Tasks

Module Program
    Private Async Function ReadAsyncTest() As Task(Of Integer)
        Dim data As Byte() = {100, 200}
        Using ms As New MemoryStream(data)
            Dim buffer(1) As Byte
            Dim readCount = Await ms.ReadAsync(buffer, 0, 2)
            Return readCount
        End Using
    End Function

    Sub Main()
        Dim t = ReadAsyncTest()
        __Check(CStr(t.Result), "2")
    End Sub
End Module
