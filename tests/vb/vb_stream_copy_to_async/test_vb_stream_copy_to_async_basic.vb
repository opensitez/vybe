' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_copy_to_async_basic
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
    Private Async Function CopyStreamAsync() As Task(Of Integer)
        Dim data As Byte() = {10, 20, 30}
        Using srcMs As New MemoryStream(data)
            Using destMs As New MemoryStream()
                Await srcMs.CopyToAsync(destMs)
                Return CInt(destMs.Length)
            End Using
        End Using
    End Function

    Sub Main()
        Dim t = CopyStreamAsync()
        __Check(CStr(t.Result), "3")
    End Sub
End Module
