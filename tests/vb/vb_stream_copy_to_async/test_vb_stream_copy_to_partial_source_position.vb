' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_copy_to_partial_source_position
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

Module Program
    Sub Main()
        Dim data As Byte() = {10, 20, 30, 40, 50}
        Using srcMs As New MemoryStream(data)
            srcMs.Position = 2 ' Skip first 2 bytes
            Using destMs As New MemoryStream()
                srcMs.CopyTo(destMs)
                __Check(CStr(String.Join(",", destMs.ToArray())), "30,40,50")
            End Using
        End Using
    End Sub
End Module
