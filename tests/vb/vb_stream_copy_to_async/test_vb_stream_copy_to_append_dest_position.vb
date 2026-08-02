' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_copy_to_append_dest_position
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
        Dim d1 As Byte() = {1, 2}
        Dim d2 As Byte() = {3, 4}
        Using destMs As New MemoryStream()
            Using s1 As New MemoryStream(d1)
                s1.CopyTo(destMs)
            End Using
            Using s2 As New MemoryStream(d2)
                s2.CopyTo(destMs)
            End Using
            __Check(CStr(String.Join(",", destMs.ToArray())), "1,2,3,4")
        End Using
    End Sub
End Module
