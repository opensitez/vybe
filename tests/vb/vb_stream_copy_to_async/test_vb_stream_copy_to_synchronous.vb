' vybe-test: vb/vb_stream_copy_to_async/test_vb_stream_copy_to_synchronous
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
        Dim sourceData As Byte() = {1, 2, 3, 4, 5}
        Using srcMs As New MemoryStream(sourceData)
            Using destMs As New MemoryStream()
                srcMs.CopyTo(destMs)
                __Check(CStr(destMs.Length & "|" & String.Join(",", destMs.ToArray())), "5|1,2,3,4,5")
            End Using
        End Using
    End Sub
End Module
