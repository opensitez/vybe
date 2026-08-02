' vybe-test: vb/vb_system_stream_matrix/stream_copy_to_memory_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_stream_matrix.rs

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
Imports System.Text

Module M
    Sub Main()
        Dim source As New MemoryStream(Encoding.UTF8.GetBytes("copy-source"))
        Dim destination As New MemoryStream()
        source.CopyTo(destination)
        __Check(CStr(destination.Length), "11")
        __Check(CStr(Encoding.UTF8.GetString(destination.ToArray())), "copy-source")
    End Sub
End Module
