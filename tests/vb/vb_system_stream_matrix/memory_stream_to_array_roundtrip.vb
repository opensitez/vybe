' vybe-test: vb/vb_system_stream_matrix/memory_stream_to_array_roundtrip
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

Module M
    Sub Main()
        Dim bytes() As Byte = {1, 2, 3}
        Using ms As New MemoryStream(bytes)
            Dim cloned As Byte() = ms.ToArray()
            __Check(CStr(cloned.Length), "3")
            __Check(CStr(cloned(0)), "1")
            __Check(CStr(cloned(2)), "3")
        End Using
    End Sub
End Module
