' vybe-test: vb/vb_system_stream_matrix/stream_seek_and_position_workflow
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
        Dim bytes() As Byte = {10, 20, 30, 40}
        Using ms As New MemoryStream(bytes)
            ms.Seek(2, SeekOrigin.Begin)
            __Check(CStr(ms.Position), "2")
            Dim one As Integer = ms.ReadByte()
            __Check(CStr(one), "30")
            __Check(CStr(ms.Position), "3")
        End Using
    End Sub
End Module
