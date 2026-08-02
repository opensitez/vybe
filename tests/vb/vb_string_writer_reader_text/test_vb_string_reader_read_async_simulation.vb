' vybe-test: vb/vb_string_writer_reader_text/test_vb_string_reader_read_async_simulation
' origin: languages/vb/tests/vb/test_vb_string_writer_reader_text.rs

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
        Dim text = "AsyncLine1" & vbCrLf & "AsyncLine2"
        Using sr As New StringReader(text)
            Dim t1 = sr.ReadLineAsync()
            Dim t2 = sr.ReadLineAsync()
            __Check(CStr(t1.Result & "|" & t2.Result), "AsyncLine1|AsyncLine2")
        End Using
    End Sub
End Module
