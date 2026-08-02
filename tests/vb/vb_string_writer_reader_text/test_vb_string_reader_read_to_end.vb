' vybe-test: vb/vb_string_writer_reader_text/test_vb_string_reader_read_to_end
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
        Dim text = "Header" & vbLf & "Body" & vbLf & "Footer"
        Using sr As New StringReader(text)
            Dim line1 = sr.ReadLine()
            Dim rest = sr.ReadToEnd()
            __Check(CStr(line1 & "|" & rest.Replace(vbLf, "-")), "Header|Body-Footer")
        End Using
    End Sub
End Module
