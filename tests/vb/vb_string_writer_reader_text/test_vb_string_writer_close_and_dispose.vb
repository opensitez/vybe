' vybe-test: vb/vb_string_writer_reader_text/test_vb_string_writer_close_and_dispose
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
        Dim sw As New StringWriter()
        sw.Write("Content")
        sw.Close()
        ' StringWriter.ToString still returns content after Dispose/Close!
        __Check(CStr(sw.ToString()), "Content")
    End Sub
End Module
