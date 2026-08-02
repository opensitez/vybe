' vybe-test: vb/vb_string_writer_reader_text/test_vb_string_writer_write_line
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
        Using sw As New StringWriter()
            sw.WriteLine("Line 1")
            sw.WriteLine("Line 2")
            __Check(CStr(sw.ToString().Contains("Line 1") AndAlso sw.ToString().Contains("Line 2")), "True")
        End Using
    End Sub
End Module
