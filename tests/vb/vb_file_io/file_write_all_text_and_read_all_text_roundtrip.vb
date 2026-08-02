' vybe-test: vb/vb_file_io/file_write_all_text_and_read_all_text_roundtrip
' origin: languages/vb/tests/vb/test_vb_file_io.rs

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
        Dim path As String = Path.GetTempFileName()
        File.WriteAllText(path, "hello")
        __Check(CStr(File.ReadAllText(path)), "hello")
        File.Delete(path)
    End Sub
End Module
