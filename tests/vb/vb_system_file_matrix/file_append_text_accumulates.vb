' vybe-test: vb/vb_system_file_matrix/file_append_text_accumulates
' origin: languages/vb/tests/vb/test_vb_system_file_matrix.rs

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

Imports System
Imports System.IO

Module M
    Sub Main()
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_file_append_" & Guid.NewGuid().ToString("N"))
        File.WriteAllText(path, "left")
        File.AppendAllText(path, "-right")
        __Check(CStr(File.ReadAllText(path)), "left-right")
        File.Delete(path)
    End Sub
End Module
