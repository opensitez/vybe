' vybe-test: vb/vb_file_io/file_read_all_bytes_returns_expected_length
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
        File.WriteAllBytes(path, New Byte() {1, 2, 3, 4, 5})
        __Check(CStr(File.ReadAllBytes(path).Length), "5")
        File.Delete(path)
    End Sub
End Module
