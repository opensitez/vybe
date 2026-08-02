' vybe-test: vb/vb_system_file_matrix/file_bytes_roundtrip
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
        Dim path As String = Path.Combine(Path.GetTempPath(), "vb_file_bytes_" & Guid.NewGuid().ToString("N"))
        Dim input() As Byte = {1, 2, 3, 4, 5}
        File.WriteAllBytes(path, input)
        Dim output() As Byte = File.ReadAllBytes(path)
        __Check(CStr(output.Length), "5")
        __Check(CStr(output(0)), "1")
        __Check(CStr(output(4)), "5")
        File.Delete(path)
    End Sub
End Module
