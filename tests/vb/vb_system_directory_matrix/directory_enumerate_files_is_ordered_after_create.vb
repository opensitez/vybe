' vybe-test: vb/vb_system_directory_matrix/directory_enumerate_files_is_ordered_after_create
' origin: languages/vb/tests/vb/test_vb_system_directory_matrix.rs

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
        Dim root As String = Path.Combine(Path.GetTempPath(), "vb_enumerate_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        File.WriteAllText(Path.Combine(root, "b.txt"), "2")
        File.WriteAllText(Path.Combine(root, "a.txt"), "1")
        Dim items As String() = Directory.GetFiles(root, "*.txt")
        __Check(CStr(items.Length), "2")
        __Check(CStr(Path.GetFileName(items(0))), "b.txt")
        Directory.Delete(root, True)
    End Module
End Module
