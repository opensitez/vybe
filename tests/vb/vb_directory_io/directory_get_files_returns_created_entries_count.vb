' vybe-test: vb/vb_directory_io/directory_get_files_returns_created_entries_count
' origin: languages/vb/tests/vb/test_vb_directory_io.rs

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
        Dim dir As String = Path.Combine(Path.GetTempPath(), "vybe_files_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(dir)

        File.WriteAllText(Path.Combine(dir, "a.txt"), "1")
        File.WriteAllText(Path.Combine(dir, "b.txt"), "2")
        File.WriteAllText(Path.Combine(dir, "c.log"), "3")

        __Check(CStr(Directory.GetFiles(dir).Length), "3")
        __Check(CStr(Directory.GetFiles(dir, "*.txt").Length), "2")

        Directory.Delete(dir, True)
    End Sub
End Module
