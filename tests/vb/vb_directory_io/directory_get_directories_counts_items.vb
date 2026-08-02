' vybe-test: vb/vb_directory_io/directory_get_directories_counts_items
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
        Dim root As String = Path.Combine(Path.GetTempPath(), "vybe_tree_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Directory.CreateDirectory(Path.Combine(root, "a"))
        Directory.CreateDirectory(Path.Combine(root, "b"))

        __Check(CStr(Directory.GetDirectories(root).Length), "2")

        Directory.Delete(root, True)
    End Sub
End Module
