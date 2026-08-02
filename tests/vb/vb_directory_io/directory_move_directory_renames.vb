' vybe-test: vb/vb_directory_io/directory_move_directory_renames
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
        Dim source As String = Path.Combine(Path.GetTempPath(), "vybe_src_" & Guid.NewGuid().ToString("N"))
        Dim target As String = Path.Combine(Path.GetTempPath(), "vybe_dst_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(source)
        Directory.Move(source, target)

        __Check(CStr(Directory.Exists(source)), "False")
        __Check(CStr(Directory.Exists(target)), "True")

        Directory.Delete(target)
    End Sub
End Module
