' vybe-test: vb/vb_system_path_matrix/path_get_file_name_and_root
' origin: languages/vb/tests/vb/test_vb_system_path_matrix.rs

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
        Dim p As String = "/tmp/reports/final.vb"
        __Check(CStr(Path.GetFileName(p)), "final.vb")
        __Check(CStr(Path.GetExtension(p)), ".vb")
        __Check(CStr(Path.GetDirectoryName(p)), "/tmp/reports")
    End Sub
End Module
