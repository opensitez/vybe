' vybe-test: vb/vb_system_directory_matrix/directory_info_full_path_and_name
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
        Dim root As New DirectoryInfo(Path.Combine(Path.GetTempPath(), "vb_info_" & Guid.NewGuid().ToString("N")))
        root.Create()
        Dim info As New DirectoryInfo(root.FullName)
        __Check(CStr(info.Exists), "True")
        __Check(CStr(info.Name.StartsWith("vb_info_")), "True")
        root.Delete()
    End Sub
End Module
