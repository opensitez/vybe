' vybe-test: vb/vb_path_api_matrix/path_get_temp_file_name_extension
' origin: languages/vb/tests/vb/test_vb_path_api_matrix.rs

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
        Dim p As String = Path.GetTempFileName()
        __Check(CStr(Path.GetExtension(p)), ".tmp")
        __Check(CStr(p.Contains(System.IO.Path.GetTempPath())), "True")
    End Sub
End Module
