' vybe-test: vb/vb_path_combine_get_filename_extension/test_vb_path_is_path_rooted
' origin: languages/vb/tests/vb/test_vb_path_combine_get_filename_extension.rs

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

Module Program
    Sub Main()
        __Check(CStr(Path.IsPathRooted("/abs/path") & "|" & Path.IsPathRooted("rel/path")), "True|False")
    End Sub
End Module
