' vybe-test: vb/vb_system_version_matrix/version_constructor_two_parts_uses_missing_build_and_revision_markers
' origin: languages/vb/tests/vb/test_vb_system_version_matrix.rs

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

Module M
    Sub Main()
        Dim version As New Version(9, 10)
        __Check(CStr(version.Major), "9")
        __Check(CStr(version.Minor), "10")
        __Check(CStr(version.Build), "-1")
        __Check(CStr(version.Revision), "-1")
    End Sub
End Module
