' vybe-test: vb/vb_system_version_matrix/version_parse_without_revision_keeps_build
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
        Dim version As Version = Version.Parse("5.6.7")
        __Check(CStr(version.Major), "5")
        __Check(CStr(version.Minor), "6")
        __Check(CStr(version.Build), "7")
        __Check(CStr(version.Revision), "-1")
        __Check(CStr(version.ToString()), "5.6.7")
    End Sub
End Module
