' vybe-test: vb/vb_system_version_matrix/version_equals_ignores_instance_identity
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
        Dim left As New Version(3, 4, 5, 6)
        Dim right As New Version(3, 4, 5, 6)
        Dim parsed As Version = Version.Parse(left.ToString())

        __Check(CStr(left.Equals(right)), "True")
        __Check(CStr(left.Equals(parsed)), "True")
        __Check(CStr(parsed.ToString() = left.ToString()), "True")
    End Sub
End Module
