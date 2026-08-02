' vybe-test: vb/vb_system_version_matrix/version_roundtrip_preserves_parse_and_format
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
        Dim source As Version = Version.Parse("10.20.30.40")
        Dim roundTrip As Version = Version.Parse(source.ToString())
        __Check(CStr(source.Major = roundTrip.Major), "True")
        __Check(CStr(source.Minor = roundTrip.Minor), "True")
        __Check(CStr(source.Build = roundTrip.Build), "True")
        __Check(CStr(source.Revision = roundTrip.Revision), "True")
    End Sub
End Module
