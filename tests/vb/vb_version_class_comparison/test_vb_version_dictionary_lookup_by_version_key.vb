' vybe-test: vb/vb_version_class_comparison/test_vb_version_dictionary_lookup_by_version_key
' origin: languages/vb/tests/vb/test_vb_version_class_comparison.rs

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
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Version, String)()
        dict(New Version(1, 0, 0)) = "Release 1.0"
        dict(New Version(2, 0, 0)) = "Release 2.0"
        __Check(CStr(dict(New Version(1, 0, 0))), "Release 1.0")
    End Sub
End Module
