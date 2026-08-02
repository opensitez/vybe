' vybe-test: vb/vb_version_class_comparison/test_vb_version_comparison_operators
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

Module Program
    Sub Main()
        Dim v1 = New Version(1, 0, 0)
        Dim v2 = New Version(1, 1, 0)
        Dim v3 = New Version(2, 0, 0)
        __Check(CStr((v1 < v2) & "|" & (v2 < v3) & "|" & (v1 = New Version(1, 0, 0))), "True|True|True")
    End Sub
End Module
