' vybe-test: vb/vb_system_guid_matrix/guid_brace_and_parentheses_forms_parse
' origin: languages/vb/tests/vb/test_vb_system_guid_matrix.rs

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
        Dim a As Guid = Guid.Parse("{4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4}")
        Dim b As Guid = Guid.Parse("(4f7f1dcb-7a39-4f9b-b0de-f7f9a2f5f8f4)")
        __Check(CStr(a = b), "True")
    End Sub
End Module
