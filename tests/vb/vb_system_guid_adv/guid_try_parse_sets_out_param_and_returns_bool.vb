' vybe-test: vb/vb_system_guid_adv/guid_try_parse_sets_out_param_and_returns_bool
' origin: languages/vb/tests/vb/test_vb_system_guid_adv.rs

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
        Dim r As Guid
        __Check(CStr(Guid.TryParse("not-a-guid", r)), "False")
        __Check(CStr(Guid.TryParse("d87a74a4-5694-4d8b-a3ed-3085794711f1", r)), "True")
        __Check(CStr(r.ToString().StartsWith("d87a74a4")), "True")
    End Sub
End Module
