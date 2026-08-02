' vybe-test: vb/vb_imports_primitive_alias/primitive_alias_guid_to_string_length
' origin: languages/vb/tests/vb/test_vb_imports_primitive_alias.rs

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

Imports GuidAlias = System.Guid

Module M
    Sub Main()
        Dim value As GuidAlias = GuidAlias.Parse("3F2504E0-4F89-11D3-9A0C-0305E82C3301")
        __Check(CStr(value.ToString().Length), "36")
        __Check(CStr(value.GetType().Name), "Guid")
    End Sub
End Module
