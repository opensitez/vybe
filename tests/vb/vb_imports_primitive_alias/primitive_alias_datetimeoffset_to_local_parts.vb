' vybe-test: vb/vb_imports_primitive_alias/primitive_alias_datetimeoffset_to_local_parts
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

Imports OffsetAlias = System.DateTimeOffset

Module M
    Sub Main()
        Dim value As OffsetAlias = OffsetAlias.Parse("2026-12-31T23:45:00+00:00")
        __Check(CStr(value.Year), "2026")
        __Check(CStr(value.Minute), "45")
    End Sub
End Module
