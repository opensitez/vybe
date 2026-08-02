' vybe-test: vb/vb_guid_parse_format/guid_parse_with_braces_is_accepted
' origin: languages/vb/tests/vb/test_vb_guid_parse_format.rs

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

Module M
    Sub Main()
        Dim g As Guid = Guid.Parse("{d87a74a4-5694-4d8b-a3ed-3085794711f1}")
        __Check(CStr(g.ToString("N")), "d87a74a4-5694-4d8b-a3ed-3085794711f1")
    End Sub
End Module
