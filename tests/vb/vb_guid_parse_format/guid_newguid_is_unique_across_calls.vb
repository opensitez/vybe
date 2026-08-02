' vybe-test: vb/vb_guid_parse_format/guid_newguid_is_unique_across_calls
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
        Dim first As Guid = Guid.NewGuid()
        Dim second As Guid = Guid.NewGuid()
        __Check(CStr(first = second), "False")
        __Check(CStr(first <> Guid.Empty), "True")
        __Check(CStr(second <> Guid.Empty), "True")
    End Sub
End Module
