' vybe-test: vb/vb_anonymous_types_key/anonymous_type_key_properties
' origin: languages/vb/tests/vb/test_vb_anonymous_types_key.rs

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
        ' The 'Key' modifier makes the property read-only and includes it in Equals/GetHashCode
        Dim p1 = New With { Key .Id = 1, .Name = "Alice" }
        Dim p2 = New With { Key .Id = 1, .Name = "Bob" }
        Dim p3 = New With { Key .Id = 2, .Name = "Alice" }
        
        ' Equality depends only on Key properties
        __Check(CStr(p1.Equals(p2)), "True")
        __Check(CStr(p1.Equals(p3)), "False")
        
        __Check(CStr(p1.Id), "1")
    End Sub
End Module
