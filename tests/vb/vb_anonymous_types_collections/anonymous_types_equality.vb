' vybe-test: vb/vb_anonymous_types_collections/anonymous_types_equality
' origin: languages/vb/tests/vb/test_vb_anonymous_types_collections.rs

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
        ' Two anonymous types with the same Key properties are considered equal
        Dim a1 = New With { Key .X = 1, Key .Y = 2, .Z = 3 }
        Dim a2 = New With { Key .X = 1, Key .Y = 2, .Z = 4 }
        Dim a3 = New With { Key .X = 2, Key .Y = 2, .Z = 3 }
        
        __Check(CStr(a1.Equals(a2)), "True")
        __Check(CStr(a1.Equals(a3)), "False")
    End Sub
End Module
