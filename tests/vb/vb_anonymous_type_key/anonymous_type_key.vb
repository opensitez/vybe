' vybe-test: vb/vb_anonymous_type_key/anonymous_type_key
' origin: languages/vb/tests/vb/test_vb_anonymous_type_key.rs

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
        ' Anonymous type with Key properties (makes them read-only and participates in Equals/GetHashCode)
        Dim a1 = New With {Key .Id = 1, .Name = "A"}
        Dim a2 = New With {Key .Id = 1, .Name = "B"}
        
        __Check(CStr(a1.Id), "1")
        __Check(CStr(a1.Equals(a2)), "True") ' Should be true because only Key properties are compared
        
        a1.Name = "C" ' Non-key is mutable
        __Check(CStr(a1.Name), "C")
    End Sub
End Module
