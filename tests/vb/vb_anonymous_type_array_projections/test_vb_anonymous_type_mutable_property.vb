' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_mutable_property
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

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

Module Program
    Sub Main()
        ' In VB, non-Key properties of anonymous types are mutable!
        Dim item = New With {.Price = 10.0}
        item.Price = 15.5
        __Check(CStr(item.Price), "15.5")
    End Sub
End Module
