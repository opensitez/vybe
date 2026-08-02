' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_null_property_values
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
        Dim obj = New With {.Text = CType(Nothing, String), .Val = CType(Nothing, Nullable(Of Integer))}
        __Check(CStr((obj.Text Is Nothing) & "|" & (Not obj.Val.HasValue)), "True|True")
    End Sub
End Module
