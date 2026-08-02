' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_key_properties_equals_comparison
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
        Dim o1 = New With {Key .ID = 1, .Name = "A"}
        Dim o2 = New With {Key .ID = 1, .Name = "B"} ' Non-key Name ignored in Equals
        Dim o3 = New With {Key .ID = 2, .Name = "A"}
        __Check(CStr(o1.Equals(o2) & "|" & o1.Equals(o3)), "True|False")
    End Sub
End Module
