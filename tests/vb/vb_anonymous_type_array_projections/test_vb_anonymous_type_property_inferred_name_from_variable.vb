' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_property_inferred_name_from_variable
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
        Dim title As String = "Manager"
        Dim level As Integer = 3
        ' Inferred property names .title and .level
        Dim obj = New With {title, level}
        __Check(CStr(obj.title & ":" & obj.level), "Manager:3")
    End Sub
End Module
