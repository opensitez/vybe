' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_reflection_property_value_retrieval
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

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
    Private Function GetPropValue(obj As Object, propName As String) As Object
        Dim prop = obj.GetType().GetProperty(propName)
        Return prop.GetValue(obj)
    End Function

    Sub Main()
        Dim anon = New With {.Title = "ReflectedTitle"}
        __Check(CStr(GetPropValue(anon, "Title")), "ReflectedTitle")
    End Sub
End Module
