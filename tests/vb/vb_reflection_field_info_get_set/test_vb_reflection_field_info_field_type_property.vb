' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_field_type_property
' origin: languages/vb/tests/vb/test_vb_reflection_field_info_get_set.rs

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

Class Entity
    Public Title As String
    Public Age As Integer
End Class

Module Program
    Sub Main()
        Dim f1 = GetType(Entity).GetField("Title")
        Dim f2 = GetType(Entity).GetField("Age")
        __Check(CStr(f1.FieldType.Name & "|" & f2.FieldType.Name), "String|Int32")
    End Sub
End Module
