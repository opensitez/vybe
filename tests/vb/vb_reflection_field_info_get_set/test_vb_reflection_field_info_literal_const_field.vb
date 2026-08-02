' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_literal_const_field
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

Class Constants
    Public Const MaxUsers As Integer = 100
End Class

Module Program
    Sub Main()
        Dim field = GetType(Constants).GetField("MaxUsers")
        __Check(CStr(field.IsLiteral & "|" & field.GetRawConstantValue()), "True|100")
    End Sub
End Module
