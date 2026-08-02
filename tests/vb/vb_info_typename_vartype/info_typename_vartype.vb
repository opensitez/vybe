' vybe-test: vb/vb_info_typename_vartype/info_typename_vartype
' origin: languages/vb/tests/vb/test_vb_info_typename_vartype.rs

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
        Dim s As String = "test"
        Dim i As Integer = 10
        
        ' TypeName returns a friendly string of the type
        __Check(CStr(TypeName(s)), "String")
        __Check(CStr(TypeName(i)), "Integer")
        
        ' VarType returns an enum value from Microsoft.VisualBasic.VariantType
        __Check(CStr(CInt(VarType(s))), "8") ' VariantType.String = 8
        __Check(CStr(CInt(VarType(i))), "3") ' VariantType.Integer = 3
    End Sub
End Module
