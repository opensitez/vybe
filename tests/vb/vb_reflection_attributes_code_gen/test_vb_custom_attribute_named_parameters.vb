' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_custom_attribute_named_parameters
' origin: languages/vb/tests/vb/test_vb_reflection_attributes_code_gen.rs

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

Imports System

<AttributeUsage(AttributeTargets.Property)>
Class ColumnAttribute
    Inherits Attribute
    Public Property Name As String
    Public Property IsKey As Boolean
End Class

Class UserRecord
    <Column(Name:="user_id", IsKey:=True)>
    Public Property UserId As Integer
End Class

Module Program
    Sub Main()
        Dim p = GetType(UserRecord).GetProperty("UserId")
        Dim attr = CType(Attribute.GetCustomAttribute(p, GetType(ColumnAttribute)), ColumnAttribute)
        __Check(CStr(attr.Name & "|" & attr.IsKey), "user_id|True")
    End Sub
End Module
