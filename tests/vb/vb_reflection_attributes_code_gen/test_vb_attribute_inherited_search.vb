' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_attribute_inherited_search
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

<AttributeUsage(AttributeTargets.Class, Inherited:=True)>
Class BaseMetadataAttribute
    Inherits Attribute
    Public Description As String
    Public Sub New(d As String)
        Description = d
    End Sub
End Class

<BaseMetadata("Base Description")>
Class BaseClass
End Class

Class SubClass
    Inherits BaseClass
End Class

Module Program
    Sub Main()
        ' Inherited search enabled (inherit:=True)
        Dim attr = CType(Attribute.GetCustomAttribute(GetType(SubClass), GetType(BaseMetadataAttribute), inherit:=True), BaseMetadataAttribute)
        __Check(CStr(attr.Description), "Base Description")
    End Sub
End Module
