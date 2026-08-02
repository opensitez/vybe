' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_typeof_parameter
' origin: languages/vb/tests/vb/test_vb_reflection_custom_attribute_inheritance.rs

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

<AttributeUsage(AttributeTargets.Class)>
Class RelatedTypeAttribute
    Inherits Attribute
    Public TargetType As Type
    Public Sub New(t As Type) : TargetType = t : End Sub
End Class

<RelatedType(GetType(String))>
Class StringProcessor : End Class

Module Program
    Sub Main()
        Dim attr = CType(GetType(StringProcessor).GetCustomAttributes(GetType(RelatedTypeAttribute), False)(0), RelatedTypeAttribute)
        __Check(CStr(attr.TargetType.Name), "String")
    End Sub
End Module
