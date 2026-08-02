' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_named_positional_arguments
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
Class MetadataAttribute
    Inherits Attribute
    Public ReadOnly ID As Integer
    Public Property Description As String
    Public Property Version As Integer = 1
    Public Sub New(idVal As Integer) : ID = idVal : End Sub
End Class

<Metadata(100, Description:="ServiceClass", Version:=2)>
Class Service : End Class

Module Program
    Sub Main()
        Dim attr = CType(GetType(Service).GetCustomAttributes(GetType(MetadataAttribute), False)(0), MetadataAttribute)
        __Check(CStr(attr.ID & "|" & attr.Description & "|v" & attr.Version), "100|ServiceClass|v2")
    End Sub
End Module
