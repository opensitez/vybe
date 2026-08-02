' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_attribute_multiple_allow_multiple
' origin: languages/vb/tests/vb/test_vb_reflection_attributes_code_gen.rs

Imports System

<AttributeUsage(AttributeTargets.Class, AllowMultiple:=True)>
Class TagAttribute
    Inherits Attribute
    Public Tag As String
    Public Sub New(t As String)
        Tag = t
    End Sub
End Class

<Tag("V1")>
<Tag("Beta")>
Class Feature
End Class

Module Program
    Sub Main()
        Dim attrs = Attribute.GetCustomAttributes(GetType(Feature), GetType(TagAttribute))
        Dim tags As New System.Collections.Generic.List(Of String)()
        For Each a In attrs
            tags.Add(CType(a, TagAttribute).Tag)
        Next
        Console.WriteLine(String.Join(",", tags))
    End Sub
End Module
