' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_get_custom_attributes_data
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
Class TagAttribute
    Inherits Attribute
    Public Sub New(tag As String) : End Sub
End Class

<Tag("SampleTag")>
Class AnnotatedClass : End Class

Module Program
    Sub Main()
        Dim customData = GetType(AnnotatedClass).GetCustomAttributesData()
        __Check(CStr(customData.Count & ":" & customData(0).ConstructorArguments(0).Value), "1:SampleTag")
    End Sub
End Module
