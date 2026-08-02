' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_get_custom_attributes
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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
Class RequiredAttribute
    Inherits Attribute
End Class

Class Model
    <Required>
    Public Property Name As String
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Model).GetProperty("Name")
        Dim attrs = prop.GetCustomAttributes(GetType(RequiredAttribute), False)
        __Check(CStr(attrs.Length), "1")
    End Sub
End Module
