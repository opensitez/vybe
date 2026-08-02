' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_allow_multiple_true
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

<AttributeUsage(AttributeTargets.Class, AllowMultiple:=True)>
Class TagAttribute
    Inherits Attribute
    Public Tag As String
    Public Sub New(t As String) : Tag = t : End Sub
End Class

<Tag("Alpha")>
<Tag("Beta")>
Class Item : End Class

Module Program
    Sub Main()
        Dim attrs = GetType(Item).GetCustomAttributes(GetType(TagAttribute), False)
        __Check(CStr(attrs.Length), "2")
    End Sub
End Module
