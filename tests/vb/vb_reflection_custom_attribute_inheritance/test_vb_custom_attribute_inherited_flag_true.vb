' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_inherited_flag_true
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

<AttributeUsage(AttributeTargets.Class, Inherited:=True)>
Class CategoryAttribute
    Inherits Attribute
    Public Tag As String
    Public Sub New(t As String) : Tag = t : End Sub
End Class

<Category("BaseTag")>
Class BaseClass : End Class

Class DerivedClass : Inherits BaseClass : End Class

Module Program
    Sub Main()
        Dim t = GetType(DerivedClass)
        Dim attrs = t.GetCustomAttributes(GetType(CategoryAttribute), True)
        __Check(CStr(attrs.Length & ":" & CType(attrs(0), CategoryAttribute).Tag), "1:BaseTag")
    End Sub
End Module
