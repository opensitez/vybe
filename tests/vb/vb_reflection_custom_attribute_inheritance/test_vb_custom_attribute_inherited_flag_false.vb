' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_inherited_flag_false
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

<AttributeUsage(AttributeTargets.Class, Inherited:=False)>
Class NonInheritedAttribute
    Inherits Attribute
End Class

<NonInherited>
Class BaseClass : End Class

Class DerivedClass : Inherits BaseClass : End Class

Module Program
    Sub Main()
        Dim attrs = GetType(DerivedClass).GetCustomAttributes(GetType(NonInheritedAttribute), True)
        __Check(CStr(attrs.Length), "0")
    End Sub
End Module
