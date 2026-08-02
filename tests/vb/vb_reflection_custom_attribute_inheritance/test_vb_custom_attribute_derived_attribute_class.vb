' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_derived_attribute_class
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

Class BaseAttribute
    Inherits Attribute
    Public Message As String
    Public Sub New(msg As String) : Message = msg : End Sub
End Class

Class SpecializedAttribute
    Inherits BaseAttribute
    Public Sub New(msg As String) : MyBase.New("Spec: " & msg) : End Sub
End Class

<Specialized("CustomNote")>
Class TargetClass : End Class

Module Program
    Sub Main()
        Dim attr = CType(GetType(TargetClass).GetCustomAttributes(GetType(BaseAttribute), True)(0), BaseAttribute)
        __Check(CStr(attr.Message), "Spec: CustomNote")
    End Sub
End Module
