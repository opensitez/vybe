' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notinheritable_attribute_usage
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

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
NotInheritable Class CustomInfoAttribute
    Inherits Attribute
    Public Property Description As String
    Public Sub New(desc As String)
        Description = desc
    End Sub
End Class

<CustomInfo("TestClass")>
Class Target
End Class

Module Program
    Sub Main()
        Dim attrs = GetType(Target).GetCustomAttributes(GetType(CustomInfoAttribute), False)
        Dim info = CType(attrs(0), CustomInfoAttribute)
        __Check(CStr(info.Description), "TestClass")
    End Sub
End Module
