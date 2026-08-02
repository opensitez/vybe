' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_applied_to_parameter
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

<AttributeUsage(AttributeTargets.Parameter)>
Class NonNullAttribute
    Inherits Attribute
End Class

Class Validator
    Public Sub Process(<NonNull> data As String) : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Validator).GetMethod("Process")
        Dim p = m.GetParameters()(0)
        Dim isDefined = p.IsDefined(GetType(NonNullAttribute), False)
        __Check(CStr(isDefined), "True")
    End Sub
End Module
