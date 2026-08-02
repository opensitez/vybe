' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_custom_attribute
' origin: languages/vb/tests/vb/test_vb_reflection_field_info_get_set.rs

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

<AttributeUsage(AttributeTargets.Field)>
Class RangeAttribute
    Inherits Attribute
    Public Min As Integer
    Public Max As Integer
    Public Sub New(min As Integer, max As Integer) : Me.Min = min : Me.Max = max : End Sub
End Class

Class Form
    <Range(1, 100)>
    Public Percentage As Integer
End Class

Module Program
    Sub Main()
        Dim field = GetType(Form).GetField("Percentage")
        Dim attr = CType(field.GetCustomAttributes(GetType(RangeAttribute), False)(0), RangeAttribute)
        __Check(CStr(attr.Min & " To " & attr.Max), "1 To 100")
    End Sub
End Module
