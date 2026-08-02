' vybe-test: vb/vb_reflection_custom_attribute_inheritance/test_vb_custom_attribute_enum_parameter
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

Enum LogLevel
    Debug
    Info
    ErrorVal
End Enum

<AttributeUsage(AttributeTargets.Method)>
Class LogAttribute
    Inherits Attribute
    Public Level As LogLevel
    Public Sub New(l As LogLevel) : Level = l : End Sub
End Class

Class Service
    <Log(LogLevel.ErrorVal)>
    Public Sub Process() : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Service).GetMethod("Process")
        Dim attr = CType(m.GetCustomAttributes(GetType(LogAttribute), False)(0), LogAttribute)
        __Check(CStr(attr.Level.ToString()), "ErrorVal")
    End Sub
End Module
