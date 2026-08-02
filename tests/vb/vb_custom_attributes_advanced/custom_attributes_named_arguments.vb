' vybe-test: vb/vb_custom_attributes_advanced/custom_attributes_named_arguments
' origin: languages/vb/tests/vb/test_vb_custom_attributes_advanced.rs

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

<AttributeUsage(AttributeTargets.Class Or AttributeTargets.Method)>
Public Class InfoAttribute
    Inherits Attribute
    
    Public Property Author As String
    Public Property Version As String
    
    Public Sub New()
    End Sub
End Class

<Info(Author:="Alice", Version:="1.0")>
Class MyComponent
End Class

Module M
    Sub Main()
        Dim t As Type = GetType(MyComponent)
        Dim attrs() As Object = t.GetCustomAttributes(GetType(InfoAttribute), False)
        Dim info As InfoAttribute = DirectCast(attrs(0), InfoAttribute)
        
        __Check(CStr(info.Author), "Alice")
        __Check(CStr(info.Version), "1.0")
    End Sub
End Module
