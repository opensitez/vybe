' vybe-test: vb/vb_custom_attribute_declaration_usage/test_vb_custom_attribute_read_properties
' origin: languages/vb/tests/vb/test_vb_custom_attribute_declaration_usage.rs

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

    Public ReadOnly Description As String
    Public Property Version As Integer = 1

    Public Sub New(desc As String)
        Me.Description = desc
    End Sub
End Class

<Info("Test Class", Version := 2)>
Class Sample
End Class

Module Program
    Sub Main()
        Dim t As Type = GetType(Sample)
        Dim attrs = t.GetCustomAttributes(GetType(InfoAttribute), False)
        If attrs.Length > 0 Then
            Dim info As InfoAttribute = CType(attrs(0), InfoAttribute)
            __Check(CStr(info.Description & ":V" & info.Version), "Test Class:V2")
        End If
    End Sub
End Module
