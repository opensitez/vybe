' vybe-test: vb/vb_reflection_constructors_create_instance/test_vb_reflection_constructor_info_get_parameters
' origin: languages/vb/tests/vb/test_vb_reflection_constructors_create_instance.rs

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
Imports System.Reflection

Class Order
    Public Sub New(id As Integer, title As String) : End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(Order)
        Dim ctor = t.GetConstructors()(0)
        Dim params = ctor.GetParameters()
        __Check(CStr(params(0).Name & ":" & params(0).ParameterType.Name & "|" & params(1).Name & ":" & params(1).ParameterType.Name), "id:Int32|title:String")
    End Sub
End Module
