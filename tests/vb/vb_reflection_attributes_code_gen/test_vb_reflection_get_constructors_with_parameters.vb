' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_reflection_get_constructors_with_parameters
' origin: languages/vb/tests/vb/test_vb_reflection_attributes_code_gen.rs

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

Class Person
    Public Sub New(name As String, age As Integer)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ctor = GetType(Person).GetConstructors()(0)
        Dim params = ctor.GetParameters()
        __Check(CStr(params(0).Name & ":" & params(0).ParameterType.Name & "|" & params(1).Name & ":" & params(1).ParameterType.Name), "name:String|age:Int32")
    End Sub
End Module
