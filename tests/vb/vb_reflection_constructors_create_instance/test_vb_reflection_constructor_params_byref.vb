' vybe-test: vb/vb_reflection_constructors_create_instance/test_vb_reflection_constructor_params_byref
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

Class ByRefCtor
    Public Sub New(ByRef val As Integer)
        val = 999
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = GetType(ByRefCtor)
        Dim ctor = t.GetConstructors()(0)
        Dim args As Object() = {10}
        ctor.Invoke(args)
        __Check(CStr(args(0)), "999")
    End Sub
End Module
