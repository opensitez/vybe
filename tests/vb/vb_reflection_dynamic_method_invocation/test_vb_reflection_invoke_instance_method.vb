' vybe-test: vb/vb_reflection_dynamic_method_invocation/test_vb_reflection_invoke_instance_method
' origin: languages/vb/tests/vb/test_vb_reflection_dynamic_method_invocation.rs

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

Imports System.Reflection

Class Calculator
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As New Calculator()
        Dim t As Type = calc.GetType()
        Dim mi As MethodInfo = t.GetMethod("Add")
        Dim res As Object = mi.Invoke(calc, New Object() {5, 10})
        __Check(CStr(res), "15")
    End Sub
End Module
