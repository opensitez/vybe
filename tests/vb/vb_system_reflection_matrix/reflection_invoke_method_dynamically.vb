' vybe-test: vb/vb_system_reflection_matrix/reflection_invoke_method_dynamically
' origin: languages/vb/tests/vb/test_vb_system_reflection_matrix.rs

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

Module M
    Sub Main()
        Dim t As Type = GetType(Operation)
        Dim obj As New Operation()
        Dim method As MethodInfo = t.GetMethod("Add")
        Dim result As Object = method.Invoke(obj, New Object(){3, 4})
        __Check(CStr(result), "7")
    End Sub

    Class Operation
        Public Function Add(a As Integer, b As Integer) As Integer
            Return a + b
        End Function
    End Class
End Module
