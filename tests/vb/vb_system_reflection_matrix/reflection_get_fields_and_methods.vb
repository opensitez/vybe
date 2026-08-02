' vybe-test: vb/vb_system_reflection_matrix/reflection_get_fields_and_methods
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
        Dim t As Type = GetType(Composite)
        __Check(CStr(t.GetFields().Length >= 1), "True")
        __Check(CStr(t.GetMethods().Length >= 1), "True")
    End Sub

    Class Composite
        Public X As Integer
        Public Y As Integer
        Public Sub Inc()
            X += 1
        End Sub
        Public Sub Dec()
            X -= 1
        End Sub
    End Class
End Module
