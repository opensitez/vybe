' vybe-test: vb/vb_system_reflection_matrix/reflection_binding_flags_controls_visibility
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
        Dim t As Type = GetType(Container)
        Dim fields = t.GetFields(BindingFlags.Instance Or BindingFlags.NonPublic)
        Dim p As PropertyInfo = t.GetProperty("Visible", BindingFlags.Instance Or BindingFlags.Public)
        __Check(CStr(fields.Length), "1")
        __Check(CStr(p.Name), "Visible")
    End Sub

    Class Container
        Private Secret As Integer = 7
        Public Property Visible As Integer
            Get
                Return Secret
            End Get
            Set
                Secret = Value
            End Set
        End Property
    End Class
End Module
