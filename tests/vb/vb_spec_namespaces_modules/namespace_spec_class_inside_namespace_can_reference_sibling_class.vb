' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_class_inside_namespace_can_reference_sibling_class
' origin: languages/vb/tests/vb/test_vb_spec_namespaces_modules.rs

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

Namespace Demo
    Public Class NameBox
        Public Shared Function Value() As Integer
            Return Counter.NextValue()
        End Function
    End Class
    Public Class Counter
        Private Shared currentValue As Integer = 0
        Public Shared Function NextValue() As Integer
            currentValue += 1
            Return currentValue
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        __Check(CStr(Demo.NameBox.Value()), "1")
        __Check(CStr(Demo.NameBox.Value()), "2")
    End Sub
End Module
