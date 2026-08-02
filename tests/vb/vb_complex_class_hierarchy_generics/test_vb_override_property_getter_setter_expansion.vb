' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_override_property_getter_setter_expansion
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Class BaseProperty
    Public Overridable Property Value As Integer = 10
End Class

Class LoggedProperty
    Inherits BaseProperty

    Public Overrides Property Value As Integer
        Get
            Return MyBase.Value
        End Get
        Set(v As Integer)
            __Check(CStr("Setting Value to " & v), "Setting Value to 42")
            MyBase.Value = v
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim p As BaseProperty = New LoggedProperty()
        p.Value = 42
        __Check(CStr(p.Value), "42")
    End Sub
End Module
