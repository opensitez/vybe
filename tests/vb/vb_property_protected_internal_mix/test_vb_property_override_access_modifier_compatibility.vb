' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_override_access_modifier_compatibility
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Class BaseClass
    Public Overridable Property Message As String
        Get
            Return "BaseMsg"
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class DerivedClass
    Inherits BaseClass
    Public Overrides Property Message As String
        Get
            Return "DerivedMsg"
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim b As BaseClass = New DerivedClass()
        __Check(CStr(b.Message), "DerivedMsg")
    End Sub
End Module
