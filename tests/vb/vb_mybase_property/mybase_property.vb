' vybe-test: vb/vb_mybase_property/mybase_property
' origin: languages/vb/tests/vb/test_vb_mybase_property.rs

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

Class Base
    Public Overridable ReadOnly Property Name As String
        Get
            Return "Base"
        End Get
    End Property
End Class

Class Derived
    Inherits Base
    
    Public Overrides ReadOnly Property Name As String
        Get
            Return MyBase.Name & "Derived"
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        __Check(CStr(d.Name), "BaseDerived")
    End Sub
End Module
