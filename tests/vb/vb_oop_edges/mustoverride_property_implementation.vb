' vybe-test: vb/vb_oop_edges/mustoverride_property_implementation
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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

MustInherit Class Base
    Public MustOverride Property Value As Integer
End Class

Class Derived
    Inherits Base
    
    Private _v As Integer
    Public Overrides Property Value As Integer
        Get
            Return _v
        End Get
        Set(v As Integer)
            _v = v
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.Value = 100
        __Check(CStr(d.Value), "100")
    End Sub
End Module
