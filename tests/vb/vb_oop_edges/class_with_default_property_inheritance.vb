' vybe-test: vb/vb_oop_edges/class_with_default_property_inheritance
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

Class Base
    Default Public Overridable Property Item(index As Integer) As String
        Get
            Return "Base" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class Derived
    Inherits Base
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        __Check(CStr(d(10)), "Base10")
    End Sub
End Module
