' vybe-test: vb/vb_abstract_class_inheritance_chain/test_vb_mustinherit_abstract_property
' origin: languages/vb/tests/vb/test_vb_abstract_class_inheritance_chain.rs

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

MustInherit Class Component
    Public MustOverride Property Name As String
End Class

Class Button
    Inherits Component
    Private _name As String
    Public Overrides Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim btn As Component = New Button()
        btn.Name = "Submit"
        __Check(CStr(btn.Name), "Submit")
    End Sub
End Module
