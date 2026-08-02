' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_abstract_mustoverride_protected
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

MustInherit Class AbstractModel
    Protected MustOverride Property Code As Integer
    Public Function ReadCode() As Integer
        Return Code
    End Function
End Class

Class ConcreteModel
    Inherits AbstractModel
    Protected Overrides Property Code As Integer = 999
End Class

Module Program
    Sub Main()
        Dim m As AbstractModel = New ConcreteModel()
        __Check(CStr(m.ReadCode()), "999")
    End Sub
End Module
