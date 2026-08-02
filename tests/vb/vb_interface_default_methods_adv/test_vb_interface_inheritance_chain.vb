' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_inheritance_chain
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IBase
    Sub MethodA()
End Interface

Interface IDerived
    Inherits IBase
    Sub MethodB()
End Interface

Class Implementation
    Implements IDerived
    Public Sub MethodA() Implements IBase.MethodA
        __Check(CStr("MethodA"), "MethodA")
    End Sub
    Public Sub MethodB() Implements IDerived.MethodB
        __Check(CStr("MethodB"), "MethodB")
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As IDerived = New Implementation()
        d.MethodA()
        d.MethodB()
    End Sub
End Module
