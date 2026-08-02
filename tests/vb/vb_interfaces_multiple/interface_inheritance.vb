' vybe-test: vb/vb_interfaces_multiple/interface_inheritance
' origin: languages/vb/tests/vb/test_vb_interfaces_multiple.rs

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

Class MyClass
    Implements IDerived
    
    Public Sub MethodA() Implements IDerived.MethodA
        __Check(CStr("A"), "A")
    End Sub
    
    Public Sub MethodB() Implements IDerived.MethodB
        __Check(CStr("B"), "B")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As IDerived = New MyClass()
        d.MethodA()
        d.MethodB()
    End Sub
End Module
