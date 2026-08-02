' vybe-test: vb/vb_oop_interfaces/interface_implementing_inherited_interface_methods
' origin: languages/vb/tests/vb/test_vb_oop_interfaces.rs

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
Sub Base()
End Interface
Interface IDerived
Inherits IBase
Sub Derived()
End Interface
Class C
Implements IDerived
Public Sub Base() Implements IDerived.Base
End Sub
Public Sub Derived() Implements IDerived.Derived
__Check(CStr("D"), "D")
End Sub
End Class
Module M
Sub Main()
Dim c1 As IDerived = New C()
c1.Derived()
End Sub
End Module
