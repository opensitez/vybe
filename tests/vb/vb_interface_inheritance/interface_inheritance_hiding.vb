' vybe-test: vb/vb_interface_inheritance/interface_inheritance_hiding
' origin: languages/vb/tests/vb/test_vb_interface_inheritance.rs

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
    Function GetVal() As Integer
End Interface

Interface IDerived
    Inherits IBase
    Shadows Function GetVal() As Integer
End Interface

Class C
    Implements IDerived
    
    Private Function IBase_GetVal() As Integer Implements IBase.GetVal
        Return 1
    End Function
    
    Private Function IDerived_GetVal() As Integer Implements IDerived.GetVal
        Return 2
    End Function
End Class

Module M
    Sub Main()
        Dim obj As New C()
        Dim b As IBase = obj
        Dim d As IDerived = obj
        
        __Check(CStr(b.GetVal()), "1")
        __Check(CStr(d.GetVal()), "2")
    End Sub
End Module
