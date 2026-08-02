' vybe-test: vb/vb_myclass_me_mybase/myclass_me_mybase
' origin: languages/vb/tests/vb/test_vb_myclass_me_mybase.rs

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
    Public Overridable Sub Print()
        __Check(CStr("Base"), "MoreDerived")
    End Sub
End Class

Class Derived
    Inherits Base
    
    Public Overrides Sub Print()
        __Check(CStr("Derived"), "Derived")
    End Sub
    
    Public Sub Test()
        ' Me calls the most derived override (Derived)
        Me.Print()
        
        ' MyClass calls the implementation in the current class, ignoring overrides
        ' Wait, MyClass calls the method as if it were not virtual, 
        ' so MyClass.Print() in Derived calls Derived.Print(), but if Derived were inherited and Print overridden again, MyClass.Print() in Derived would still call Derived.Print().
        MyClass.Print()
        
        ' MyBase calls the base class implementation
        MyBase.Print()
    End Sub
End Class

Class MoreDerived
    Inherits Derived
    
    Public Overrides Sub Print()
        __Check(CStr("MoreDerived"), "Base")
    End Sub
End Class

Module M
    Sub Main()
        Dim obj As New MoreDerived()
        obj.Test()
    End Sub
End Module
