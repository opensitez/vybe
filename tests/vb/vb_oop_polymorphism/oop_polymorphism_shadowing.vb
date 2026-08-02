' vybe-test: vb/vb_oop_polymorphism/oop_polymorphism_shadowing
' origin: languages/vb/tests/vb/test_vb_oop_polymorphism.rs

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
    Public Overridable Function GetName() As String
        Return "Base"
    End Function
End Class

Class Derived1
    Inherits Base
    Public Overrides Function GetName() As String
        Return "Derived1"
    End Function
End Class

Class Derived2
    Inherits Base
    ' Shadows the base method, doesn't override
    Public Shadows Function GetName() As String
        Return "Derived2"
    End Function
End Class

Module M
    Sub Main()
        Dim d1 As New Derived1()
        Dim d2 As New Derived2()
        
        Dim b1 As Base = d1
        Dim b2 As Base = d2
        
        __Check(CStr(b1.GetName()), "Derived1")
        __Check(CStr(b2.GetName()), "Base")
        __Check(CStr(d2.GetName()), "Derived2")
    End Sub
End Module
