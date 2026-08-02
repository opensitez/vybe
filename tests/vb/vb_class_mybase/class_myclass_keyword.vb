' vybe-test: vb/vb_class_mybase/class_myclass_keyword
' origin: languages/vb/tests/vb/test_vb_class_mybase.rs

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

Class BasePrinter
    Public Overridable Function GetName() As String
        Return "Base"
    End Function
    
    Public Function PrintName() As String
        ' MyClass forces call to this class's implementation, ignoring overrides
        Return MyClass.GetName()
    End Function
    
    Public Function PrintNamePolymorphic() As String
        Return Me.GetName()
    End Function
End Class

Class DerivedPrinter
    Inherits BasePrinter
    
    Public Overrides Function GetName() As String
        Return "Derived"
    End Function
End Class

Module M
    Sub Main()
        Dim d As New DerivedPrinter()
        __Check(CStr(d.PrintName()), "Base")
        __Check(CStr(d.PrintNamePolymorphic()), "Derived")
    End Sub
End Module
