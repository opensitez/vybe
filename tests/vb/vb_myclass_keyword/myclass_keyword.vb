' vybe-test: vb/vb_myclass_keyword/myclass_keyword
' origin: languages/vb/tests/vb/test_vb_myclass_keyword.rs

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
    Public Overridable Sub Show()
        __Check(CStr("Base"), "Derived")
    End Sub
    
    Public Sub CallShow()
        ' Me.Show() calls the overridden version in Derived
        Me.Show()
        ' MyClass.Show() statically calls the version defined in this class, bypassing overrides
        MyClass.Show()
    End Sub
End Class

Class Derived
    Inherits Base
    Public Overrides Sub Show()
        __Check(CStr("Derived"), "Base")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.CallShow()
    End Sub
End Module
