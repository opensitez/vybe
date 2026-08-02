' vybe-test: vb/vb_method_shadows_adv/method_shadows_adv
' origin: languages/vb/tests/vb/test_vb_method_shadows_adv.rs

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
    Public Sub Process()
        __Check(CStr("Base"), "Derived")
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' Shadows hides the base method instead of overriding
    Public Shadows Sub Process()
        __Check(CStr("Derived"), "Base")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        Dim b As Base = d
        
        d.Process()
        b.Process()
    End Sub
End Module
