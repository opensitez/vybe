' vybe-test: vb/vb_casts_advanced/casts_trycast
' origin: languages/vb/tests/vb/test_vb_casts_advanced.rs

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

Class Animal
End Class

Class Dog
    Inherits Animal
    Public Sub Bark()
        __Check(CStr("Woof"), "Woof")
    End Sub
End Class

Module M
    Sub Main()
        Dim a1 As Animal = New Dog()
        Dim a2 As Animal = New Animal()
        
        ' TryCast returns Nothing if the cast fails (only for reference types)
        Dim d1 As Dog = TryCast(a1, Dog)
        Dim d2 As Dog = TryCast(a2, Dog)
        
        If d1 IsNot Nothing Then d1.Bark()
        __Check(CStr(d2 Is Nothing), "True")
    End Sub
End Module
