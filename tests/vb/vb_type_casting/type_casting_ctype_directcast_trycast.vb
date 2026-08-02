' vybe-test: vb/vb_type_casting/type_casting_ctype_directcast_trycast
' origin: languages/vb/tests/vb/test_vb_type_casting.rs

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
End Class

Module M
    Sub Main()
        Dim d As New Dog()
        Dim a As Animal = d
        
        ' DirectCast
        Dim d2 As Dog = DirectCast(a, Dog)
        __Check(CStr(d2 IsNot Nothing), "True")
        
        ' TryCast
        Dim d3 As Dog = TryCast(a, Dog)
        __Check(CStr(d3 IsNot Nothing), "True")
        
        ' CType
        Dim d4 As Dog = CType(a, Dog)
        __Check(CStr(d4 IsNot Nothing), "True")
    End Sub
End Module
