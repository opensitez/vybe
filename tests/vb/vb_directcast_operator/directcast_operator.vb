' vybe-test: vb/vb_directcast_operator/directcast_operator
' origin: languages/vb/tests/vb/test_vb_directcast_operator.rs

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
        Dim a As Animal = New Dog()
        
        ' DirectCast requires the run-time type of an object variable to be the same as the specified type.
        ' It is faster than CType but throws an exception if the cast fails.
        Dim d As Dog = DirectCast(a, Dog)
        d.Bark()
    End Sub
End Module
