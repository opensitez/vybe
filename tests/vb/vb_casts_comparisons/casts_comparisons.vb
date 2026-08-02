' vybe-test: vb/vb_casts_comparisons/casts_comparisons
' origin: languages/vb/tests/vb/test_vb_casts_comparisons.rs

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
        
        ' DirectCast requires an inheritance or implementation relationship
        Dim d1 As Dog = DirectCast(a, Dog)
        d1.Bark()
        
        ' TryCast returns Nothing if the cast fails (only for reference types)
        Dim a2 As New Animal()
        Dim d2 As Dog = TryCast(a2, Dog)
        If d2 Is Nothing Then
            __Check(CStr("Cast Failed"), "Cast Failed")
        End If
        
        ' CType can do conversions as well as casts (e.g. String to Integer)
        Dim numStr As Object = "123"
        Dim num As Integer = CType(numStr, Integer)
        __Check(CStr(num + 1), "124")
    End Sub
End Module
