' vybe-test: vb/vb_trycast_reference_types/test_vb_trycast_compatible_reference_type_succeeds
' origin: languages/vb/tests/vb/test_vb_trycast_reference_types.rs

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
    Public ReadOnly Name As String = "Rover"
End Class

Module Program
    Sub Main()
        Dim a As Animal = New Dog()
        Dim d As Dog = TryCast(a, Dog)
        __Check(CStr(d IsNot Nothing AndAlso d.Name = "Rover"), "True")
    End Sub
End Module
