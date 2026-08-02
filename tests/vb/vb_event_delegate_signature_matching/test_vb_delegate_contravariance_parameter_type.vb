' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_contravariance_parameter_type
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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

Imports System

Class Animal : End Class
Class Dog : Inherits Animal : End Class

Delegate Sub DogHandler(d As Dog)

Module Program
    Private Sub ProcessAnimal(a As Animal)
        __Check(CStr("Processed Animal: " & a.GetType().Name), "Processed Animal: Dog")
    End Sub

    Sub Main()
        Dim h As DogHandler = AddressOf ProcessAnimal
        h(New Dog())
    End Sub
End Module
