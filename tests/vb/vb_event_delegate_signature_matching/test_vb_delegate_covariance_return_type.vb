' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_covariance_return_type
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

Delegate Function AnimalFactory() As Animal

Module Program
    Private Function CreateDog() As Dog
        Return New Dog()
    End Function

    Sub Main()
        Dim f As AnimalFactory = AddressOf CreateDog
        Dim a As Animal = f()
        __Check(CStr(a IsNot Nothing), "True")
    End Sub
End Module
