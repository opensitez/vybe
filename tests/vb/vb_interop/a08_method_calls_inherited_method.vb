' vybe-test: vb/vb_interop/a08_method_calls_inherited_method
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Animal
    Dim species As String
    Public Sub New(s As String)
        species = s
    End Sub
    Public Function GetSpecies() As String
        Return species
    End Function
End Class
Public Class Dog
    Inherits Animal
    Public Sub New()
        MyBase.New("Canine")
    End Sub
    Public Function Describe() As String
        Return "Dog: " & GetSpecies()
    End Function
End Class
Dim d As New Dog()
__Check(CStr(d.Describe()), "Dog: Canine")
