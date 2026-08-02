' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_self_referential_constraint
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface IComparableEntity(Of T As IComparableEntity(Of T))
    Function CompareTo(other As T) As Integer
End Interface

Class Account
    Implements IComparableEntity(Of Account)
    Public Property ID As Integer
    Public Function CompareTo(other As Account) As Integer Implements IComparableEntity(Of Account).CompareTo
        Return ID.CompareTo(other.ID)
    End Function
End Class

Module Program
    Sub Main()
        Dim a1 As New Account With {.ID = 10}
        Dim a2 As New Account With {.ID = 20}
        __Check(CStr(a1.CompareTo(a2) < 0), "True")
    End Sub
End Module
