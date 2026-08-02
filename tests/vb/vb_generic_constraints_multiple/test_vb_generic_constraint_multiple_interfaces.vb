' vybe-test: vb/vb_generic_constraints_multiple/test_vb_generic_constraint_multiple_interfaces
' origin: languages/vb/tests/vb/test_vb_generic_constraints_multiple.rs

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
Imports System.Collections

Interface IIdentifiable
    ReadOnly Property Id As Integer
End Interface

Class EntityContainer(Of T As {IIdentifiable, IComparable})
    Public Sub Process(item As T)
        __Check(CStr("Id: " & item.Id), "Id: 100")
    End Sub
End Class

Class Product
    Implements IIdentifiable, IComparable
    Public ReadOnly Property Id As Integer Implements IIdentifiable.Id
        Get
            Return 100
        End Get
    End Property

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
        Return 0
    End Function
End Class

Module Program
    Sub Main()
        Dim container As New EntityContainer(Of Product)()
        container.Process(New Product())
    End Sub
End Module
