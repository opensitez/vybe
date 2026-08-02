' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_generic_interface
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IRepository(Of T)
    Sub Add(entity As T)
    Function GetById(id As Integer) As T
End Interface

Class ProductRepository
    Implements IRepository(Of String)
    Private item As String = ""
    Public Sub Add(entity As String) Implements IRepository(Of String).Add
        item = entity
    End Sub
    Public Function GetById(id As Integer) As String Implements IRepository(Of String).GetById
        Return item
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As IRepository(Of String) = New ProductRepository()
        repo.Add("Laptop")
        __Check(CStr(repo.GetById(1)), "Laptop")
    End Sub
End Module
