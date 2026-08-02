' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_generic_interface_inheritance_specialization
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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
    Sub Add(item As T)
End Interface

Interface ICustomerRepository
    Inherits IRepository(Of String)
    Function GetCustomerName(id As Integer) As String
End Interface

Class CustomerService
    Implements ICustomerRepository
    Private customer As String
    Public Sub Add(item As String) Implements IRepository(Of String).Add
        customer = item
    End Sub
    Public Function GetCustomerName(id As Integer) As String Implements ICustomerRepository.GetCustomerName
        Return customer
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As ICustomerRepository = New CustomerService()
        repo.Add("Alice")
        __Check(CStr(repo.GetCustomerName(1)), "Alice")
    End Sub
End Module
