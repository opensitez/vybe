' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_order_repository_in_memory
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

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

Imports System.Collections.Generic
Imports System.Linq

Class OrderEntity
    Public Property OrderId As String
    Public Property CustomerName As String
End Class

Class OrderRepository
    Private db As New List(Of OrderEntity)()

    Public Sub Save(ord As OrderEntity)
        db.Add(ord)
    End Sub

    Public Function FindByCustomer(name As String) As List(Of OrderEntity)
        Return db.Where(Function(o) o.CustomerName = name).ToList()
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As New OrderRepository()
        repo.Save(New OrderEntity With {.OrderId = "O1", .CustomerName = "Alice"})
        repo.Save(New OrderEntity With {.OrderId = "O2", .CustomerName = "Alice"})
        repo.Save(New OrderEntity With {.OrderId = "O3", .CustomerName = "Bob"})

        __Check(CStr(repo.FindByCustomer("Alice").Count), "2")
    End Sub
End Module
