' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_order_builder_fluent_interface
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

Class FluentOrderBuilder
    Private custName As String
    Private items As New List(Of String)()

    Public Function WithCustomer(name As String) As FluentOrderBuilder
        custName = name
        Return Me
    End Function

    Public Function AddItem(item As String) As FluentOrderBuilder
        items.Add(item)
        Return Me
    End Function

    Public Function Build() As String
        Return custName & ":" & String.Join(",", items)
    End Function
End Class

Module Program
    Sub Main()
        Dim summary = New FluentOrderBuilder().WithCustomer("Charlie").AddItem("Laptop").AddItem("Mouse").Build()
        __Check(CStr(summary), "Charlie:Laptop,Mouse")
    End Sub
End Module
