' vybe-test: vb/vb_system_linq_join_matrix/linq_left_join_style_with_default_if_empty
' origin: languages/vb/tests/vb/test_vb_system_linq_join_matrix.rs

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

Class User
    Public Id As Integer
    Public Name As String
    Public Sub New(id As Integer, name As String)
        Me.Id = id
        Me.Name = name
    End Sub
End Class

Class Order
    Public UserId As Integer
    Public Amount As Integer
    Public Sub New(id As Integer, amount As Integer)
        Me.UserId = id
        Me.Amount = amount
    End Sub
End Class

Module M
    Sub Main()
        Dim users = {New User(1, "Ada"), New User(2, "Bob")}
        Dim orders = {New Order(1, 10)}

        Dim grouped = From u In users _
            Group Join o In orders On u.Id Equals o.UserId _
            Into g = Group _
            Select User = u.Name, Total = g.Sum(Function(x) x.Amount)

        Dim rows As Integer = grouped.Count()
        Dim firstTotal As Integer = grouped(0).Total
        Dim secondTotal As Integer = grouped(1).Total

        __Check(CStr(rows), "2")
        __Check(CStr(firstTotal), "10")
        __Check(CStr(secondTotal), "0")
    End Sub
End Module
