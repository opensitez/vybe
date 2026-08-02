' vybe-test: vb/vb_system_linq_join_matrix/linq_join_pairs_on_ids
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
    Public Total As Integer

    Public Sub New(userId As Integer, total As Integer)
        Me.UserId = userId
        Me.Total = total
    End Sub
End Class

Module M
    Sub Main()
        Dim users = {New User(1, "Ada"), New User(2, "Bob")}
        Dim orders = {New Order(1, 100), New Order(1, 200), New Order(2, 30)}

        Dim joined = From u In users _
            Join o In orders On u.Id Equals o.UserId _
            Select Name = u.Name, Total = o.Total

        Dim sum = joined.Sum(Function(x) x.Total)
        Dim hasBob As Boolean = joined.Any(Function(x) x.Name = "Bob")

        __Check(CStr(sum), "330")
        __Check(CStr(hasBob), "True")
        __Check(CStr(joined.Count()), "3")
    End Sub
End Module
