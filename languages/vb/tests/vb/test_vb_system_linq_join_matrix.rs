use super::helpers::run_vb;

#[test]
fn linq_join_pairs_on_ids() {
    let out = run_vb(
        r#"
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

        Console.WriteLine(sum)
        Console.WriteLine(hasBob)
        Console.WriteLine(joined.Count())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["330", "True", "3"]);
}

#[test]
fn linq_left_join_style_with_default_if_empty() {
    let out = run_vb(
        r#"
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

        Console.WriteLine(rows)
        Console.WriteLine(firstTotal)
        Console.WriteLine(secondTotal)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "10", "0"]);
}
