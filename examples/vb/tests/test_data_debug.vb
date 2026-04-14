Imports System
Imports System.Data

Module TestDataDebug
    Dim counter As Integer = 0

    Function Increment() As Integer
        counter = counter + 1
        Return counter
    End Function

    Sub Main()
        ' Test: method call inside concat should only execute once
        counter = 0
        Console.WriteLine("val=" & Increment())
        Console.WriteLine("counter should be 1: " & counter)

        ' Test: DB method call inside concat should only execute once
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE T (X INTEGER)"
        cmd.ExecuteNonQuery()

        Dim cmdI As Object = conn.CreateCommand()
        cmdI.CommandText = "INSERT INTO T VALUES (1)"
        Console.WriteLine("result=" & cmdI.ExecuteNonQuery())

        Dim cmdC As Object = conn.CreateCommand()
        cmdC.CommandText = "SELECT COUNT(*) FROM T"
        Console.WriteLine("count should be 1: " & cmdC.ExecuteScalar())

        conn.Close()
    End Sub
End Module
