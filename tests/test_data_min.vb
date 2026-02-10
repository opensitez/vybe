Imports System

Module TestDataMin
    Sub Main()
        Console.WriteLine("Start")

        Dim conn As New System.Data.SqlClient.SqlConnection()
        conn.ConnectionString = "Data Source=:memory:"
        conn.Open()
        Console.WriteLine("Opened")

        Dim cmd As Object
        cmd = conn.CreateCommand()
        Console.WriteLine("Got command")
        cmd.CommandText = "CREATE TABLE T1 (Id INTEGER, Name TEXT)"
        Console.WriteLine("Set command text: " & cmd.CommandText)

        Dim r As Integer = cmd.ExecuteNonQuery()
        Console.WriteLine("ExecuteNonQuery returned: " & r)

        ' Insert rows
        Console.WriteLine("About to set command text for insert")
        cmd.CommandText = "INSERT INTO T1 (Id, Name) VALUES (1, 'Alice')"
        Console.WriteLine("Set insert text: " & cmd.CommandText)
        Console.WriteLine("conn_id check")
        
        ' Try creating a new command instead
        Dim cmd2 As Object = conn.CreateCommand()
        cmd2.CommandText = "INSERT INTO T1 (Id, Name) VALUES (1, 'Alice')"
        Console.WriteLine("About to execute insert via cmd2")
        Dim r2 As Integer = cmd2.ExecuteNonQuery()
        Console.WriteLine("Insert 1: " & r2)

        cmd.CommandText = "INSERT INTO T1 (Id, Name) VALUES (2, 'Bob')"
        Dim r3 As Integer = cmd.ExecuteNonQuery()
        Console.WriteLine("Insert 2: " & r3)

        ' ExecuteScalar
        cmd.CommandText = "SELECT COUNT(*) FROM T1"
        Dim cnt As String = cmd.ExecuteScalar()
        Console.WriteLine("Count: " & cnt)

        ' ExecuteReader
        cmd.CommandText = "SELECT Id, Name FROM T1 ORDER BY Id"
        Dim reader As Object = cmd.ExecuteReader()
        Console.WriteLine("HasRows: " & reader.HasRows)

        Do While reader.Read()
            Console.WriteLine("Row: " & reader.GetString(0) & " - " & reader.GetString(1))
        Loop

        reader.Close()
        conn.Close()

        Console.WriteLine("Done")
    End Sub
End Module
