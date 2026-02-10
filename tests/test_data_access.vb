Imports System
Imports System.Data

Module TestDataAccess

    Sub Main()
        Console.WriteLine("=== Data Access Tests ===")
        Console.WriteLine()

        ' Test 1: ADO.NET style with SQLite in-memory
        TestAdoNet()

        ' Test 2: ADODB style with SQLite in-memory
        TestAdodb()

        Console.WriteLine()
        Console.WriteLine("=== All Data Access Tests Passed ===")
    End Sub

    Sub TestAdoNet()
        Console.WriteLine("--- Test ADO.NET Style ---")

        ' Create connection
        Dim conn As New System.Data.SqlClient.SqlConnection()
        conn.ConnectionString = "Data Source=:memory:"
        conn.Open()
        Console.WriteLine("Connection opened: State = " & conn.State)

        ' Create table
        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE Users (Id INTEGER PRIMARY KEY, Name TEXT, Age INTEGER)"
        Dim createResult As Integer = cmd.ExecuteNonQuery()
        Console.WriteLine("Table created")

        ' Insert data
        Dim cmdInsert As Object = conn.CreateCommand()
        cmdInsert.CommandText = "INSERT INTO Users (Name, Age) VALUES ('Alice', 30)"
        Dim rows1 As Integer = cmdInsert.ExecuteNonQuery()
        Console.WriteLine("Inserted row 1, affected: " & rows1)

        Dim cmdInsert2 As Object = conn.CreateCommand()
        cmdInsert2.CommandText = "INSERT INTO Users (Name, Age) VALUES ('Bob', 25)"
        Dim rows2 As Integer = cmdInsert2.ExecuteNonQuery()
        Console.WriteLine("Inserted row 2, affected: " & rows2)

        Dim cmdInsert3 As Object = conn.CreateCommand()
        cmdInsert3.CommandText = "INSERT INTO Users (Name, Age) VALUES ('Charlie', 35)"
        Dim rows3 As Integer = cmdInsert3.ExecuteNonQuery()
        Console.WriteLine("Inserted row 3, affected: " & rows3)

        ' ExecuteScalar
        Dim cmdScalar As Object = conn.CreateCommand()
        cmdScalar.CommandText = "SELECT COUNT(*) FROM Users"
        Dim count As Object = cmdScalar.ExecuteScalar()
        Console.WriteLine("ExecuteScalar COUNT(*) = " & count)

        ' ExecuteReader
        Dim cmdSelect As Object = conn.CreateCommand()
        cmdSelect.CommandText = "SELECT Name, Age FROM Users ORDER BY Age"
        Dim reader As Object = cmdSelect.ExecuteReader()

        Console.WriteLine("HasRows: " & reader.HasRows)
        Console.WriteLine("FieldCount: " & reader.FieldCount)

        Dim rowNum As Integer = 0
        Do While reader.Read()
            rowNum = rowNum + 1
            Dim name As String = reader.GetString(0)
            Dim age As String = reader.GetValue(1)
            Console.WriteLine("  Row " & rowNum & ": " & name & ", age " & age)
        Loop

        reader.Close()
        Console.WriteLine("Reader closed, IsClosed: " & reader.IsClosed)

        ' Close connection
        conn.Close()
        Console.WriteLine("Connection closed")
        Console.WriteLine()
    End Sub

    Sub TestAdodb()
        Console.WriteLine("--- Test ADODB Style ---")

        ' Create ADODB connection
        Dim conn As New ADODB.Connection()
        conn.ConnectionString = "Data Source=:memory:"
        conn.Open()
        Console.WriteLine("ADODB Connection opened")

        ' Create table using Execute
        conn.Execute("CREATE TABLE Products (Id INTEGER PRIMARY KEY, Name TEXT, Price REAL)")
        Console.WriteLine("Table created via Execute")

        ' Insert rows
        conn.Execute("INSERT INTO Products (Name, Price) VALUES ('Widget', 9.99)")
        conn.Execute("INSERT INTO Products (Name, Price) VALUES ('Gadget', 19.99)")
        conn.Execute("INSERT INTO Products (Name, Price) VALUES ('Doohickey', 4.50)")
        Console.WriteLine("3 rows inserted")

        ' Query using Recordset
        Dim rsQuery As New ADODB.Recordset()
        rsQuery.Open("SELECT Name, Price FROM Products ORDER BY Price", conn)
        Console.WriteLine("RecordCount: " & rsQuery.RecordCount)

        Dim n As Integer = 0
        Do While Not rsQuery.EOF
            n = n + 1
            Dim prodName As String = rsQuery.Fields("Name").Value
            Dim prodPrice As String = rsQuery.Fields("Price").Value
            Console.WriteLine("Product " & n & ": " & prodName & " - $" & prodPrice)
            rsQuery.MoveNext()
        Loop

        rsQuery.Close()
        Console.WriteLine("Recordset closed")

        conn.Close()
        Console.WriteLine("ADODB Connection closed")
    End Sub

End Module
