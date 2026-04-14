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

        ' Test 3: Parameterized queries
        TestParameters()

        ' Test 4: Transactions
        TestTransactions()

        ' Test 5: DataAdapter + DataTable
        TestDataAdapter()

        ' Test 6: Constructor args
        TestConstructorArgs()

        Console.WriteLine()
        Console.WriteLine("=== All Data Access Tests Passed ===")
    End Sub

    Sub TestAdoNet()
        Console.WriteLine("--- Test ADO.NET Style ---")

        Dim conn As New System.Data.SqlClient.SqlConnection()
        conn.ConnectionString = "Data Source=:memory:"
        conn.Open()
        Console.WriteLine("Connection opened: State = " & conn.State)

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE Users (Id INTEGER PRIMARY KEY, Name TEXT, Age INTEGER)"
        cmd.ExecuteNonQuery()
        Console.WriteLine("Table created")

        Dim cmdI1 As Object = conn.CreateCommand()
        cmdI1.CommandText = "INSERT INTO Users (Name, Age) VALUES ('Alice', 30)"
        Console.WriteLine("Inserted: " & cmdI1.ExecuteNonQuery())

        Dim cmdI2 As Object = conn.CreateCommand()
        cmdI2.CommandText = "INSERT INTO Users (Name, Age) VALUES ('Bob', 25)"
        Console.WriteLine("Inserted: " & cmdI2.ExecuteNonQuery())

        Dim cmdI3 As Object = conn.CreateCommand()
        cmdI3.CommandText = "INSERT INTO Users (Name, Age) VALUES ('Charlie', 35)"
        Console.WriteLine("Inserted: " & cmdI3.ExecuteNonQuery())

        ' ExecuteScalar
        Dim cmdS As Object = conn.CreateCommand()
        cmdS.CommandText = "SELECT COUNT(*) FROM Users"
        Console.WriteLine("COUNT(*) = " & cmdS.ExecuteScalar())

        ' ExecuteReader
        Dim cmdR As Object = conn.CreateCommand()
        cmdR.CommandText = "SELECT Name, Age FROM Users ORDER BY Age"
        Dim reader As Object = cmdR.ExecuteReader()

        Console.WriteLine("HasRows: " & reader.HasRows)
        Console.WriteLine("FieldCount: " & reader.FieldCount)

        Do While reader.Read()
            Dim name As String = reader.GetString(0)
            Dim age As String = reader.GetValue(1)
            Console.WriteLine("  " & name & ", age " & age)
        Loop

        reader.Close()
        Console.WriteLine("IsClosed: " & reader.IsClosed)

        ' IsDBNull test
        Dim cmdNull As Object = conn.CreateCommand()
        cmdNull.CommandText = "SELECT NULL AS NullCol, 'OK' AS OkCol"
        Dim rdr2 As Object = cmdNull.ExecuteReader()
        If rdr2.Read() Then
            Console.WriteLine("IsDBNull(0): " & rdr2.IsDBNull(0))
            Console.WriteLine("IsDBNull(1): " & rdr2.IsDBNull(1))
            Console.WriteLine("GetName(1): " & rdr2.GetName(1))
        End If
        rdr2.Close()

        conn.Close()
        Console.WriteLine("Connection closed")
        Console.WriteLine()
    End Sub

    Sub TestAdodb()
        Console.WriteLine("--- Test ADODB Style ---")

        Dim conn As New ADODB.Connection()
        conn.ConnectionString = "Data Source=:memory:"
        conn.Open()
        Console.WriteLine("ADODB Connection opened")

        conn.Execute("CREATE TABLE Products (Id INTEGER PRIMARY KEY, Name TEXT, Price REAL)")
        conn.Execute("INSERT INTO Products (Name, Price) VALUES ('Widget', 9.99)")
        conn.Execute("INSERT INTO Products (Name, Price) VALUES ('Gadget', 19.99)")
        conn.Execute("INSERT INTO Products (Name, Price) VALUES ('Doohickey', 4.50)")
        Console.WriteLine("Table + 3 rows inserted")

        Dim rs As New ADODB.Recordset()
        rs.Open("SELECT Name, Price FROM Products ORDER BY Price", conn)
        Console.WriteLine("RecordCount: " & rs.RecordCount)

        ' Test Fields by name
        Do While Not rs.EOF
            Console.WriteLine("  " & rs.Fields("Name").Value & " - $" & rs.Fields("Price").Value)
            rs.MoveNext()
        Loop

        ' Test Fields by index
        rs.MoveFirst()
        Console.WriteLine("Field(0) by index: " & rs.Fields(0).Value)

        rs.Close()
        conn.Close()
        Console.WriteLine("ADODB done")
        Console.WriteLine()
    End Sub

    Sub TestParameters()
        Console.WriteLine("--- Test Parameterized Queries ---")

        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmdC As Object = conn.CreateCommand()
        cmdC.CommandText = "CREATE TABLE Items (Id INTEGER PRIMARY KEY, Name TEXT, Qty INTEGER)"
        cmdC.ExecuteNonQuery()

        ' Use Parameters.AddWithValue
        Dim cmdIns As Object = conn.CreateCommand()
        cmdIns.CommandText = "INSERT INTO Items (Name, Qty) VALUES (@name, @qty)"
        cmdIns.Parameters.AddWithValue("@name", "Apples")
        cmdIns.Parameters.AddWithValue("@qty", "10")
        cmdIns.ExecuteNonQuery()
        Console.WriteLine("Inserted with params: Apples, 10")

        ' Reuse with new params
        cmdIns.Parameters.Clear()
        cmdIns.Parameters.AddWithValue("@name", "Bananas")
        cmdIns.Parameters.AddWithValue("@qty", "20")
        cmdIns.ExecuteNonQuery()
        Console.WriteLine("Inserted with params: Bananas, 20")

        ' Verify
        Dim cmdV As Object = conn.CreateCommand()
        cmdV.CommandText = "SELECT Name, Qty FROM Items ORDER BY Qty"
        Dim r As Object = cmdV.ExecuteReader()
        Do While r.Read()
            Console.WriteLine("  " & r.GetString(0) & " qty=" & r.GetValue(1))
        Loop
        r.Close()

        conn.Close()
        Console.WriteLine("Parameters done")
        Console.WriteLine()
    End Sub

    Sub TestTransactions()
        Console.WriteLine("--- Test Transactions ---")

        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmdC As Object = conn.CreateCommand()
        cmdC.CommandText = "CREATE TABLE Accounts (Name TEXT, Balance INTEGER)"
        cmdC.ExecuteNonQuery()

        Dim cmdI As Object = conn.CreateCommand()
        cmdI.CommandText = "INSERT INTO Accounts VALUES ('Alice', 100)"
        cmdI.ExecuteNonQuery()

        ' Start transaction and update
        Dim trans As Object = conn.BeginTransaction()
        Dim cmdU As Object = conn.CreateCommand()
        cmdU.CommandText = "UPDATE Accounts SET Balance = 200 WHERE Name = 'Alice'"
        cmdU.ExecuteNonQuery()

        ' Rollback
        trans.Rollback()
        Console.WriteLine("Rolled back")

        ' Check balance is still 100
        Dim cmdQ As Object = conn.CreateCommand()
        cmdQ.CommandText = "SELECT Balance FROM Accounts WHERE Name = 'Alice'"
        Dim bal As Object = cmdQ.ExecuteScalar()
        Console.WriteLine("Balance after rollback: " & bal)

        ' Now commit a change
        Dim trans2 As Object = conn.BeginTransaction()
        Dim cmdU2 As Object = conn.CreateCommand()
        cmdU2.CommandText = "UPDATE Accounts SET Balance = 500 WHERE Name = 'Alice'"
        cmdU2.ExecuteNonQuery()
        trans2.Commit()
        Console.WriteLine("Committed")

        Dim bal2 As Object = cmdQ.ExecuteScalar()
        Console.WriteLine("Balance after commit: " & bal2)

        conn.Close()
        Console.WriteLine("Transactions done")
        Console.WriteLine()
    End Sub

    Sub TestDataAdapter()
        Console.WriteLine("--- Test DataAdapter ---")

        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmdC As Object = conn.CreateCommand()
        cmdC.CommandText = "CREATE TABLE Colors (Id INTEGER, Name TEXT)"
        cmdC.ExecuteNonQuery()

        Dim cmdI As Object = conn.CreateCommand()
        cmdI.CommandText = "INSERT INTO Colors VALUES (1, 'Red')"
        cmdI.ExecuteNonQuery()
        Dim cmdI2 As Object = conn.CreateCommand()
        cmdI2.CommandText = "INSERT INTO Colors VALUES (2, 'Blue')"
        cmdI2.ExecuteNonQuery()
        Dim cmdI3 As Object = conn.CreateCommand()
        cmdI3.CommandText = "INSERT INTO Colors VALUES (3, 'Green')"
        cmdI3.ExecuteNonQuery()

        Dim dt As New DataTable("Colors")
        Dim adapter As New System.Data.SqlClient.SqlDataAdapter("SELECT Id, Name FROM Colors ORDER BY Id", conn)
        Dim rowCount As Integer = adapter.Fill(dt)
        Console.WriteLine("Filled " & rowCount & " rows")

        ' Access rows
        Dim rows As Object = dt.Rows
        Console.WriteLine("Rows.Length: " & rows.Length)

        conn.Close()
        Console.WriteLine("DataAdapter done")
        Console.WriteLine()
    End Sub

    Sub TestConstructorArgs()
        Console.WriteLine("--- Test Constructor Args ---")

        ' New SqlConnection(connStr)
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()
        Console.WriteLine("SqlConnection(connStr) opened: State = " & conn.State)

        Dim cmdC As Object = conn.CreateCommand()
        cmdC.CommandText = "CREATE TABLE T1 (X INTEGER)"
        cmdC.ExecuteNonQuery()

        ' New SqlCommand(sql, conn)
        Dim cmd As New System.Data.SqlClient.SqlCommand("INSERT INTO T1 VALUES (42)", conn)
        cmd.ExecuteNonQuery()
        Console.WriteLine("SqlCommand(sql, conn) inserted row")

        Dim cmdQ As New System.Data.SqlClient.SqlCommand("SELECT X FROM T1", conn)
        Dim val As Object = cmdQ.ExecuteScalar()
        Console.WriteLine("Value: " & val)

        conn.Close()
        Console.WriteLine("Constructor args done")
        Console.WriteLine()
    End Sub

End Module
