Imports System
Imports System.Data

Module TestDataAccessFull

    Sub Main()
        Console.WriteLine("=== Full Data Access Feature Tests ===")
        Console.WriteLine()

        TestBasicAdoNet()
        TestAsyncMethods()
        TestCommandTimeout()
        TestCommandTypeEnum()
        TestSchema()
        TestDataSet()
        TestAdodbCommand()
        TestDbNullValue()
        TestDataTableMethods()
        TestParametersCollection()
        TestServerVersion()

        Console.WriteLine()
        Console.WriteLine("=== All Full Data Access Tests Passed ===")
    End Sub

    Sub TestBasicAdoNet()
        Console.WriteLine("--- Test Basic ADO.NET ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()
        Console.WriteLine("State: " & conn.State)

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE T1 (Id INTEGER PRIMARY KEY, Name TEXT, Score REAL)"
        cmd.ExecuteNonQuery()

        Dim ins As Object = conn.CreateCommand()
        ins.CommandText = "INSERT INTO T1 VALUES (1, 'Alice', 95.5)"
        ins.ExecuteNonQuery()
        ins.CommandText = "INSERT INTO T1 VALUES (2, 'Bob', 87.0)"
        ins.ExecuteNonQuery()

        Dim q As Object = conn.CreateCommand()
        q.CommandText = "SELECT COUNT(*) FROM T1"
        Console.WriteLine("Count: " & q.ExecuteScalar())

        Dim r As Object = conn.CreateCommand()
        r.CommandText = "SELECT Name, Score FROM T1 ORDER BY Id"
        Dim reader As Object = r.ExecuteReader()
        Console.WriteLine("HasRows: " & reader.HasRows)
        Console.WriteLine("FieldCount: " & reader.FieldCount)
        Do While reader.Read()
            Console.WriteLine("  " & reader.GetString(0) & " = " & reader.GetValue(1))
        Loop
        Console.WriteLine("GetName(0): " & reader.GetName(0))
        reader.Close()
        Console.WriteLine("IsClosed: " & reader.IsClosed)

        conn.Close()
        Console.WriteLine("Basic ADO.NET done")
        Console.WriteLine()
    End Sub

    Sub TestAsyncMethods()
        Console.WriteLine("--- Test Async Wrappers ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE AsyncTest (Id INTEGER, Val TEXT)"
        cmd.ExecuteNonQuery()

        ' Test ExecuteNonQueryAsync
        Dim ins As Object = conn.CreateCommand()
        ins.CommandText = "INSERT INTO AsyncTest VALUES (1, 'async1')"
        ins.ExecuteNonQueryAsync()
        Console.WriteLine("ExecuteNonQueryAsync done")

        ' Test ExecuteScalarAsync
        Dim qs As Object = conn.CreateCommand()
        qs.CommandText = "SELECT COUNT(*) FROM AsyncTest"
        Dim cnt As Object = qs.ExecuteScalarAsync()
        Console.WriteLine("ExecuteScalarAsync count: " & cnt)

        ' Test ExecuteReaderAsync
        Dim qr As Object = conn.CreateCommand()
        qr.CommandText = "SELECT Val FROM AsyncTest"
        Dim rdr As Object = qr.ExecuteReaderAsync()
        Do While rdr.Read()
            Console.WriteLine("  async val: " & rdr.GetString(0))
        Loop
        rdr.Close()

        conn.Close()
        Console.WriteLine("Async wrappers done")
        Console.WriteLine()
    End Sub

    Sub TestCommandTimeout()
        Console.WriteLine("--- Test Command Timeout ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmd As Object = conn.CreateCommand()
        Console.WriteLine("Default timeout: " & cmd.CommandTimeout)

        cmd.CommandTimeout = 60
        Console.WriteLine("Updated timeout: " & cmd.CommandTimeout)

        cmd.CommandText = "SELECT 1"
        Console.WriteLine("CommandText: " & cmd.CommandText)

        conn.Close()
        Console.WriteLine("Command timeout done")
        Console.WriteLine()
    End Sub

    Sub TestCommandTypeEnum()
        Console.WriteLine("--- Test CommandType Enum ---")
        Console.WriteLine("CommandType.Text = " & CommandType.Text)
        Console.WriteLine("CommandType.StoredProcedure = " & CommandType.StoredProcedure)
        Console.WriteLine("ConnectionState.Open = " & ConnectionState.Open)
        Console.WriteLine("ConnectionState.Closed = " & ConnectionState.Closed)
        Console.WriteLine("Enum values done")
        Console.WriteLine()
    End Sub

    Sub TestSchema()
        Console.WriteLine("--- Test Schema Discovery ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE SchemaTest (Col1 INTEGER, Col2 TEXT, Col3 REAL)"
        cmd.ExecuteNonQuery()

        ' GetSchema - tables
        Dim schemaTable As Object = conn.GetSchema("Tables")
        Dim schemaRows As Object = schemaTable.Rows
        Console.WriteLine("Schema tables found: " & schemaRows.Length)

        ' Reader GetSchemaTable
        Dim qr As Object = conn.CreateCommand()
        qr.CommandText = "SELECT Col1, Col2, Col3 FROM SchemaTest"
        Dim rdr As Object = qr.ExecuteReader()
        Dim st As Object = rdr.GetSchemaTable()
        Dim stRows As Object = st.Rows
        Console.WriteLine("Schema columns: " & stRows.Length)
        rdr.Close()

        conn.Close()
        Console.WriteLine("Schema done")
        Console.WriteLine()
    End Sub

    Sub TestDataSet()
        Console.WriteLine("--- Test DataSet ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE DSTest (Id INTEGER, Color TEXT)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "INSERT INTO DSTest VALUES (1, 'Red')"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "INSERT INTO DSTest VALUES (2, 'Blue')"
        cmd.ExecuteNonQuery()

        Dim ds As New DataSet("TestDS")
        Console.WriteLine("DataSet name: " & ds.DataSetName)

        Dim adapter As New System.Data.SqlClient.SqlDataAdapter("SELECT Id, Color FROM DSTest ORDER BY Id", conn)
        Dim rowCount As Integer = adapter.Fill(ds)
        Console.WriteLine("Filled " & rowCount & " rows into DataSet")

        ' Access tables
        Dim tables As Object = ds.Tables
        Console.WriteLine("Tables count: " & tables.Length)

        ' Access first table by index
        Dim dt As Object = ds.Tables(0)
        Dim rows As Object = dt.Rows
        Console.WriteLine("Rows in table: " & rows.Length)

        conn.Close()
        Console.WriteLine("DataSet done")
        Console.WriteLine()
    End Sub

    Sub TestAdodbCommand()
        Console.WriteLine("--- Test ADODB Command ---")
        Dim conn As New ADODB.Connection()
        conn.ConnectionString = "Data Source=:memory:"
        conn.Open()

        conn.Execute("CREATE TABLE CmdTest (Id INTEGER, Name TEXT)")
        conn.Execute("INSERT INTO CmdTest VALUES (1, 'test1')")
        conn.Execute("INSERT INTO CmdTest VALUES (2, 'test2')")

        ' Create a command
        Dim cmd As New ADODB.Command()
        cmd.CommandText = "SELECT Name FROM CmdTest ORDER BY Id"
        cmd.CommandType = 1
        cmd.ConnectionString = conn.ConnectionString

        ' Use a new connection for the command (ADODB style uses ActiveConnection)
        Dim conn2 As New ADODB.Connection()
        conn2.ConnectionString = "Data Source=:memory:"
        conn2.Open()
        conn2.Execute("CREATE TABLE CmdTest (Id INTEGER, Name TEXT)")
        conn2.Execute("INSERT INTO CmdTest VALUES (1, 'cmd1')")
        conn2.Execute("INSERT INTO CmdTest VALUES (2, 'cmd2')")

        ' Use CreateParameter
        Dim param As Object = cmd.CreateParameter("Name", 200, 1, 50, "test")
        Console.WriteLine("Parameter created: Name")

        conn.Close()
        conn2.Close()
        Console.WriteLine("ADODB Command done")
        Console.WriteLine()
    End Sub

    Sub TestDbNullValue()
        Console.WriteLine("--- Test DBNull.Value ---")
        Dim nullVal As Object = DBNull.Value
        If nullVal Is Nothing Then
            Console.WriteLine("DBNull.Value is Nothing: True")
        Else
            Console.WriteLine("DBNull.Value is Nothing: False")
        End If
        Console.WriteLine("DBNull done")
        Console.WriteLine()
    End Sub

    Sub TestDataTableMethods()
        Console.WriteLine("--- Test DataTable Methods ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE DTMethods (Id INTEGER, Val TEXT)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "INSERT INTO DTMethods VALUES (1, 'A')"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "INSERT INTO DTMethods VALUES (2, 'B')"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "INSERT INTO DTMethods VALUES (3, 'C')"
        cmd.ExecuteNonQuery()

        Dim dt As New DataTable("DTMethods")
        Dim adapter As New System.Data.SqlClient.SqlDataAdapter("SELECT Id, Val FROM DTMethods ORDER BY Id", conn)
        adapter.Fill(dt)

        ' Test .Rows access
        Dim rows As Object = dt.Rows
        Console.WriteLine("Rows: " & rows.Length)

        ' Test .Columns access
        Dim cols As Object = dt.Columns
        Console.WriteLine("Columns: " & cols.Length)

        ' Test .TableName
        Console.WriteLine("TableName: " & dt.TableName)

        ' Test DataTable.Select()
        Dim allRows As Object = dt.Select()
        Console.WriteLine("Select() count: " & allRows.Length)

        ' Test DataRow.Item
        Dim firstRow As Object = rows(0)
        Console.WriteLine("Row(0).Item(""val""): " & firstRow.Item("Val"))
        Console.WriteLine("Row(0).Item(0): " & firstRow.Item(0))

        ' Test DataRow.IsNull
        Console.WriteLine("Row(0).IsNull(""val""): " & firstRow.IsNull("Val"))

        conn.Close()
        Console.WriteLine("DataTable methods done")
        Console.WriteLine()
    End Sub

    Sub TestParametersCollection()
        Console.WriteLine("--- Test Parameters Collection ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Dim cmd As Object = conn.CreateCommand()
        cmd.CommandText = "CREATE TABLE ParamTest (Name TEXT, Age INTEGER)"
        cmd.ExecuteNonQuery()

        Dim ins As Object = conn.CreateCommand()
        ins.CommandText = "INSERT INTO ParamTest VALUES (@name, @age)"
        ins.Parameters.AddWithValue("@name", "Dave")
        ins.Parameters.AddWithValue("@age", "28")

        ' Test Parameters.Count
        Dim cnt As Integer = ins.Parameters.Count
        Console.WriteLine("Params count: " & cnt)

        ins.ExecuteNonQuery()

        ' Verify
        Dim q As Object = conn.CreateCommand()
        q.CommandText = "SELECT Name, Age FROM ParamTest"
        Dim rdr As Object = q.ExecuteReader()
        Do While rdr.Read()
            Console.WriteLine("  " & rdr.GetString(0) & " age=" & rdr.GetValue(1))
        Loop
        rdr.Close()

        ' Test clear and re-add
        ins.Parameters.Clear()
        ins.Parameters.AddWithValue("@name", "Eve")
        ins.Parameters.AddWithValue("@age", "35")
        ins.ExecuteNonQuery()

        Dim q2 As Object = conn.CreateCommand()
        q2.CommandText = "SELECT COUNT(*) FROM ParamTest"
        Console.WriteLine("Total rows: " & q2.ExecuteScalar())

        conn.Close()
        Console.WriteLine("Parameters collection done")
        Console.WriteLine()
    End Sub

    Sub TestServerVersion()
        Console.WriteLine("--- Test Server Version ---")
        Dim conn As New System.Data.SqlClient.SqlConnection("Data Source=:memory:")
        conn.Open()

        Console.WriteLine("ServerVersion: " & conn.ServerVersion)
        Console.WriteLine("Provider: " & conn.Provider)
        Console.WriteLine("ConnectionTimeout: " & conn.ConnectionTimeout)

        conn.Close()
        Console.WriteLine("Server version done")
        Console.WriteLine()
    End Sub

End Module
