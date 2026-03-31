' SQLite Data Binding Demo
' Demonstrates: SqlConnection, SqlDataAdapter, DataTable, BindingSource,
'               DataGridView binding, TextBox DataBindings, CRUD operations
'
' The form creates/opens a local SQLite database (contacts.db),
' loads contacts into a DataTable via a DataAdapter, binds them
' through a BindingSource to a DataGridView and TextBox controls.

Imports System.Data

Public Class Form1

    Dim conn As SqlConnection
    Dim adapter As SqlDataAdapter
    Dim dt As DataTable

    Private Sub Form1_Load() Handles Me.Load
        ' Create and open the SQLite connection
        conn = New SqlConnection("Data Source=contacts.db")
        conn.Open()

        ' Create the contacts table if it doesn't exist
        Dim createCmd As New SqlCommand( _
            "CREATE TABLE IF NOT EXISTS contacts (" & _
            "  Id INTEGER PRIMARY KEY AUTOINCREMENT," & _
            "  Name TEXT NOT NULL," & _
            "  Email TEXT," & _
            "  Phone TEXT" & _
            ")", conn)
        createCmd.ExecuteNonQuery()

        ' Insert sample data if table is empty
        Dim countCmd As New SqlCommand("SELECT COUNT(*) FROM contacts", conn)
        Dim result = countCmd.ExecuteScalar()
        If CInt(result) = 0 Then
            Dim ins1 As New SqlCommand("INSERT INTO contacts (Name, Email, Phone) VALUES ('Alice Smith', 'alice@example.com', '555-0101')", conn)
            ins1.ExecuteNonQuery()
            Dim ins2 As New SqlCommand("INSERT INTO contacts (Name, Email, Phone) VALUES ('Bob Jones', 'bob@example.com', '555-0202')", conn)
            ins2.ExecuteNonQuery()
            Dim ins3 As New SqlCommand("INSERT INTO contacts (Name, Email, Phone) VALUES ('Carol White', 'carol@example.com', '555-0303')", conn)
            ins3.ExecuteNonQuery()
        End If

        ' Load data
        LoadData()
        lblStatus.Text = "Database loaded: contacts.db"
    End Sub

    Private Sub LoadData()
        ' Create adapter and fill DataTable
        adapter = New SqlDataAdapter("SELECT * FROM contacts ORDER BY Name", conn)
        dt = New DataTable("contacts")
        adapter.Fill(dt)

        ' Bind to BindingSource, then to grid
        bs1.DataSource = dt
        lblStatus.Text = "Loaded " & bs1.Count & " contacts"
    End Sub

    Private Sub btnAdd_Click() Handles btnAdd.Click
        Dim name As String = txtName.Text
        Dim email As String = txtEmail.Text
        Dim phone As String = txtPhone.Text

        If name = "" Then
            MsgBox("Please enter a name")
            Exit Sub
        End If

        ' Insert into database
        Dim cmd As New SqlCommand( _
            "INSERT INTO contacts (Name, Email, Phone) VALUES ('" & name & "', '" & email & "', '" & phone & "')", conn)
        cmd.ExecuteNonQuery()

        ' Clear inputs
        txtName.Text = ""
        txtEmail.Text = ""
        txtPhone.Text = ""

        ' Reload data
        LoadData()
        lblStatus.Text = "Added: " & name
    End Sub

    Private Sub btnDelete_Click() Handles btnDelete.Click
        ' Get current position from BindingSource
        Dim pos As Integer = bs1.Position
        If pos < 0 Then
            MsgBox("No contact selected")
            Exit Sub
        End If

        ' Get the current row to find the ID
        Dim row = bs1.Current
        If row Is Nothing Then
            MsgBox("No contact selected")
            Exit Sub
        End If

        Dim id As String = row("Id")
        Dim name As String = row("Name")

        ' Delete from database
        Dim cmd As New SqlCommand("DELETE FROM contacts WHERE Id = " & id, conn)
        cmd.ExecuteNonQuery()

        ' Reload
        LoadData()
        lblStatus.Text = "Deleted: " & name
    End Sub

    Private Sub btnRefresh_Click() Handles btnRefresh.Click
        LoadData()
    End Sub

End Class
