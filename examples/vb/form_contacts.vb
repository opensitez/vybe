' VB Contacts Manager — WinForms-style GUI
' Demonstrates: Form, Labels, TextBoxes, Buttons, ListBox, DataGridView

Imports System.Windows.Forms

Module Program
    Dim txtName As TextBox
    Dim txtEmail As TextBox
    Dim txtPhone As TextBox
    Dim grid As DataGridView
    Dim status As Label
    Dim contactCount As Integer = 0

    Sub ClearInputs()
        txtName.Text = ""
        txtEmail.Text = ""
        txtPhone.Text = ""
    End Sub

    Sub OnAddContact()
        If txtName.Text = "" Then
            status.Text = "Enter a name"
            Return
        End If
        grid.Rows.Add(txtName.Text, txtEmail.Text, txtPhone.Text)
        contactCount += 1
        status.Text = CStr(contactCount) & " contacts"
        ClearInputs()
    End Sub

    Sub OnClear()
        ClearInputs()
        status.Text = "Ready - " & CStr(contactCount) & " contacts"
    End Sub

    Sub Main()
        Dim form As New Form()
        form.Text = "Contact Manager"

        ' Header
        Dim header As New Label()
        header.Text = "Contact Manager"
        header.Left = 10
        header.Top = 10
        header.Width = 300
        header.Height = 30
        form.Controls.Add(header)

        ' Name input
        Dim lblName As New Label()
        lblName.Text = "Name:"
        lblName.Left = 10
        lblName.Top = 50
        lblName.Width = 60
        lblName.Height = 25
        form.Controls.Add(lblName)

        txtName = New TextBox()
        txtName.Left = 80
        txtName.Top = 50
        txtName.Width = 200
        txtName.Height = 25
        form.Controls.Add(txtName)

        ' Email input
        Dim lblEmail As New Label()
        lblEmail.Text = "Email:"
        lblEmail.Left = 10
        lblEmail.Top = 85
        lblEmail.Width = 60
        lblEmail.Height = 25
        form.Controls.Add(lblEmail)

        txtEmail = New TextBox()
        txtEmail.Left = 80
        txtEmail.Top = 85
        txtEmail.Width = 200
        txtEmail.Height = 25
        form.Controls.Add(txtEmail)

        ' Phone input
        Dim lblPhone As New Label()
        lblPhone.Text = "Phone:"
        lblPhone.Left = 10
        lblPhone.Top = 120
        lblPhone.Width = 60
        lblPhone.Height = 25
        form.Controls.Add(lblPhone)

        txtPhone = New TextBox()
        txtPhone.Left = 80
        txtPhone.Top = 120
        txtPhone.Width = 200
        txtPhone.Height = 25
        form.Controls.Add(txtPhone)

        ' Buttons
        Dim btnAdd As New Button()
        btnAdd.Text = "Add Contact"
        btnAdd.Left = 80
        btnAdd.Top = 160
        btnAdd.Width = 95
        btnAdd.Height = 35
        form.Controls.Add(btnAdd)
        AddHandler btnAdd.Click, AddressOf OnAddContact

        Dim btnClear As New Button()
        btnClear.Text = "Clear"
        btnClear.Left = 185
        btnClear.Top = 160
        btnClear.Width = 95
        btnClear.Height = 35
        form.Controls.Add(btnClear)
        AddHandler btnClear.Click, AddressOf OnClear

        ' Contacts grid
        grid = New DataGridView()
        grid.Left = 10
        grid.Top = 210
        grid.Width = 560
        grid.Height = 300
        form.Controls.Add(grid)
        grid.Columns.Add("Name", "Name")
        grid.Columns.Add("Email", "Email")
        grid.Columns.Add("Phone", "Phone")

        ' Status bar
        status = New Label()
        status.Text = "Ready — 0 contacts"
        status.Left = 10
        status.Top = 520
        status.Width = 560
        status.Height = 25
        form.Controls.Add(status)

        Application.Run(form)
    End Sub
End Module
