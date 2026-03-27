' VB Contacts Manager — WinForms-style GUI
' Demonstrates: Form, Labels, TextBoxes, Buttons, ListBox, DataGridView

Module Program
    Sub Main()
        Dim form As Object = Window.Forms.Form("Contact Manager")

        ' Header
        Dim header As Object = Window.Forms.Label()
        header.text = "Contact Manager"
        header.left = 10
        header.top = 10
        header.width = 300
        header.height = 30
        vybe.gui.controlsAdd("Contact Manager", header)

        ' Name input
        Dim lblName As Object = Window.Forms.Label()
        lblName.text = "Name:"
        lblName.left = 10
        lblName.top = 50
        lblName.width = 60
        lblName.height = 25
        vybe.gui.controlsAdd("Contact Manager", lblName)

        Dim txtName As Object = Window.Forms.TextBox()
        txtName.left = 80
        txtName.top = 50
        txtName.width = 200
        txtName.height = 25
        vybe.gui.controlsAdd("Contact Manager", txtName)

        ' Email input
        Dim lblEmail As Object = Window.Forms.Label()
        lblEmail.text = "Email:"
        lblEmail.left = 10
        lblEmail.top = 85
        lblEmail.width = 60
        lblEmail.height = 25
        vybe.gui.controlsAdd("Contact Manager", lblEmail)

        Dim txtEmail As Object = Window.Forms.TextBox()
        txtEmail.left = 80
        txtEmail.top = 85
        txtEmail.width = 200
        txtEmail.height = 25
        vybe.gui.controlsAdd("Contact Manager", txtEmail)

        ' Phone input
        Dim lblPhone As Object = Window.Forms.Label()
        lblPhone.text = "Phone:"
        lblPhone.left = 10
        lblPhone.top = 120
        lblPhone.width = 60
        lblPhone.height = 25
        vybe.gui.controlsAdd("Contact Manager", lblPhone)

        Dim txtPhone As Object = Window.Forms.TextBox()
        txtPhone.left = 80
        txtPhone.top = 120
        txtPhone.width = 200
        txtPhone.height = 25
        vybe.gui.controlsAdd("Contact Manager", txtPhone)

        ' Buttons
        Dim btnAdd As Object = Window.Forms.Button()
        btnAdd.text = "Add Contact"
        btnAdd.left = 80
        btnAdd.top = 160
        btnAdd.width = 95
        btnAdd.height = 35
        vybe.gui.controlsAdd("Contact Manager", btnAdd)

        Dim btnClear As Object = Window.Forms.Button()
        btnClear.text = "Clear"
        btnClear.left = 185
        btnClear.top = 160
        btnClear.width = 95
        btnClear.height = 35
        vybe.gui.controlsAdd("Contact Manager", btnClear)

        ' Contacts grid
        Dim grid As Object = Window.Forms.DataGridView()
        grid.left = 10
        grid.top = 210
        grid.width = 560
        grid.height = 300
        vybe.gui.controlsAdd("Contact Manager", grid)

        ' Status bar
        Dim status As Object = Window.Forms.Label()
        status.text = "Ready — 0 contacts"
        status.left = 10
        status.top = 520
        status.width = 560
        status.height = 25
        vybe.gui.controlsAdd("Contact Manager", status)

        Application.Run("Contact Manager")
    End Sub
End Module
