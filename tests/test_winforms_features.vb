Imports System
Imports System.Windows.Forms

Module WinFormsFeatureTest
    Sub Main()
        Console.WriteLine("=== WinForms Feature Tests ===")

        ' --- TextBox properties ---
        Dim txt As New System.Windows.Forms.TextBox()
        txt.Name = "txtTest"
        txt.Text = "Hello"
        txt.ReadOnly = True
        txt.Multiline = True
        txt.PasswordChar = "*"
        Console.WriteLine("TextBox.ReadOnly: " & txt.ReadOnly)
        Console.WriteLine("TextBox.Multiline: " & txt.Multiline)
        Console.WriteLine("TextBox.PasswordChar: " & txt.PasswordChar)

        ' --- ComboBox with Items ---
        Dim cbo As New System.Windows.Forms.ComboBox()
        cbo.Name = "cboTest"
        cbo.Items.Add("Apple")
        cbo.Items.Add("Banana")
        cbo.Items.Add("Cherry")
        Console.WriteLine("ComboBox.Items.Count: " & cbo.Items.Count)
        cbo.SelectedIndex = 1
        Console.WriteLine("ComboBox.SelectedIndex: " & cbo.SelectedIndex)

        ' --- ListBox with Items ---
        Dim lst As New System.Windows.Forms.ListBox()
        lst.Name = "lstTest"
        lst.Items.Add("Red")
        lst.Items.Add("Green")
        lst.Items.Add("Blue")
        Console.WriteLine("ListBox.Items.Count: " & lst.Items.Count)
        lst.Items.Clear()
        Console.WriteLine("ListBox after Clear: " & lst.Items.Count)

        ' --- ProgressBar ---
        Dim pb As New System.Windows.Forms.ProgressBar()
        pb.Name = "pbTest"
        pb.Minimum = 0
        pb.Maximum = 100
        pb.Value = 10
        Console.WriteLine("ProgressBar.Value: " & pb.Value)
        Console.WriteLine("ProgressBar.Maximum: " & pb.Maximum)
        pb.PerformStep()
        Console.WriteLine("ProgressBar after PerformStep: " & pb.Value)

        ' --- NumericUpDown ---
        Dim nud As New System.Windows.Forms.NumericUpDown()
        nud.Name = "nudTest"
        nud.Minimum = 0
        nud.Maximum = 50
        nud.Increment = 5
        nud.Value = 10
        Console.WriteLine("NumericUpDown.Value: " & nud.Value)
        nud.UpButton()
        Console.WriteLine("NumericUpDown after UpButton: " & nud.Value)
        nud.DownButton()
        Console.WriteLine("NumericUpDown after DownButton: " & nud.Value)

        ' --- TreeView with Nodes ---
        Dim tv As New System.Windows.Forms.TreeView()
        tv.Name = "tvTest"
        tv.Nodes.Add("Root1")
        tv.Nodes.Add("Root2")
        Console.WriteLine("TreeView.Nodes.Count: " & tv.Nodes.Count)

        ' --- ListView with Items and Columns ---
        Dim lv As New System.Windows.Forms.ListView()
        lv.Name = "lvTest"
        lv.Columns.Add("Name")
        lv.Columns.Add("Value")
        lv.Items.Add("Item1")
        lv.Items.Add("Item2")
        Console.WriteLine("ListView.Columns.Count: " & lv.Columns.Count)
        Console.WriteLine("ListView.Items.Count: " & lv.Items.Count)

        ' --- DataGridView with Rows and Columns ---
        Dim dgv As New System.Windows.Forms.DataGridView()
        dgv.Name = "dgvTest"
        dgv.Columns.Add("Col1")
        dgv.Columns.Add("Col2")
        dgv.Rows.Add("Row1")
        Console.WriteLine("DataGridView.ColumnCount: " & dgv.ColumnCount)
        Console.WriteLine("DataGridView.RowCount: " & dgv.RowCount)

        ' --- TabControl with TabPages ---
        Dim tab As New System.Windows.Forms.TabControl()
        tab.Name = "tabTest"
        tab.TabPages.Add("Page1")
        tab.TabPages.Add("Page2")
        tab.TabPages.Add("Page3")
        Console.WriteLine("TabControl.TabCount: " & tab.TabCount)
        Console.WriteLine("TabControl.SelectedIndex: " & tab.SelectedIndex)
        tab.SelectedIndex = 2
        Console.WriteLine("TabControl.SelectedIndex after set: " & tab.SelectedIndex)

        ' --- MenuStrip with Items ---
        Dim ms As New System.Windows.Forms.MenuStrip()
        ms.Name = "msTest"
        ms.Items.Add("File")
        ms.Items.Add("Edit")
        Console.WriteLine("MenuStrip.Items.Count: " & ms.Items.Count)

        ' --- ToolStripMenuItem with DropDownItems ---
        Dim tsmi As New System.Windows.Forms.ToolStripMenuItem()
        tsmi.Name = "tsmiFile"
        tsmi.Text = "File"
        tsmi.DropDownItems.Add("New")
        tsmi.DropDownItems.Add("Open")
        tsmi.DropDownItems.Add("Save")
        Console.WriteLine("ToolStripMenuItem.Text: " & tsmi.Text)
        Console.WriteLine("ToolStripMenuItem.DropDownItems.Count: " & tsmi.DropDownItems.Count)

        ' --- Anchor / Dock ---
        Dim pnl As New System.Windows.Forms.Panel()
        pnl.Name = "pnlTest"
        pnl.Dock = DockStyle.Fill
        Console.WriteLine("Panel.Dock: " & pnl.Dock)
        pnl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Console.WriteLine("Panel.Anchor: " & pnl.Anchor)

        Console.WriteLine("=== WinForms Feature Tests Complete ===")
    End Sub
End Module
