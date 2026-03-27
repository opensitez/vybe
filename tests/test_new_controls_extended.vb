Imports System.Windows.Forms
Imports System.ComponentModel

Module TestNewControlsExtended

    Sub Pass(msg As String)
        Console.WriteLine("PASS: " & msg)
    End Sub

    Sub Fail(msg As String)
        Console.WriteLine("FAIL: " & msg)
    End Sub

    Sub Check(cond As Boolean, msg As String)
        If cond Then
            Pass(msg)
        Else
            Fail(msg)
        End If
    End Sub

    Sub Main()

        ' ===== CheckedListBox =====
        Dim clb As New System.Windows.Forms.CheckedListBox()
        Check(clb.Items IsNot Nothing, "CheckedListBox.Items not Nothing")
        Check(clb.SelectedIndex = -1, "CheckedListBox.SelectedIndex default is -1")
        Check(clb.SelectionMode = 1, "CheckedListBox.SelectionMode default is 1")
        Check(clb.Sorted = False, "CheckedListBox.Sorted default is False")
        Check(clb.CheckOnClick = False, "CheckedListBox.CheckOnClick default is False")

        ' Add items and test SetItemChecked / GetItemChecked
        clb.Items.Add("Apple")
        clb.Items.Add("Banana")
        clb.Items.Add("Cherry")
        Check(clb.Items.Count = 3, "CheckedListBox.Items.Count after Add is 3")

        clb.SetItemChecked(0, True)
        Check(clb.GetItemChecked(0) = True, "CheckedListBox.GetItemChecked(0) after SetItemChecked(0,True)")
        Check(clb.GetItemChecked(1) = False, "CheckedListBox.GetItemChecked(1) default is False")

        clb.SetItemChecked(1, True)
        clb.SetItemChecked(1, False)
        Check(clb.GetItemChecked(1) = False, "CheckedListBox.GetItemChecked(1) after uncheck")

        ' GetItemCheckState returns integer: 0=Unchecked, 1=Checked
        Dim state0 As Integer = clb.GetItemCheckState(0)
        Check(state0 = 1, "CheckedListBox.GetItemCheckState(0) is 1 (Checked)")
        Dim state1 As Integer = clb.GetItemCheckState(1)
        Check(state1 = 0, "CheckedListBox.GetItemCheckState(1) is 0 (Unchecked)")

        ' SetItemCheckState
        clb.SetItemCheckState(2, 1)
        Check(clb.GetItemCheckState(2) = 1, "CheckedListBox.SetItemCheckState(2,1) works")

        ' Property setters
        clb.Sorted = True
        Check(clb.Sorted = True, "CheckedListBox.Sorted set to True")
        clb.CheckOnClick = True
        Check(clb.CheckOnClick = True, "CheckedListBox.CheckOnClick set to True")

        ' ===== DomainUpDown =====
        Dim dud As New System.Windows.Forms.DomainUpDown()
        Check(dud.Items IsNot Nothing, "DomainUpDown.Items not Nothing")
        Check(dud.SelectedIndex = -1, "DomainUpDown.SelectedIndex default is -1")
        Check(dud.Text = "", "DomainUpDown.Text default is empty")
        Check(dud.ReadOnly = False, "DomainUpDown.ReadOnly default is False")
        Check(dud.Wrap = False, "DomainUpDown.Wrap default is False")
        Check(dud.Sorted = False, "DomainUpDown.Sorted default is False")

        dud.Items.Add("Red")
        dud.Items.Add("Green")
        dud.Items.Add("Blue")
        Check(dud.Items.Count = 3, "DomainUpDown.Items.Count after Add is 3")

        ' DownButton moves forward through items
        dud.DownButton()
        Check(dud.SelectedIndex = 0, "DomainUpDown.DownButton sets SelectedIndex to 0")
        Check(dud.Text = "Red", "DomainUpDown.Text is Red after first DownButton")

        dud.DownButton()
        Check(dud.SelectedIndex = 1, "DomainUpDown.DownButton advances to index 1")
        Check(dud.Text = "Green", "DomainUpDown.Text is Green")

        dud.DownButton()
        Check(dud.SelectedIndex = 2, "DomainUpDown.DownButton advances to index 2")
        Check(dud.Text = "Blue", "DomainUpDown.Text is Blue")

        ' DownButton at end should clamp
        dud.DownButton()
        Check(dud.SelectedIndex = 2, "DomainUpDown.DownButton clamps at last item")

        ' UpButton moves backward
        dud.UpButton()
        Check(dud.SelectedIndex = 1, "DomainUpDown.UpButton moves to index 1")
        Check(dud.Text = "Green", "DomainUpDown.Text is Green after UpButton")

        dud.UpButton()
        Check(dud.SelectedIndex = 0, "DomainUpDown.UpButton moves to index 0")

        ' UpButton at start should clamp
        dud.UpButton()
        Check(dud.SelectedIndex = 0, "DomainUpDown.UpButton clamps at first item")

        ' Property setters
        dud.ReadOnly = True
        Check(dud.ReadOnly = True, "DomainUpDown.ReadOnly set to True")
        dud.Wrap = True
        Check(dud.Wrap = True, "DomainUpDown.Wrap set to True")
        dud.Sorted = True
        Check(dud.Sorted = True, "DomainUpDown.Sorted set to True")

        ' ===== BackgroundWorker =====
        Dim bgw As New System.ComponentModel.BackgroundWorker()
        Check(bgw.IsBusy = False, "BackgroundWorker.IsBusy default is False")
        Check(bgw.CancellationPending = False, "BackgroundWorker.CancellationPending default is False")
        Check(bgw.WorkerReportsProgress = False, "BackgroundWorker.WorkerReportsProgress default is False")
        Check(bgw.WorkerSupportsCancellation = False, "BackgroundWorker.WorkerSupportsCancellation default is False")

        bgw.WorkerReportsProgress = True
        Check(bgw.WorkerReportsProgress = True, "BackgroundWorker.WorkerReportsProgress set to True")

        bgw.WorkerSupportsCancellation = True
        Check(bgw.WorkerSupportsCancellation = True, "BackgroundWorker.WorkerSupportsCancellation set to True")

        ' RunWorkerAsync / CancelAsync are no-ops in the interpreter but should not crash
        bgw.RunWorkerAsync()
        Pass("BackgroundWorker.RunWorkerAsync does not crash")
        bgw.CancelAsync()
        Pass("BackgroundWorker.CancelAsync does not crash")

        ' ===== HelpProvider =====
        Dim hp As New System.Windows.Forms.HelpProvider()
        Check(hp.HelpNamespace = "", "HelpProvider.HelpNamespace default is empty")

        hp.HelpNamespace = "MyApp.chm"
        Check(hp.HelpNamespace = "MyApp.chm", "HelpProvider.HelpNamespace set")

        ' SetHelpString / GetHelpString
        Dim btn As New System.Windows.Forms.Button()
        btn.Name = "btn1"
        hp.SetHelpString(btn, "Click this to save")
        Pass("HelpProvider.SetHelpString does not crash")

        ' ===== PrintDialog =====
        Dim pd As New System.Windows.Forms.PrintDialog()
        Check(pd.AllowPrintToFile = True, "PrintDialog.AllowPrintToFile default is True")
        Check(pd.AllowSelection = False, "PrintDialog.AllowSelection default is False")
        Check(pd.AllowSomePages = False, "PrintDialog.AllowSomePages default is False")
        Check(pd.PrintToFile = False, "PrintDialog.PrintToFile default is False")
        Check(pd.ShowHelp = False, "PrintDialog.ShowHelp default is False")
        Check(pd.ShowNetwork = True, "PrintDialog.ShowNetwork default is True")

        pd.AllowSomePages = True
        Check(pd.AllowSomePages = True, "PrintDialog.AllowSomePages set to True")

        ' ShowDialog in headless mode returns Cancel (2)
        Dim pdResult As Integer = pd.ShowDialog()
        Check(pdResult = 2, "PrintDialog.ShowDialog returns DialogResult.Cancel in headless mode")

        ' ===== PrintPreviewDialog =====
        Dim ppd As New System.Windows.Forms.PrintPreviewDialog()
        Pass("PrintPreviewDialog created successfully")
        Dim ppdResult As Integer = ppd.ShowDialog()
        Check(ppdResult = 2, "PrintPreviewDialog.ShowDialog returns DialogResult.Cancel in headless mode")

        ' ===== PageSetupDialog =====
        Dim psd As New System.Windows.Forms.PageSetupDialog()
        Check(psd.AllowMargins = True, "PageSetupDialog.AllowMargins default is True")
        Check(psd.AllowOrientation = True, "PageSetupDialog.AllowOrientation default is True")
        Check(psd.AllowPaper = True, "PageSetupDialog.AllowPaper default is True")
        Check(psd.AllowPrinter = False, "PageSetupDialog.AllowPrinter default is False")

        psd.AllowPrinter = True
        Check(psd.AllowPrinter = True, "PageSetupDialog.AllowPrinter set to True")

        Dim psdResult As Integer = psd.ShowDialog()
        Check(psdResult = 2, "PageSetupDialog.ShowDialog returns DialogResult.Cancel in headless mode")

        ' ===== PropertyGrid =====
        Dim pg As New System.Windows.Forms.PropertyGrid()
        Check(pg.Text = "", "PropertyGrid.Text default is empty")
        Check(pg.Visible = True, "PropertyGrid.Visible default is True")
        Check(pg.Enabled = True, "PropertyGrid.Enabled default is True")

        pg.SelectedObject = btn
        Pass("PropertyGrid.SelectedObject setter does not crash")

        ' ===== Splitter =====
        Dim spl As New System.Windows.Forms.Splitter()
        Check(spl.Visible = True, "Splitter.Visible default is True")
        Check(spl.Enabled = True, "Splitter.Enabled default is True")

        spl.Dock = 4  ' Right
        Check(spl.Dock = 4, "Splitter.Dock set to Right")

        ' ===== DataGrid (legacy) =====
        Dim dg As New System.Windows.Forms.DataGrid()
        Check(dg.Visible = True, "DataGrid.Visible default is True")
        Check(dg.Enabled = True, "DataGrid.Enabled default is True")
        Check(dg.Text = "", "DataGrid.Text default is empty")

        ' ===== UserControl =====
        Dim uc As New System.Windows.Forms.UserControl()
        Check(uc.Visible = True, "UserControl.Visible default is True")
        Check(uc.Enabled = True, "UserControl.Enabled default is True")
        Check(uc.BackColor = "", "UserControl.BackColor default is empty")

        ' ===== ToolStrip sub-items =====
        ' ToolStripButton
        Dim tsb As New System.Windows.Forms.ToolStripButton()
        Check(tsb.Text = "", "ToolStripButton.Text default is empty")
        Check(tsb.Enabled = True, "ToolStripButton.Enabled default is True")
        Check(tsb.Visible = True, "ToolStripButton.Visible default is True")
        tsb.Text = "Save"
        Check(tsb.Text = "Save", "ToolStripButton.Text set to Save")
        tsb.ToolTipText = "Save the file"
        Check(tsb.ToolTipText = "Save the file", "ToolStripButton.ToolTipText set")

        ' ToolStripLabel
        Dim tsl As New System.Windows.Forms.ToolStripLabel()
        Check(tsl.Text = "", "ToolStripLabel.Text default is empty")
        tsl.Text = "Status:"
        Check(tsl.Text = "Status:", "ToolStripLabel.Text set to Status:")

        ' ToolStripSeparator
        Dim tss As New System.Windows.Forms.ToolStripSeparator()
        Check(tss.Visible = True, "ToolStripSeparator.Visible default is True")

        ' ToolStripComboBox
        Dim tscb As New System.Windows.Forms.ToolStripComboBox()
        Check(tscb.Text = "", "ToolStripComboBox.Text default is empty")
        Check(tscb.Items IsNot Nothing, "ToolStripComboBox.Items not Nothing")
        tscb.Items.Add("Option A")
        tscb.Items.Add("Option B")
        Check(tscb.Items.Count = 2, "ToolStripComboBox.Items.Count after Add is 2")

        ' ToolStripTextBox
        Dim tstb As New System.Windows.Forms.ToolStripTextBox()
        Check(tstb.Text = "", "ToolStripTextBox.Text default is empty")
        tstb.Text = "search..."
        Check(tstb.Text = "search...", "ToolStripTextBox.Text set")

        ' ToolStripProgressBar
        Dim tspb As New System.Windows.Forms.ToolStripProgressBar()
        Check(tspb.Value = 0, "ToolStripProgressBar.Value default is 0")
        Check(tspb.Minimum = 0, "ToolStripProgressBar.Minimum default is 0")
        Check(tspb.Maximum = 100, "ToolStripProgressBar.Maximum default is 100")
        tspb.Value = 50
        Check(tspb.Value = 50, "ToolStripProgressBar.Value set to 50")

        ' ToolStripDropDownButton
        Dim tsddb As New System.Windows.Forms.ToolStripDropDownButton()
        Check(tsddb.Text = "", "ToolStripDropDownButton.Text default is empty")
        tsddb.Text = "File"
        Check(tsddb.Text = "File", "ToolStripDropDownButton.Text set to File")

        ' ToolStripSplitButton
        Dim tssb2 As New System.Windows.Forms.ToolStripSplitButton()
        Check(tssb2.Text = "", "ToolStripSplitButton.Text default is empty")
        tssb2.Text = "New"
        Check(tssb2.Text = "New", "ToolStripSplitButton.Text set to New")

        ' ===== SqlConnection / OleDbConnection (non-visual, just check construction) =====
        Dim sqlConn As New System.Data.SqlClient.SqlConnection()
        sqlConn.ConnectionString = "Server=localhost;Database=test;Uid=root;Pwd=;"
        Check(sqlConn.ConnectionString = "Server=localhost;Database=test;Uid=root;Pwd=;", "SqlConnection.ConnectionString set")
        Pass("SqlConnection created successfully")

        Dim oleConn As New System.Data.OleDb.OleDbConnection()
        oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb"
        Check(oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb", "OleDbConnection.ConnectionString set")
        Pass("OleDbConnection created successfully")

        ' ===== PrintPreviewControl =====
        Dim ppc As New System.Windows.Forms.PrintPreviewControl()
        Check(ppc.Visible = True, "PrintPreviewControl.Visible default is True")
        Check(ppc.Enabled = True, "PrintPreviewControl.Enabled default is True")
        Pass("PrintPreviewControl created successfully")

    End Sub

End Module
