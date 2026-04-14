' Test: All Controls Complete - Properties, Methods, Events
' Comprehensive test covering EVERY control type with all default properties.

Imports System.Windows.Forms

Module TestAllControlsComplete
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Check(description As String, actual As Object, expected As Object)
        If actual.ToString() = expected.ToString() Then
            Console.WriteLine("PASS: " & description)
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: " & description & " (expected=" & expected.ToString() & " actual=" & actual.ToString() & ")")
            failed = failed + 1
        End If
    End Sub

    Sub Main()
        TestButton()
        TestLabel()
        TestCheckBox()
        TestRadioButton()
        TestGroupBox()
        TestPanel()
        TestPictureBox()
        TestTextBox()
        TestComboBox()
        TestListBox()
        TestRichTextBox()
        TestProgressBar()
        TestNumericUpDown()
        TestTreeView()
        TestListView()
        TestDataGridView()
        TestTabControl()
        TestTabPage()
        TestMenuStrip()
        TestStatusStrip()
        TestToolStripStatusLabel()
        TestToolStripMenuItem()
        TestDateTimePicker()
        TestLinkLabel()
        TestToolStrip()
        TestTrackBar()
        TestMaskedTextBox()
        TestSplitContainer()
        TestFlowLayoutPanel()
        TestTableLayoutPanel()
        TestMonthCalendar()
        TestHScrollBar()
        TestVScrollBar()
        TestToolTip()
        TestWebBrowser()

        Console.WriteLine("")
        Console.WriteLine("=== FINAL RESULTS ===")
        Console.WriteLine("Passed: " & passed.ToString())
        Console.WriteLine("Failed: " & failed.ToString())
        Console.WriteLine("Total:  " & (passed + failed).ToString())
        If failed = 0 Then
            Console.WriteLine("ALL CONTROLS COMPLETE TEST PASSED")
        End If
    End Sub

    ' ===== BUTTON =====
    Sub TestButton()
        Console.WriteLine("--- Button ---")
        Dim c As New Button()
        Check("Button.FlatStyle", c.FlatStyle, "Standard")
        Check("Button.DialogResult", c.DialogResult, 0)
        Check("Button.AutoSize", c.AutoSize, False)
        Check("Button.TextAlign", c.TextAlign, "MiddleCenter")
        Check("Button.AutoEllipsis", c.AutoEllipsis, False)
        Check("Button.UseMnemonic", c.UseMnemonic, True)
        c.Text = "OK"
        Check("Button.Text set", c.Text, "OK")
        c.DialogResult = 1
        Check("Button.DialogResult set", c.DialogResult, 1)
        c.Enabled = False
        Check("Button.Enabled set", c.Enabled, False)
    End Sub

    ' ===== LABEL =====
    Sub TestLabel()
        Console.WriteLine("--- Label ---")
        Dim c As New Label()
        Check("Label.AutoSize", c.AutoSize, True)
        Check("Label.TextAlign", c.TextAlign, "TopLeft")
        Check("Label.FlatStyle", c.FlatStyle, "Standard")
        Check("Label.BorderStyle", c.BorderStyle, "None")
        Check("Label.AutoEllipsis", c.AutoEllipsis, False)
        Check("Label.UseMnemonic", c.UseMnemonic, True)
        c.Text = "Hello"
        Check("Label.Text set", c.Text, "Hello")
    End Sub

    ' ===== CHECKBOX =====
    Sub TestCheckBox()
        Console.WriteLine("--- CheckBox ---")
        Dim c As New CheckBox()
        Check("CheckBox.Checked", c.Checked, False)
        Check("CheckBox.CheckState", c.CheckState, "Unchecked")
        Check("CheckBox.ThreeState", c.ThreeState, False)
        Check("CheckBox.AutoCheck", c.AutoCheck, True)
        Check("CheckBox.CheckAlign", c.CheckAlign, "MiddleLeft")
        Check("CheckBox.FlatStyle", c.FlatStyle, "Standard")
        Check("CheckBox.Appearance", c.Appearance, "Normal")
        Check("CheckBox.AutoSize", c.AutoSize, True)
        c.Checked = True
        Check("CheckBox.Checked set", c.Checked, True)
        c.ThreeState = True
        Check("CheckBox.ThreeState set", c.ThreeState, True)
    End Sub

    ' ===== RADIOBUTTON =====
    Sub TestRadioButton()
        Console.WriteLine("--- RadioButton ---")
        Dim c As New RadioButton()
        Check("RadioButton.Checked", c.Checked, False)
        Check("RadioButton.AutoCheck", c.AutoCheck, True)
        Check("RadioButton.CheckAlign", c.CheckAlign, "MiddleLeft")
        Check("RadioButton.FlatStyle", c.FlatStyle, "Standard")
        Check("RadioButton.Appearance", c.Appearance, "Normal")
        Check("RadioButton.AutoSize", c.AutoSize, True)
        c.Checked = True
        Check("RadioButton.Checked set", c.Checked, True)
    End Sub

    ' ===== GROUPBOX =====
    Sub TestGroupBox()
        Console.WriteLine("--- GroupBox ---")
        Dim c As New GroupBox()
        Check("GroupBox.FlatStyle", c.FlatStyle, "Standard")
        Check("GroupBox.AutoSize", c.AutoSize, False)
        Check("GroupBox.Padding", c.Padding, 3)
        c.Text = "Settings"
        Check("GroupBox.Text set", c.Text, "Settings")
    End Sub

    ' ===== PANEL =====
    Sub TestPanel()
        Console.WriteLine("--- Panel ---")
        Dim c As New Panel()
        Check("Panel.BorderStyle", c.BorderStyle, "None")
        Check("Panel.AutoSize", c.AutoSize, False)
        Check("Panel.AutoScroll", c.AutoScroll, False)
        Check("Panel.Padding", c.Padding, 0)
        c.BorderStyle = "FixedSingle"
        Check("Panel.BorderStyle set", c.BorderStyle, "FixedSingle")
    End Sub

    ' ===== PICTUREBOX =====
    Sub TestPictureBox()
        Console.WriteLine("--- PictureBox ---")
        Dim c As New PictureBox()
        Check("PictureBox.SizeMode", c.SizeMode, "Normal")
        Check("PictureBox.BorderStyle", c.BorderStyle, "None")
        Check("PictureBox.WaitOnLoad", c.WaitOnLoad, False)
        c.SizeMode = "StretchImage"
        Check("PictureBox.SizeMode set", c.SizeMode, "StretchImage")
    End Sub

    ' ===== TEXTBOX =====
    Sub TestTextBox()
        Console.WriteLine("--- TextBox ---")
        Dim c As New TextBox()
        Check("TextBox.ReadOnly", c.ReadOnly, False)
        Check("TextBox.Multiline", c.Multiline, False)
        Check("TextBox.MaxLength", c.MaxLength, 32767)
        Check("TextBox.WordWrap", c.WordWrap, True)
        Check("TextBox.AcceptsReturn", c.AcceptsReturn, False)
        Check("TextBox.AcceptsTab", c.AcceptsTab, False)
        Check("TextBox.CharacterCasing", c.CharacterCasing, "Normal")
        Check("TextBox.SelectionStart", c.SelectionStart, 0)
        Check("TextBox.HideSelection", c.HideSelection, True)
        Check("TextBox.BorderStyle", c.BorderStyle, "Fixed3D")
        Check("TextBox.PlaceholderText", c.PlaceholderText, "")
        c.Text = "Hello World"
        Check("TextBox.Text set", c.Text, "Hello World")
        c.ReadOnly = True
        Check("TextBox.ReadOnly set", c.ReadOnly, True)
        c.Multiline = True
        Check("TextBox.Multiline set", c.Multiline, True)
    End Sub

    ' ===== COMBOBOX =====
    Sub TestComboBox()
        Console.WriteLine("--- ComboBox ---")
        Dim c As New ComboBox()
        Check("ComboBox.SelectedIndex", c.SelectedIndex, -1)
        Check("ComboBox.DropDownStyle", c.DropDownStyle, 0)
        Check("ComboBox.MaxDropDownItems", c.MaxDropDownItems, 8)
        Check("ComboBox.DropDownWidth", c.DropDownWidth, 121)
        Check("ComboBox.Sorted", c.Sorted, False)
        Check("ComboBox.FlatStyle", c.FlatStyle, "Standard")
        Check("ComboBox.AutoCompleteMode", c.AutoCompleteMode, 0)
        c.Items.Add("A")
        c.Items.Add("B")
        Check("ComboBox.Items.Count", c.Items.Count, 2)
    End Sub

    ' ===== LISTBOX =====
    Sub TestListBox()
        Console.WriteLine("--- ListBox ---")
        Dim c As New ListBox()
        Check("ListBox.SelectedIndex", c.SelectedIndex, -1)
        Check("ListBox.SelectionMode", c.SelectionMode, 1)
        Check("ListBox.Sorted", c.Sorted, False)
        Check("ListBox.IntegralHeight", c.IntegralHeight, True)
        Check("ListBox.MultiColumn", c.MultiColumn, False)
        Check("ListBox.HorizontalScrollbar", c.HorizontalScrollbar, False)
        c.Items.Add("X")
        c.Items.Add("Y")
        c.Items.Add("Z")
        Check("ListBox.Items.Count", c.Items.Count, 3)
    End Sub

    ' ===== RICHTEXTBOX =====
    Sub TestRichTextBox()
        Console.WriteLine("--- RichTextBox ---")
        Dim c As New RichTextBox()
        Check("RichTextBox.ReadOnly", c.ReadOnly, False)
        Check("RichTextBox.Multiline", c.Multiline, True)
        Check("RichTextBox.WordWrap", c.WordWrap, True)
        Check("RichTextBox.DetectUrls", c.DetectUrls, True)
        Check("RichTextBox.HideSelection", c.HideSelection, True)
        Check("RichTextBox.AcceptsTab", c.AcceptsTab, False)
        Check("RichTextBox.SelectionStart", c.SelectionStart, 0)
        Check("RichTextBox.Modified", c.Modified, False)
        c.Text = "Rich text"
        Check("RichTextBox.Text set", c.Text, "Rich text")
    End Sub

    ' ===== PROGRESSBAR =====
    Sub TestProgressBar()
        Console.WriteLine("--- ProgressBar ---")
        Dim c As New ProgressBar()
        Check("ProgressBar.Value", c.Value, 0)
        Check("ProgressBar.Minimum", c.Minimum, 0)
        Check("ProgressBar.Maximum", c.Maximum, 100)
        Check("ProgressBar.Step", c.Step, 10)
        Check("ProgressBar.Style", c.Style, 0)
        c.Value = 50
        Check("ProgressBar.Value set", c.Value, 50)
        c.Maximum = 200
        Check("ProgressBar.Maximum set", c.Maximum, 200)
    End Sub

    ' ===== NUMERICUPDOWN =====
    Sub TestNumericUpDown()
        Console.WriteLine("--- NumericUpDown ---")
        Dim c As New NumericUpDown()
        Check("NumericUpDown.Value", c.Value, 0)
        Check("NumericUpDown.Minimum", c.Minimum, 0)
        Check("NumericUpDown.Maximum", c.Maximum, 100)
        Check("NumericUpDown.Increment", c.Increment, 1)
        Check("NumericUpDown.DecimalPlaces", c.DecimalPlaces, 0)
        Check("NumericUpDown.ReadOnly", c.ReadOnly, False)
        Check("NumericUpDown.Hexadecimal", c.Hexadecimal, False)
        Check("NumericUpDown.ThousandsSeparator", c.ThousandsSeparator, False)
        Check("NumericUpDown.TextAlign", c.TextAlign, "Left")
        c.Value = 42
        Check("NumericUpDown.Value set", c.Value, 42)
    End Sub

    ' ===== TREEVIEW =====
    Sub TestTreeView()
        Console.WriteLine("--- TreeView ---")
        Dim c As New TreeView()
        Check("TreeView.CheckBoxes", c.CheckBoxes, False)
        Check("TreeView.ShowLines", c.ShowLines, True)
        Check("TreeView.ShowRootLines", c.ShowRootLines, True)
        Check("TreeView.ShowPlusMinus", c.ShowPlusMinus, True)
        Check("TreeView.FullRowSelect", c.FullRowSelect, False)
        Check("TreeView.HideSelection", c.HideSelection, True)
        Check("TreeView.LabelEdit", c.LabelEdit, False)
        Check("TreeView.Scrollable", c.Scrollable, True)
        Check("TreeView.Sorted", c.Sorted, False)
        Check("TreeView.Indent", c.Indent, 19)
        Check("TreeView.ItemHeight", c.ItemHeight, 16)
        Check("TreeView.PathSeparator", c.PathSeparator, "\")
    End Sub

    ' ===== LISTVIEW =====
    Sub TestListView()
        Console.WriteLine("--- ListView ---")
        Dim c As New ListView()
        Check("ListView.View", c.View, 1)
        Check("ListView.FullRowSelect", c.FullRowSelect, False)
        Check("ListView.GridLines", c.GridLines, False)
        Check("ListView.CheckBoxes", c.CheckBoxes, False)
        Check("ListView.MultiSelect", c.MultiSelect, True)
        Check("ListView.ShowGroups", c.ShowGroups, False)
        Check("ListView.Sorting", c.Sorting, "None")
        Check("ListView.LabelEdit", c.LabelEdit, False)
        Check("ListView.LabelWrap", c.LabelWrap, True)
        Check("ListView.AllowColumnReorder", c.AllowColumnReorder, False)
        Check("ListView.HideSelection", c.HideSelection, True)
        Check("ListView.Scrollable", c.Scrollable, True)
    End Sub

    ' ===== DATAGRIDVIEW =====
    Sub TestDataGridView()
        Console.WriteLine("--- DataGridView ---")
        Dim c As New DataGridView()
        Check("DataGridView.AllowUserToAddRows", c.AllowUserToAddRows, True)
        Check("DataGridView.AllowUserToDeleteRows", c.AllowUserToDeleteRows, True)
        Check("DataGridView.ReadOnly", c.ReadOnly, False)
        Check("DataGridView.AutoGenerateColumns", c.AutoGenerateColumns, True)
        Check("DataGridView.MultiSelect", c.MultiSelect, True)
        Check("DataGridView.SelectionMode", c.SelectionMode, "RowHeaderSelect")
        Check("DataGridView.ColumnHeadersVisible", c.ColumnHeadersVisible, True)
        Check("DataGridView.RowHeadersVisible", c.RowHeadersVisible, True)
        Check("DataGridView.RowHeadersWidth", c.RowHeadersWidth, 43)
        Check("DataGridView.EditMode", c.EditMode, "EditOnKeystrokeOrF2")
        Check("DataGridView.BorderStyle", c.BorderStyle, "FixedSingle")
        Check("DataGridView.CellBorderStyle", c.CellBorderStyle, "Single")
        Check("DataGridView.AllowUserToResizeColumns", c.AllowUserToResizeColumns, True)
        Check("DataGridView.AllowUserToOrderColumns", c.AllowUserToOrderColumns, False)
    End Sub

    ' ===== TABCONTROL =====
    Sub TestTabControl()
        Console.WriteLine("--- TabControl ---")
        Dim c As New TabControl()
        Check("TabControl.SelectedIndex", c.SelectedIndex, 0)
        Check("TabControl.Alignment", c.Alignment, "Top")
        Check("TabControl.Appearance", c.Appearance, "Normal")
        Check("TabControl.Multiline", c.Multiline, False)
        Check("TabControl.SizeMode", c.SizeMode, "Normal")
        Check("TabControl.HotTrack", c.HotTrack, False)
        Check("TabControl.ShowToolTips", c.ShowToolTips, False)
    End Sub

    ' ===== TABPAGE =====
    Sub TestTabPage()
        Console.WriteLine("--- TabPage ---")
        Dim c As New TabPage()
        Check("TabPage.BorderStyle", c.BorderStyle, "None")
        Check("TabPage.AutoSize", c.AutoSize, False)
        Check("TabPage.AutoScroll", c.AutoScroll, False)
        Check("TabPage.UseVisualStyleBackColor", c.UseVisualStyleBackColor, True)
        Check("TabPage.Padding", c.Padding, 3)
        c.Text = "Tab1"
        Check("TabPage.Text set", c.Text, "Tab1")
    End Sub

    ' ===== MENUSTRIP =====
    Sub TestMenuStrip()
        Console.WriteLine("--- MenuStrip ---")
        Dim c As New MenuStrip()
        Check("MenuStrip.Dock", c.Dock, 1)
        Check("MenuStrip.RenderMode", c.RenderMode, "ManagerRenderMode")
        Check("MenuStrip.Stretch", c.Stretch, True)
    End Sub

    ' ===== STATUSSTRIP =====
    Sub TestStatusStrip()
        Console.WriteLine("--- StatusStrip ---")
        Dim c As New StatusStrip()
        Check("StatusStrip.Dock", c.Dock, 2)
        Check("StatusStrip.RenderMode", c.RenderMode, "ManagerRenderMode")
        Check("StatusStrip.SizingGrip", c.SizingGrip, True)
        Check("StatusStrip.Stretch", c.Stretch, True)
    End Sub

    ' ===== TOOLSTRIPSTATUSLABEL =====
    Sub TestToolStripStatusLabel()
        Console.WriteLine("--- ToolStripStatusLabel ---")
        Dim c As New ToolStripStatusLabel()
        Check("ToolStripStatusLabel.Spring", c.Spring, False)
        Check("ToolStripStatusLabel.AutoSize", c.AutoSize, True)
        Check("ToolStripStatusLabel.BorderSides", c.BorderSides, "None")
        Check("ToolStripStatusLabel.BorderStyle", c.BorderStyle, "None")
        Check("ToolStripStatusLabel.IsLink", c.IsLink, False)
        Check("ToolStripStatusLabel.Alignment", c.Alignment, "Left")
        c.Text = "Ready"
        Check("ToolStripStatusLabel.Text set", c.Text, "Ready")
    End Sub

    ' ===== TOOLSTRIPMENUITEM =====
    Sub TestToolStripMenuItem()
        Console.WriteLine("--- ToolStripMenuItem ---")
        Dim c As New ToolStripMenuItem()
        Check("ToolStripMenuItem.Checked", c.Checked, False)
        Check("ToolStripMenuItem.CheckState", c.CheckState, "Unchecked")
        Check("ToolStripMenuItem.CheckOnClick", c.CheckOnClick, False)
        Check("ToolStripMenuItem.ShowShortcutKeys", c.ShowShortcutKeys, True)
        Check("ToolStripMenuItem.AutoSize", c.AutoSize, True)
        Check("ToolStripMenuItem.DisplayStyle", c.DisplayStyle, "ImageAndText")
        Check("ToolStripMenuItem.Alignment", c.Alignment, "Left")
        c.Text = "File"
        Check("ToolStripMenuItem.Text set", c.Text, "File")
        c.Checked = True
        Check("ToolStripMenuItem.Checked set", c.Checked, True)
    End Sub

    ' ===== DATETIMEPICKER =====
    Sub TestDateTimePicker()
        Console.WriteLine("--- DateTimePicker ---")
        Dim c As New DateTimePicker()
        Check("DateTimePicker.Format", c.Format, "Long")
        Check("DateTimePicker.ShowCheckBox", c.ShowCheckBox, False)
        Check("DateTimePicker.Checked", c.Checked, True)
        Check("DateTimePicker.ShowUpDown", c.ShowUpDown, False)
        Check("DateTimePicker.DropDownAlign", c.DropDownAlign, "Left")
        Check("DateTimePicker.MinDate", c.MinDate, "1/1/1753")
        Check("DateTimePicker.MaxDate", c.MaxDate, "12/31/9998")
        c.Format = "Short"
        Check("DateTimePicker.Format set", c.Format, "Short")
        c.CustomFormat = "yyyy-MM-dd"
        Check("DateTimePicker.CustomFormat set", c.CustomFormat, "yyyy-MM-dd")
    End Sub

    ' ===== LINKLABEL =====
    Sub TestLinkLabel()
        Console.WriteLine("--- LinkLabel ---")
        Dim c As New LinkLabel()
        Check("LinkLabel.LinkColor", c.LinkColor, "#0066cc")
        Check("LinkLabel.VisitedLinkColor", c.VisitedLinkColor, "#800080")
        Check("LinkLabel.ActiveLinkColor", c.ActiveLinkColor, "Red")
        Check("LinkLabel.LinkVisited", c.LinkVisited, False)
        Check("LinkLabel.LinkBehavior", c.LinkBehavior, "SystemDefault")
        Check("LinkLabel.AutoSize", c.AutoSize, True)
        Check("LinkLabel.TextAlign", c.TextAlign, "TopLeft")
        c.Text = "Click here"
        Check("LinkLabel.Text set", c.Text, "Click here")
        c.LinkVisited = True
        Check("LinkLabel.LinkVisited set", c.LinkVisited, True)
    End Sub

    ' ===== TOOLSTRIP =====
    Sub TestToolStrip()
        Console.WriteLine("--- ToolStrip ---")
        Dim c As New ToolStrip()
        Check("ToolStrip.Dock", c.Dock, 1)
        Check("ToolStrip.RenderMode", c.RenderMode, "ManagerRenderMode")
        Check("ToolStrip.Stretch", c.Stretch, True)
        Check("ToolStrip.ShowItemToolTips", c.ShowItemToolTips, True)
        Check("ToolStrip.LayoutStyle", c.LayoutStyle, "HorizontalStackWithOverflow")
    End Sub

    ' ===== TRACKBAR =====
    Sub TestTrackBar()
        Console.WriteLine("--- TrackBar ---")
        Dim c As New TrackBar()
        Check("TrackBar.Value", c.Value, 0)
        Check("TrackBar.Minimum", c.Minimum, 0)
        Check("TrackBar.Maximum", c.Maximum, 10)
        Check("TrackBar.TickFrequency", c.TickFrequency, 1)
        Check("TrackBar.SmallChange", c.SmallChange, 1)
        Check("TrackBar.LargeChange", c.LargeChange, 5)
        Check("TrackBar.Orientation", c.Orientation, "Horizontal")
        Check("TrackBar.TickStyle", c.TickStyle, "BottomRight")
        c.Value = 5
        Check("TrackBar.Value set", c.Value, 5)
        c.Maximum = 20
        Check("TrackBar.Maximum set", c.Maximum, 20)
    End Sub

    ' ===== MASKEDTEXTBOX =====
    Sub TestMaskedTextBox()
        Console.WriteLine("--- MaskedTextBox ---")
        Dim c As New MaskedTextBox()
        Check("MaskedTextBox.PromptChar", c.PromptChar, "_")
        Check("MaskedTextBox.MaskCompleted", c.MaskCompleted, False)
        Check("MaskedTextBox.ReadOnly", c.ReadOnly, False)
        Check("MaskedTextBox.HidePromptOnLeave", c.HidePromptOnLeave, False)
        Check("MaskedTextBox.AsciiOnly", c.AsciiOnly, False)
        Check("MaskedTextBox.SkipLiterals", c.SkipLiterals, True)
        Check("MaskedTextBox.TextAlign", c.TextAlign, "Left")
        c.Mask = "000-00-0000"
        Check("MaskedTextBox.Mask set", c.Mask, "000-00-0000")
    End Sub

    ' ===== SPLITCONTAINER =====
    Sub TestSplitContainer()
        Console.WriteLine("--- SplitContainer ---")
        Dim c As New SplitContainer()
        Check("SplitContainer.Orientation", c.Orientation, "Vertical")
        Check("SplitContainer.SplitterDistance", c.SplitterDistance, 100)
        Check("SplitContainer.SplitterIncrement", c.SplitterIncrement, 1)
        Check("SplitContainer.SplitterWidth", c.SplitterWidth, 4)
        Check("SplitContainer.FixedPanel", c.FixedPanel, "None")
        Check("SplitContainer.IsSplitterFixed", c.IsSplitterFixed, False)
        Check("SplitContainer.Panel1Collapsed", c.Panel1Collapsed, False)
        Check("SplitContainer.Panel2Collapsed", c.Panel2Collapsed, False)
        Check("SplitContainer.Panel1MinSize", c.Panel1MinSize, 25)
        Check("SplitContainer.Panel2MinSize", c.Panel2MinSize, 25)
        Check("SplitContainer.BorderStyle", c.BorderStyle, "None")
        c.SplitterDistance = 200
        Check("SplitContainer.SplitterDistance set", c.SplitterDistance, 200)
    End Sub

    ' ===== FLOWLAYOUTPANEL =====
    Sub TestFlowLayoutPanel()
        Console.WriteLine("--- FlowLayoutPanel ---")
        Dim c As New FlowLayoutPanel()
        Check("FlowLayoutPanel.FlowDirection", c.FlowDirection, "LeftToRight")
        Check("FlowLayoutPanel.WrapContents", c.WrapContents, True)
        Check("FlowLayoutPanel.AutoSize", c.AutoSize, False)
        Check("FlowLayoutPanel.AutoSizeMode", c.AutoSizeMode, "GrowOnly")
        Check("FlowLayoutPanel.AutoScroll", c.AutoScroll, False)
        Check("FlowLayoutPanel.BorderStyle", c.BorderStyle, "None")
        c.FlowDirection = "TopDown"
        Check("FlowLayoutPanel.FlowDirection set", c.FlowDirection, "TopDown")
    End Sub

    ' ===== TABLELAYOUTPANEL =====
    Sub TestTableLayoutPanel()
        Console.WriteLine("--- TableLayoutPanel ---")
        Dim c As New TableLayoutPanel()
        Check("TableLayoutPanel.ColumnCount", c.ColumnCount, 2)
        Check("TableLayoutPanel.RowCount", c.RowCount, 2)
        Check("TableLayoutPanel.AutoSize", c.AutoSize, False)
        Check("TableLayoutPanel.AutoScroll", c.AutoScroll, False)
        Check("TableLayoutPanel.BorderStyle", c.BorderStyle, "None")
        Check("TableLayoutPanel.CellBorderStyle", c.CellBorderStyle, "None")
        Check("TableLayoutPanel.GrowStyle", c.GrowStyle, "AddRows")
        c.ColumnCount = 3
        Check("TableLayoutPanel.ColumnCount set", c.ColumnCount, 3)
    End Sub

    ' ===== MONTHCALENDAR =====
    Sub TestMonthCalendar()
        Console.WriteLine("--- MonthCalendar ---")
        Dim c As New MonthCalendar()
        Check("MonthCalendar.ShowToday", c.ShowToday, True)
        Check("MonthCalendar.ShowTodayCircle", c.ShowTodayCircle, True)
        Check("MonthCalendar.ShowWeekNumbers", c.ShowWeekNumbers, False)
        Check("MonthCalendar.MaxSelectionCount", c.MaxSelectionCount, 7)
        Check("MonthCalendar.FirstDayOfWeek", c.FirstDayOfWeek, "Default")
        Check("MonthCalendar.ScrollChange", c.ScrollChange, 1)
        c.ShowWeekNumbers = True
        Check("MonthCalendar.ShowWeekNumbers set", c.ShowWeekNumbers, True)
        c.MaxSelectionCount = 14
        Check("MonthCalendar.MaxSelectionCount set", c.MaxSelectionCount, 14)
    End Sub

    ' ===== HSCROLLBAR =====
    Sub TestHScrollBar()
        Console.WriteLine("--- HScrollBar ---")
        Dim c As New HScrollBar()
        Check("HScrollBar.Value", c.Value, 0)
        Check("HScrollBar.Minimum", c.Minimum, 0)
        Check("HScrollBar.Maximum", c.Maximum, 100)
        Check("HScrollBar.SmallChange", c.SmallChange, 1)
        Check("HScrollBar.LargeChange", c.LargeChange, 10)
        c.Value = 50
        Check("HScrollBar.Value set", c.Value, 50)
        c.Maximum = 200
        Check("HScrollBar.Maximum set", c.Maximum, 200)
    End Sub

    ' ===== VSCROLLBAR =====
    Sub TestVScrollBar()
        Console.WriteLine("--- VScrollBar ---")
        Dim c As New VScrollBar()
        Check("VScrollBar.Value", c.Value, 0)
        Check("VScrollBar.Minimum", c.Minimum, 0)
        Check("VScrollBar.Maximum", c.Maximum, 100)
        Check("VScrollBar.SmallChange", c.SmallChange, 1)
        Check("VScrollBar.LargeChange", c.LargeChange, 10)
        c.Value = 75
        Check("VScrollBar.Value set", c.Value, 75)
    End Sub

    ' ===== TOOLTIP =====
    Sub TestToolTip()
        Console.WriteLine("--- ToolTip ---")
        Dim c As New ToolTip()
        Check("ToolTip.Active", c.Active, True)
        Check("ToolTip.AutoPopDelay", c.AutoPopDelay, 5000)
        Check("ToolTip.InitialDelay", c.InitialDelay, 500)
        Check("ToolTip.ReshowDelay", c.ReshowDelay, 100)
        Check("ToolTip.ShowAlways", c.ShowAlways, False)
        Check("ToolTip.UseFading", c.UseFading, True)
        Check("ToolTip.UseAnimation", c.UseAnimation, True)
        Check("ToolTip.ToolTipTitle", c.ToolTipTitle, "")
        ' SetToolTip/GetToolTip
        Dim btn As New Button()
        btn.Name = "btnOK"
        c.SetToolTip(btn, "Press OK")
        Dim tip As String = c.GetToolTip(btn)
        Check("ToolTip.SetToolTip/GetToolTip", tip, "Press OK")
        c.RemoveAll()
        tip = c.GetToolTip(btn)
        Check("ToolTip.RemoveAll", tip, "")
    End Sub

    ' ===== WEBBROWSER =====
    Sub TestWebBrowser()
        Console.WriteLine("--- WebBrowser ---")
        Dim c As New WebBrowser()
        Check("WebBrowser.CanGoBack", c.CanGoBack, False)
        Check("WebBrowser.CanGoForward", c.CanGoForward, False)
        Check("WebBrowser.IsBusy", c.IsBusy, False)
        Check("WebBrowser.ReadyState", c.ReadyState, "Uninitialized")
        Check("WebBrowser.AllowNavigation", c.AllowNavigation, True)
        Check("WebBrowser.ScrollBarsEnabled", c.ScrollBarsEnabled, True)
    End Sub
End Module
