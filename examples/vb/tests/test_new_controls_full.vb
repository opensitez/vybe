' Test: New Controls - MonthCalendar, HScrollBar, VScrollBar, ToolTip
' Tests creation, property access/set, events, and methods for all new + enhanced controls.

Imports System.Windows.Forms

Module TestNewControlsFull
    Sub Main()
        Dim passed As Integer = 0
        Dim failed As Integer = 0

        ' ===== MonthCalendar =====
        Console.WriteLine("=== MonthCalendar Tests ===")
        
        Dim cal As New MonthCalendar()
        ' Default properties
        If cal.ShowToday = True Then
            Console.WriteLine("PASS: MonthCalendar.ShowToday default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.ShowToday default")
            failed = failed + 1
        End If
        
        If cal.ShowTodayCircle = True Then
            Console.WriteLine("PASS: MonthCalendar.ShowTodayCircle default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.ShowTodayCircle default")
            failed = failed + 1
        End If
        
        If cal.ShowWeekNumbers = False Then
            Console.WriteLine("PASS: MonthCalendar.ShowWeekNumbers default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.ShowWeekNumbers default")
            failed = failed + 1
        End If
        
        If cal.MaxSelectionCount = 7 Then
            Console.WriteLine("PASS: MonthCalendar.MaxSelectionCount default is 7")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.MaxSelectionCount default")
            failed = failed + 1
        End If
        
        If cal.FirstDayOfWeek = "Default" Then
            Console.WriteLine("PASS: MonthCalendar.FirstDayOfWeek default is Default")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.FirstDayOfWeek default")
            failed = failed + 1
        End If

        If cal.ScrollChange = 1 Then
            Console.WriteLine("PASS: MonthCalendar.ScrollChange default is 1")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.ScrollChange default")
            failed = failed + 1
        End If

        ' Set properties
        cal.ShowWeekNumbers = True
        If cal.ShowWeekNumbers = True Then
            Console.WriteLine("PASS: MonthCalendar.ShowWeekNumbers set to True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.ShowWeekNumbers set")
            failed = failed + 1
        End If

        cal.MaxSelectionCount = 14
        If cal.MaxSelectionCount = 14 Then
            Console.WriteLine("PASS: MonthCalendar.MaxSelectionCount set to 14")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MonthCalendar.MaxSelectionCount set")
            failed = failed + 1
        End If

        ' ===== HScrollBar =====
        Console.WriteLine("")
        Console.WriteLine("=== HScrollBar Tests ===")

        Dim hbar As New HScrollBar()
        If hbar.Value = 0 Then
            Console.WriteLine("PASS: HScrollBar.Value default is 0")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: HScrollBar.Value default")
            failed = failed + 1
        End If

        If hbar.Minimum = 0 Then
            Console.WriteLine("PASS: HScrollBar.Minimum default is 0")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: HScrollBar.Minimum default")
            failed = failed + 1
        End If

        If hbar.Maximum = 100 Then
            Console.WriteLine("PASS: HScrollBar.Maximum default is 100")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: HScrollBar.Maximum default")
            failed = failed + 1
        End If

        If hbar.SmallChange = 1 Then
            Console.WriteLine("PASS: HScrollBar.SmallChange default is 1")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: HScrollBar.SmallChange default")
            failed = failed + 1
        End If

        If hbar.LargeChange = 10 Then
            Console.WriteLine("PASS: HScrollBar.LargeChange default is 10")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: HScrollBar.LargeChange default")
            failed = failed + 1
        End If

        hbar.Value = 50
        hbar.Minimum = 10
        hbar.Maximum = 200
        hbar.SmallChange = 5
        hbar.LargeChange = 20

        If hbar.Value = 50 Then
            Console.WriteLine("PASS: HScrollBar.Value set to 50")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: HScrollBar.Value set")
            failed = failed + 1
        End If

        If hbar.Maximum = 200 Then
            Console.WriteLine("PASS: HScrollBar.Maximum set to 200")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: HScrollBar.Maximum set")
            failed = failed + 1
        End If

        ' ===== VScrollBar =====
        Console.WriteLine("")
        Console.WriteLine("=== VScrollBar Tests ===")

        Dim vbar As New VScrollBar()
        If vbar.Value = 0 Then
            Console.WriteLine("PASS: VScrollBar.Value default is 0")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: VScrollBar.Value default")
            failed = failed + 1
        End If

        If vbar.Maximum = 100 Then
            Console.WriteLine("PASS: VScrollBar.Maximum default is 100")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: VScrollBar.Maximum default")
            failed = failed + 1
        End If

        If vbar.LargeChange = 10 Then
            Console.WriteLine("PASS: VScrollBar.LargeChange default is 10")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: VScrollBar.LargeChange default")
            failed = failed + 1
        End If

        vbar.Value = 75
        If vbar.Value = 75 Then
            Console.WriteLine("PASS: VScrollBar.Value set to 75")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: VScrollBar.Value set")
            failed = failed + 1
        End If

        ' ===== ToolTip =====
        Console.WriteLine("")
        Console.WriteLine("=== ToolTip Tests ===")

        Dim tt As New ToolTip()
        If tt.Active = True Then
            Console.WriteLine("PASS: ToolTip.Active default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.Active default")
            failed = failed + 1
        End If

        If tt.AutoPopDelay = 5000 Then
            Console.WriteLine("PASS: ToolTip.AutoPopDelay default is 5000")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.AutoPopDelay default")
            failed = failed + 1
        End If

        If tt.InitialDelay = 500 Then
            Console.WriteLine("PASS: ToolTip.InitialDelay default is 500")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.InitialDelay default")
            failed = failed + 1
        End If

        If tt.ReshowDelay = 100 Then
            Console.WriteLine("PASS: ToolTip.ReshowDelay default is 100")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.ReshowDelay default")
            failed = failed + 1
        End If

        If tt.ShowAlways = False Then
            Console.WriteLine("PASS: ToolTip.ShowAlways default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.ShowAlways default")
            failed = failed + 1
        End If

        If tt.UseFading = True Then
            Console.WriteLine("PASS: ToolTip.UseFading default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.UseFading default")
            failed = failed + 1
        End If

        If tt.UseAnimation = True Then
            Console.WriteLine("PASS: ToolTip.UseAnimation default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.UseAnimation default")
            failed = failed + 1
        End If

        ' Set properties
        tt.AutoPopDelay = 10000
        If tt.AutoPopDelay = 10000 Then
            Console.WriteLine("PASS: ToolTip.AutoPopDelay set to 10000")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.AutoPopDelay set")
            failed = failed + 1
        End If

        ' ToolTip methods
        Dim btn As New Button()
        btn.Name = "btnTest"
        tt.SetToolTip(btn, "Click me!")
        Dim tipText As String = tt.GetToolTip(btn)
        If tipText = "Click me!" Then
            Console.WriteLine("PASS: ToolTip.SetToolTip/GetToolTip works")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.SetToolTip/GetToolTip: got '" & tipText & "'")
            failed = failed + 1
        End If

        ' RemoveAll
        tt.RemoveAll()
        tipText = tt.GetToolTip(btn)
        If tipText = "" Then
            Console.WriteLine("PASS: ToolTip.RemoveAll clears tooltips")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ToolTip.RemoveAll: got '" & tipText & "'")
            failed = failed + 1
        End If

        ' ===== Enhanced Existing Controls =====
        Console.WriteLine("")
        Console.WriteLine("=== Enhanced Control Properties ===")

        ' TextBox enhancements
        Dim tb As New TextBox()
        If tb.AcceptsReturn = False Then
            Console.WriteLine("PASS: TextBox.AcceptsReturn default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TextBox.AcceptsReturn default")
            failed = failed + 1
        End If

        If tb.AcceptsTab = False Then
            Console.WriteLine("PASS: TextBox.AcceptsTab default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TextBox.AcceptsTab default")
            failed = failed + 1
        End If

        If tb.CharacterCasing = "Normal" Then
            Console.WriteLine("PASS: TextBox.CharacterCasing default is Normal")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TextBox.CharacterCasing default")
            failed = failed + 1
        End If

        If tb.SelectionStart = 0 Then
            Console.WriteLine("PASS: TextBox.SelectionStart default is 0")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TextBox.SelectionStart default")
            failed = failed + 1
        End If

        If tb.HideSelection = True Then
            Console.WriteLine("PASS: TextBox.HideSelection default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TextBox.HideSelection default")
            failed = failed + 1
        End If

        ' RichTextBox enhancements
        Dim rtb As New RichTextBox()
        If rtb.Multiline = True Then
            Console.WriteLine("PASS: RichTextBox.Multiline default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: RichTextBox.Multiline default")
            failed = failed + 1
        End If

        If rtb.DetectUrls = True Then
            Console.WriteLine("PASS: RichTextBox.DetectUrls default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: RichTextBox.DetectUrls default")
            failed = failed + 1
        End If

        If rtb.ZoomFactor = 1.0 Then
            Console.WriteLine("PASS: RichTextBox.ZoomFactor default is 1.0")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: RichTextBox.ZoomFactor default")
            failed = failed + 1
        End If

        ' DateTimePicker enhancements
        Dim dtp As New DateTimePicker()
        If dtp.ShowUpDown = False Then
            Console.WriteLine("PASS: DateTimePicker.ShowUpDown default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: DateTimePicker.ShowUpDown default")
            failed = failed + 1
        End If

        If dtp.DropDownAlign = "Left" Then
            Console.WriteLine("PASS: DateTimePicker.DropDownAlign default is Left")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: DateTimePicker.DropDownAlign default")
            failed = failed + 1
        End If

        ' LinkLabel enhancements
        Dim ll As New LinkLabel()
        If ll.ActiveLinkColor = "Red" Then
            Console.WriteLine("PASS: LinkLabel.ActiveLinkColor default is Red")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: LinkLabel.ActiveLinkColor default")
            failed = failed + 1
        End If

        If ll.LinkBehavior = "SystemDefault" Then
            Console.WriteLine("PASS: LinkLabel.LinkBehavior default is SystemDefault")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: LinkLabel.LinkBehavior default")
            failed = failed + 1
        End If

        If ll.AutoSize = True Then
            Console.WriteLine("PASS: LinkLabel.AutoSize default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: LinkLabel.AutoSize default")
            failed = failed + 1
        End If

        ' TrackBar enhancements
        Dim tr As New TrackBar()
        If tr.TickStyle = "BottomRight" Then
            Console.WriteLine("PASS: TrackBar.TickStyle default is BottomRight")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TrackBar.TickStyle default")
            failed = failed + 1
        End If

        ' MaskedTextBox enhancements
        Dim mtb As New MaskedTextBox()
        If mtb.HidePromptOnLeave = False Then
            Console.WriteLine("PASS: MaskedTextBox.HidePromptOnLeave default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MaskedTextBox.HidePromptOnLeave default")
            failed = failed + 1
        End If

        If mtb.AsciiOnly = False Then
            Console.WriteLine("PASS: MaskedTextBox.AsciiOnly default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MaskedTextBox.AsciiOnly default")
            failed = failed + 1
        End If

        If mtb.SkipLiterals = True Then
            Console.WriteLine("PASS: MaskedTextBox.SkipLiterals default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: MaskedTextBox.SkipLiterals default")
            failed = failed + 1
        End If

        ' SplitContainer enhancements
        Dim sc As New SplitContainer()
        If sc.IsSplitterFixed = False Then
            Console.WriteLine("PASS: SplitContainer.IsSplitterFixed default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: SplitContainer.IsSplitterFixed default")
            failed = failed + 1
        End If

        If sc.Panel1Collapsed = False Then
            Console.WriteLine("PASS: SplitContainer.Panel1Collapsed default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: SplitContainer.Panel1Collapsed default")
            failed = failed + 1
        End If

        If sc.Panel1MinSize = 25 Then
            Console.WriteLine("PASS: SplitContainer.Panel1MinSize default is 25")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: SplitContainer.Panel1MinSize default")
            failed = failed + 1
        End If

        If sc.SplitterWidth = 4 Then
            Console.WriteLine("PASS: SplitContainer.SplitterWidth default is 4")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: SplitContainer.SplitterWidth default")
            failed = failed + 1
        End If

        ' FlowLayoutPanel enhancements
        Dim flp As New FlowLayoutPanel()
        If flp.AutoSize = False Then
            Console.WriteLine("PASS: FlowLayoutPanel.AutoSize default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: FlowLayoutPanel.AutoSize default")
            failed = failed + 1
        End If

        If flp.BorderStyle = "None" Then
            Console.WriteLine("PASS: FlowLayoutPanel.BorderStyle default is None")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: FlowLayoutPanel.BorderStyle default")
            failed = failed + 1
        End If

        ' TableLayoutPanel enhancements
        Dim tlp As New TableLayoutPanel()
        If tlp.CellBorderStyle = "None" Then
            Console.WriteLine("PASS: TableLayoutPanel.CellBorderStyle default is None")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TableLayoutPanel.CellBorderStyle default")
            failed = failed + 1
        End If

        If tlp.GrowStyle = "AddRows" Then
            Console.WriteLine("PASS: TableLayoutPanel.GrowStyle default is AddRows")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TableLayoutPanel.GrowStyle default")
            failed = failed + 1
        End If

        ' ComboBox enhancements
        Dim cb As New ComboBox()
        If cb.Sorted = False Then
            Console.WriteLine("PASS: ComboBox.Sorted default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ComboBox.Sorted default")
            failed = failed + 1
        End If

        If cb.MaxDropDownItems = 8 Then
            Console.WriteLine("PASS: ComboBox.MaxDropDownItems default is 8")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ComboBox.MaxDropDownItems default")
            failed = failed + 1
        End If

        ' ListBox enhancements
        Dim lb As New ListBox()
        If lb.Sorted = False Then
            Console.WriteLine("PASS: ListBox.Sorted default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ListBox.Sorted default")
            failed = failed + 1
        End If

        If lb.IntegralHeight = True Then
            Console.WriteLine("PASS: ListBox.IntegralHeight default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ListBox.IntegralHeight default")
            failed = failed + 1
        End If

        ' NumericUpDown enhancements
        Dim nud As New NumericUpDown()
        If nud.Hexadecimal = False Then
            Console.WriteLine("PASS: NumericUpDown.Hexadecimal default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: NumericUpDown.Hexadecimal default")
            failed = failed + 1
        End If

        If nud.ThousandsSeparator = False Then
            Console.WriteLine("PASS: NumericUpDown.ThousandsSeparator default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: NumericUpDown.ThousandsSeparator default")
            failed = failed + 1
        End If

        ' DataGridView enhancements
        Dim dgv As New DataGridView()
        If dgv.AutoGenerateColumns = True Then
            Console.WriteLine("PASS: DataGridView.AutoGenerateColumns default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: DataGridView.AutoGenerateColumns default")
            failed = failed + 1
        End If

        If dgv.MultiSelect = True Then
            Console.WriteLine("PASS: DataGridView.MultiSelect default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: DataGridView.MultiSelect default")
            failed = failed + 1
        End If

        If dgv.ColumnHeadersVisible = True Then
            Console.WriteLine("PASS: DataGridView.ColumnHeadersVisible default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: DataGridView.ColumnHeadersVisible default")
            failed = failed + 1
        End If

        If dgv.RowHeadersVisible = True Then
            Console.WriteLine("PASS: DataGridView.RowHeadersVisible default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: DataGridView.RowHeadersVisible default")
            failed = failed + 1
        End If

        ' TabControl enhancements
        Dim tc As New TabControl()
        If tc.Alignment = "Top" Then
            Console.WriteLine("PASS: TabControl.Alignment default is Top")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TabControl.Alignment default")
            failed = failed + 1
        End If

        If tc.Appearance = "Normal" Then
            Console.WriteLine("PASS: TabControl.Appearance default is Normal")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TabControl.Appearance default")
            failed = failed + 1
        End If

        ' TreeView enhancements
        Dim tv As New TreeView()
        If tv.ShowPlusMinus = True Then
            Console.WriteLine("PASS: TreeView.ShowPlusMinus default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TreeView.ShowPlusMinus default")
            failed = failed + 1
        End If

        If tv.LabelEdit = False Then
            Console.WriteLine("PASS: TreeView.LabelEdit default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TreeView.LabelEdit default")
            failed = failed + 1
        End If

        If tv.Scrollable = True Then
            Console.WriteLine("PASS: TreeView.Scrollable default is True")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: TreeView.Scrollable default")
            failed = failed + 1
        End If

        ' ListView enhancements
        Dim lv As New ListView()
        If lv.ShowGroups = False Then
            Console.WriteLine("PASS: ListView.ShowGroups default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ListView.ShowGroups default")
            failed = failed + 1
        End If

        If lv.Sorting = "None" Then
            Console.WriteLine("PASS: ListView.Sorting default is None")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ListView.Sorting default")
            failed = failed + 1
        End If

        If lv.AllowColumnReorder = False Then
            Console.WriteLine("PASS: ListView.AllowColumnReorder default is False")
            passed = passed + 1
        Else
            Console.WriteLine("FAIL: ListView.AllowColumnReorder default")
            failed = failed + 1
        End If

        Console.WriteLine("")
        Console.WriteLine("=== Results ===")
        Console.WriteLine("Passed: " & passed.ToString())
        Console.WriteLine("Failed: " & failed.ToString())
        Console.WriteLine("Total:  " & (passed + failed).ToString())
    End Sub
End Module
