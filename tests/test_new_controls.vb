' Test new controls: DateTimePicker, LinkLabel, ToolStrip, TrackBar,
' MaskedTextBox, SplitContainer, FlowLayoutPanel, TableLayoutPanel
' Tests cover: creation, properties, and events

Module TestNewControls

    ' ===== DateTimePicker =====
    Sub TestDateTimePicker()
        Console.WriteLine("=== DateTimePicker Tests ===")
        
        ' Creation
        Dim dtp As New DateTimePicker()
        Console.WriteLine("DateTimePicker created: " & (dtp IsNot Nothing).ToString())
        
        ' Default properties
        Console.WriteLine("Default Format: " & dtp.Format)
        Console.WriteLine("Default Checked: " & dtp.Checked.ToString())
        Console.WriteLine("Default ShowCheckBox: " & dtp.ShowCheckBox.ToString())
        Console.WriteLine("Default Enabled: " & dtp.Enabled.ToString())
        Console.WriteLine("Default Visible: " & dtp.Visible.ToString())
        
        ' Set properties
        dtp.Format = "Short"
        Console.WriteLine("Format after set: " & dtp.Format)
        
        dtp.CustomFormat = "yyyy-MM-dd"
        Console.WriteLine("CustomFormat: " & dtp.CustomFormat)
        
        dtp.Value = "2025-01-15"
        Console.WriteLine("Value: " & dtp.Value)
        
        dtp.ShowCheckBox = True
        Console.WriteLine("ShowCheckBox: " & dtp.ShowCheckBox.ToString())
        
        dtp.Checked = False
        Console.WriteLine("Checked: " & dtp.Checked.ToString())
        
        dtp.Enabled = False
        Console.WriteLine("Enabled after disable: " & dtp.Enabled.ToString())
        
        dtp.Visible = False
        Console.WriteLine("Visible after hide: " & dtp.Visible.ToString())
        
        Console.WriteLine("DateTimePicker PASS")
    End Sub

    ' ===== LinkLabel =====
    Sub TestLinkLabel()
        Console.WriteLine("=== LinkLabel Tests ===")
        
        ' Creation
        Dim lnk As New LinkLabel()
        Console.WriteLine("LinkLabel created: " & (lnk IsNot Nothing).ToString())
        
        ' Default properties
        Console.WriteLine("Default LinkColor: " & lnk.LinkColor)
        Console.WriteLine("Default VisitedLinkColor: " & lnk.VisitedLinkColor)
        Console.WriteLine("Default LinkVisited: " & lnk.LinkVisited.ToString())
        Console.WriteLine("Default Enabled: " & lnk.Enabled.ToString())
        Console.WriteLine("Default Visible: " & lnk.Visible.ToString())
        
        ' Set properties
        lnk.Text = "Click Here"
        Console.WriteLine("Text: " & lnk.Text)
        
        lnk.LinkColor = "#FF0000"
        Console.WriteLine("LinkColor: " & lnk.LinkColor)
        
        lnk.VisitedLinkColor = "#00FF00"
        Console.WriteLine("VisitedLinkColor: " & lnk.VisitedLinkColor)
        
        lnk.LinkVisited = True
        Console.WriteLine("LinkVisited: " & lnk.LinkVisited.ToString())
        
        Console.WriteLine("LinkLabel PASS")
    End Sub

    ' ===== ToolStrip =====
    Sub TestToolStrip()
        Console.WriteLine("=== ToolStrip Tests ===")
        
        ' Creation
        Dim ts As New ToolStrip()
        Console.WriteLine("ToolStrip created: " & (ts IsNot Nothing).ToString())
        
        ' Properties
        Console.WriteLine("Default Enabled: " & ts.Enabled.ToString())
        Console.WriteLine("Default Visible: " & ts.Visible.ToString())
        
        ' Items collection
        Console.WriteLine("Items exists: " & (ts.Items IsNot Nothing).ToString())
        
        ts.Enabled = False
        Console.WriteLine("Enabled after disable: " & ts.Enabled.ToString())
        
        Console.WriteLine("ToolStrip PASS")
    End Sub

    ' ===== TrackBar =====
    Sub TestTrackBar()
        Console.WriteLine("=== TrackBar Tests ===")
        
        ' Creation
        Dim trk As New TrackBar()
        Console.WriteLine("TrackBar created: " & (trk IsNot Nothing).ToString())
        
        ' Default properties
        Console.WriteLine("Default Value: " & trk.Value.ToString())
        Console.WriteLine("Default Minimum: " & trk.Minimum.ToString())
        Console.WriteLine("Default Maximum: " & trk.Maximum.ToString())
        Console.WriteLine("Default TickFrequency: " & trk.TickFrequency.ToString())
        Console.WriteLine("Default SmallChange: " & trk.SmallChange.ToString())
        Console.WriteLine("Default LargeChange: " & trk.LargeChange.ToString())
        Console.WriteLine("Default Orientation: " & trk.Orientation)
        Console.WriteLine("Default Enabled: " & trk.Enabled.ToString())
        
        ' Set properties
        trk.Value = 5
        Console.WriteLine("Value after set: " & trk.Value.ToString())
        
        trk.Minimum = 1
        Console.WriteLine("Minimum: " & trk.Minimum.ToString())
        
        trk.Maximum = 20
        Console.WriteLine("Maximum: " & trk.Maximum.ToString())
        
        trk.TickFrequency = 2
        Console.WriteLine("TickFrequency: " & trk.TickFrequency.ToString())
        
        trk.SmallChange = 2
        Console.WriteLine("SmallChange: " & trk.SmallChange.ToString())
        
        trk.LargeChange = 10
        Console.WriteLine("LargeChange: " & trk.LargeChange.ToString())
        
        trk.Orientation = "Vertical"
        Console.WriteLine("Orientation: " & trk.Orientation)
        
        Console.WriteLine("TrackBar PASS")
    End Sub

    ' ===== MaskedTextBox =====
    Sub TestMaskedTextBox()
        Console.WriteLine("=== MaskedTextBox Tests ===")
        
        ' Creation
        Dim mtxt As New MaskedTextBox()
        Console.WriteLine("MaskedTextBox created: " & (mtxt IsNot Nothing).ToString())
        
        ' Default properties
        Console.WriteLine("Default Mask: [" & mtxt.Mask & "]")
        Console.WriteLine("Default PromptChar: " & mtxt.PromptChar)
        Console.WriteLine("Default Enabled: " & mtxt.Enabled.ToString())
        Console.WriteLine("Default Visible: " & mtxt.Visible.ToString())
        
        ' Set properties
        mtxt.Text = "555-12-3456"
        Console.WriteLine("Text: " & mtxt.Text)
        
        mtxt.Mask = "000-00-0000"
        Console.WriteLine("Mask: " & mtxt.Mask)
        
        mtxt.PromptChar = "#"
        Console.WriteLine("PromptChar: " & mtxt.PromptChar)
        
        Console.WriteLine("MaskedTextBox PASS")
    End Sub

    ' ===== SplitContainer =====
    Sub TestSplitContainer()
        Console.WriteLine("=== SplitContainer Tests ===")
        
        ' Creation
        Dim sc As New SplitContainer()
        Console.WriteLine("SplitContainer created: " & (sc IsNot Nothing).ToString())
        
        ' Default properties
        Console.WriteLine("Default Orientation: " & sc.Orientation)
        Console.WriteLine("Default SplitterDistance: " & sc.SplitterDistance.ToString())
        Console.WriteLine("Default Enabled: " & sc.Enabled.ToString())
        Console.WriteLine("Default Visible: " & sc.Visible.ToString())
        
        ' Set properties
        sc.Orientation = "Horizontal"
        Console.WriteLine("Orientation: " & sc.Orientation)
        
        sc.SplitterDistance = 200
        Console.WriteLine("SplitterDistance: " & sc.SplitterDistance.ToString())
        
        sc.Enabled = False
        Console.WriteLine("Enabled: " & sc.Enabled.ToString())
        
        Console.WriteLine("SplitContainer PASS")
    End Sub

    ' ===== FlowLayoutPanel =====
    Sub TestFlowLayoutPanel()
        Console.WriteLine("=== FlowLayoutPanel Tests ===")
        
        ' Creation
        Dim flp As New FlowLayoutPanel()
        Console.WriteLine("FlowLayoutPanel created: " & (flp IsNot Nothing).ToString())
        
        ' Default properties
        Console.WriteLine("Default FlowDirection: " & flp.FlowDirection)
        Console.WriteLine("Default WrapContents: " & flp.WrapContents.ToString())
        Console.WriteLine("Default Enabled: " & flp.Enabled.ToString())
        Console.WriteLine("Default Visible: " & flp.Visible.ToString())
        
        ' Set properties
        flp.FlowDirection = "TopDown"
        Console.WriteLine("FlowDirection: " & flp.FlowDirection)
        
        flp.WrapContents = False
        Console.WriteLine("WrapContents: " & flp.WrapContents.ToString())
        
        Console.WriteLine("FlowLayoutPanel PASS")
    End Sub

    ' ===== TableLayoutPanel =====
    Sub TestTableLayoutPanel()
        Console.WriteLine("=== TableLayoutPanel Tests ===")
        
        ' Creation
        Dim tlp As New TableLayoutPanel()
        Console.WriteLine("TableLayoutPanel created: " & (tlp IsNot Nothing).ToString())
        
        ' Default properties
        Console.WriteLine("Default ColumnCount: " & tlp.ColumnCount.ToString())
        Console.WriteLine("Default RowCount: " & tlp.RowCount.ToString())
        Console.WriteLine("Default Enabled: " & tlp.Enabled.ToString())
        Console.WriteLine("Default Visible: " & tlp.Visible.ToString())
        
        ' Set properties
        tlp.ColumnCount = 3
        Console.WriteLine("ColumnCount: " & tlp.ColumnCount.ToString())
        
        tlp.RowCount = 4
        Console.WriteLine("RowCount: " & tlp.RowCount.ToString())
        
        Console.WriteLine("TableLayoutPanel PASS")
    End Sub

    ' ===== StatusStrip (verify existing) =====
    Sub TestStatusStrip()
        Console.WriteLine("=== StatusStrip Tests ===")
        
        Dim ss As New StatusStrip()
        Console.WriteLine("StatusStrip created: " & (ss IsNot Nothing).ToString())
        Console.WriteLine("Items exists: " & (ss.Items IsNot Nothing).ToString())
        Console.WriteLine("Default Enabled: " & ss.Enabled.ToString())
        Console.WriteLine("Default Visible: " & ss.Visible.ToString())
        
        Console.WriteLine("StatusStrip PASS")
    End Sub

    ' ===== ToolStripStatusLabel =====
    Sub TestToolStripStatusLabel()
        Console.WriteLine("=== ToolStripStatusLabel Tests ===")
        
        Dim tssl As New ToolStripStatusLabel()
        Console.WriteLine("ToolStripStatusLabel created: " & (tssl IsNot Nothing).ToString())
        
        tssl.Text = "Ready"
        Console.WriteLine("Text: " & tssl.Text)
        
        tssl.Spring = True
        Console.WriteLine("Spring: " & tssl.Spring.ToString())
        
        tssl.AutoSize = False
        Console.WriteLine("AutoSize: " & tssl.AutoSize.ToString())
        
        Console.WriteLine("ToolStripStatusLabel PASS")
    End Sub

    Sub Main()
        TestDateTimePicker()
        TestLinkLabel()
        TestToolStrip()
        TestTrackBar()
        TestMaskedTextBox()
        TestSplitContainer()
        TestFlowLayoutPanel()
        TestTableLayoutPanel()
        TestStatusStrip()
        TestToolStripStatusLabel()
        Console.WriteLine("")
        Console.WriteLine("ALL NEW CONTROL TESTS PASSED")
    End Sub

End Module
