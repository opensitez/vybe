Module TestCollectionsControls
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, label As String)
        If condition Then
            Console.WriteLine("  PASS: " & label)
            passed = passed + 1
        Else
            Console.WriteLine("  FAIL: " & label)
            failed = failed + 1
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== Collections & Controls Test Suite ===")
        Console.WriteLine()

        ' ── 1. Variant / Object types ────────────────────────────────
        Console.WriteLine("--- 1. Variant / Object types ---")
        Dim v As Variant
        Assert(v Is Nothing, "Variant starts as Nothing")
        v = 42
        Assert(v = 42, "Variant can hold Integer")
        v = "hello"
        Assert(v = "hello", "Variant can hold String")
        v = True
        Assert(v = True, "Variant can hold Boolean")

        Dim obj As Object
        Assert(obj Is Nothing, "Object starts as Nothing")
        obj = 3.14
        Assert(obj = 3.14, "Object can hold Double")
        obj = "world"
        Assert(obj = "world", "Object can hold String")
        Console.WriteLine()

        ' ── 2. New Collection (VB6 Collection) ───────────────────────
        Console.WriteLine("--- 2. New Collection ---")
        Dim col As Object = New Collection()
        Assert(col IsNot Nothing, "New Collection creates object")
        col.Add("Apple")
        col.Add("Banana")
        col.Add("Cherry")
        Assert(col.Count = 3, "Collection.Count = 3")
        Console.WriteLine()

        ' ── 3. ArrayList ─────────────────────────────────────────────
        Console.WriteLine("--- 3. ArrayList ---")
        Dim al As Object = New ArrayList()
        al.Add(10)
        al.Add(20)
        al.Add(30)
        Assert(al.Count = 3, "ArrayList.Count = 3")
        Assert(al.Contains(20), "ArrayList.Contains(20)")
        Assert(al.IndexOf(30) = 2, "ArrayList.IndexOf(30) = 2")
        al.Remove(20)
        Assert(al.Count = 2, "After Remove: Count = 2")
        al.Insert(0, 5)
        Assert(al.Count = 3, "After Insert: Count = 3")
        Assert(al.Item(0) = 5, "Item(0) = 5 after insert at 0")
        Console.WriteLine()

        ' ── 4. List(Of T) ────────────────────────────────────────────
        Console.WriteLine("--- 4. List(Of T) ---")
        Dim lst As New List(Of String)
        lst.Add("x")
        lst.Add("y")
        lst.Add("z")
        Assert(lst.Count = 3, "List(Of String).Count = 3")
        Assert(lst.Contains("y"), "List Contains 'y'")
        lst.Clear()
        Assert(lst.Count = 0, "After Clear: Count = 0")
        Console.WriteLine()

        ' ── 5. Queue ─────────────────────────────────────────────────
        Console.WriteLine("--- 5. Queue ---")
        Dim q As New Queue()
        q.Enqueue("A")
        q.Enqueue("B")
        q.Enqueue("C")
        Assert(q.Count = 3, "Queue.Count = 3")
        Dim deq As String = q.Dequeue()
        Assert(deq = "A", "Dequeue returns 'A' (FIFO)")
        Assert(q.Count = 2, "After Dequeue: Count = 2")
        Assert(q.Peek() = "B", "Peek returns 'B'")
        Assert(q.Contains("C"), "Queue contains 'C'")
        Console.WriteLine()

        ' ── 6. Stack ─────────────────────────────────────────────────
        Console.WriteLine("--- 6. Stack ---")
        Dim s As New Stack()
        s.Push("X")
        s.Push("Y")
        s.Push("Z")
        Assert(s.Count = 3, "Stack.Count = 3")
        Dim popped As String = s.Pop()
        Assert(popped = "Z", "Pop returns 'Z' (LIFO)")
        Assert(s.Count = 2, "After Pop: Count = 2")
        Assert(s.Peek() = "Y", "Peek returns 'Y'")
        Console.WriteLine()

        ' ── 7. Dictionary ────────────────────────────────────────────
        Console.WriteLine("--- 7. Dictionary ---")
        Dim d As New Dictionary(Of String, Integer)
        d.Add("one", 1)
        d.Add("two", 2)
        d.Add("three", 3)
        Assert(d.Count = 3, "Dictionary.Count = 3")
        Assert(d.ContainsKey("two"), "ContainsKey('two')")
        Assert(d.ContainsValue(3), "ContainsValue(3)")
        Assert(d.Item("one") = 1, "Item('one') = 1")
        d.Remove("two")
        Assert(d.Count = 2, "After Remove: Count = 2")
        Console.WriteLine()

        ' ── 8. HashSet ───────────────────────────────────────────────
        Console.WriteLine("--- 8. HashSet ---")
        Dim hs As New HashSet(Of String)
        hs.Add("a")
        hs.Add("b")
        hs.Add("a")  ' Duplicate
        Assert(hs.Count = 2, "HashSet.Count = 2 (no dupes)")
        Assert(hs.Contains("b"), "HashSet contains 'b'")
        hs.Remove("a")
        Assert(hs.Count = 1, "After Remove: Count = 1")
        Console.WriteLine()

        ' ── 9. For Each on collections ───────────────────────────────
        Console.WriteLine("--- 9. For Each iteration ---")
        Dim list2 As New List(Of Integer)
        list2.Add(10)
        list2.Add(20)
        list2.Add(30)
        Dim total As Integer = 0
        For Each item As Integer In list2
            total = total + item
        Next
        Assert(total = 60, "For Each on List: sum = 60")

        ' For Each on Dictionary
        Dim d2 As New Dictionary(Of String, Integer)
        d2.Add("x", 1)
        d2.Add("y", 2)
        Dim keyCount As Integer = 0
        For Each kvp In d2
            keyCount = keyCount + 1
        Next
        Assert(keyCount = 2, "For Each on Dictionary: 2 items")

        ' For Each on Array
        Dim arr() As String = {"red", "green", "blue"}
        Dim colorCount As Integer = 0
        For Each c As String In arr
            colorCount = colorCount + 1
        Next
        Assert(colorCount = 3, "For Each on Array: 3 items")

        ' For Each on String
        Dim charCount As Integer = 0
        For Each ch In "Hello"
            charCount = charCount + 1
        Next
        Assert(charCount = 5, "For Each on String: 5 chars")

        ' For Each on Queue
        Dim q2 As New Queue()
        q2.Enqueue(1)
        q2.Enqueue(2)
        q2.Enqueue(3)
        Dim qSum As Integer = 0
        For Each qi In q2
            qSum = qSum + qi
        Next
        Assert(qSum = 6, "For Each on Queue: sum = 6")
        Console.WriteLine()

        ' ── 10. Simulated Controls collection ────────────────────────
        Console.WriteLine("--- 10. Controls collection ---")
        ' Simulate a form object with controls as fields
        Dim btn1 As New Button()
        btn1.Name = "Button1"
        btn1.Text = "Click Me"
        btn1.Enabled = True
        btn1.Visible = True

        Dim btn2 As New Button()
        btn2.Name = "Button2"
        btn2.Text = "Cancel"
        btn2.Enabled = True
        btn2.Visible = True

        Dim txt1 As New TextBox()
        txt1.Name = "TextBox1"
        txt1.Text = "Hello"
        txt1.Enabled = True
        txt1.Visible = True

        Dim lbl1 As New Label()
        lbl1.Name = "Label1"
        lbl1.Text = "Status"
        lbl1.Enabled = True
        lbl1.Visible = True

        ' Verify initial state
        Assert(btn1.Enabled = True, "btn1 starts Enabled")
        Assert(btn2.Visible = True, "btn2 starts Visible")
        Assert(txt1.Enabled = True, "txt1 starts Enabled")
        Assert(lbl1.Visible = True, "lbl1 starts Visible")

        ' Put controls in an array and iterate to disable/hide them
        Dim controls() As Object = {btn1, btn2, txt1, lbl1}

        For Each ctrl As Object In controls
            ctrl.Enabled = False
            ctrl.Visible = False
        Next

        ' Verify changes were applied
        Assert(btn1.Enabled = False, "btn1 disabled after loop")
        Assert(btn1.Visible = False, "btn1 hidden after loop")
        Assert(btn2.Enabled = False, "btn2 disabled after loop")
        Assert(btn2.Visible = False, "btn2 hidden after loop")
        Assert(txt1.Enabled = False, "txt1 disabled after loop")
        Assert(txt1.Visible = False, "txt1 hidden after loop")
        Assert(lbl1.Enabled = False, "lbl1 disabled after loop")
        Assert(lbl1.Visible = False, "lbl1 hidden after loop")
        Console.WriteLine()

        ' ── 11. Control property completeness ────────────────────────
        Console.WriteLine("--- 11. Control properties ---")
        Dim b As New Button()
        b.Name = "TestBtn"
        b.Text = "Test"
        b.Width = 200
        b.Height = 50
        b.Left = 10
        b.Top = 20
        b.TabIndex = 3
        b.BackColor = "Red"
        b.ForeColor = "White"
        b.Tag = "mytag"

        Assert(b.Name = "TestBtn", "Button.Name")
        Assert(b.Text = "Test", "Button.Text")
        Assert(b.Width = 200, "Button.Width")
        Assert(b.Height = 50, "Button.Height")
        Assert(b.Left = 10, "Button.Left")
        Assert(b.Top = 20, "Button.Top")
        Assert(b.TabIndex = 3, "Button.TabIndex")
        Assert(b.BackColor = "Red", "Button.BackColor")
        Assert(b.ForeColor = "White", "Button.ForeColor")
        Assert(b.Tag = "mytag", "Button.Tag")

        ' CheckBox
        Dim cb As New CheckBox()
        cb.Checked = True
        Assert(cb.Checked = True, "CheckBox.Checked = True")
        cb.Checked = False
        Assert(cb.Checked = False, "CheckBox.Checked = False")

        ' ComboBox items
        Dim cmb As New ComboBox()
        cmb.Items.Add("Option 1")
        cmb.Items.Add("Option 2")
        cmb.Items.Add("Option 3")
        Assert(cmb.Items.Count = 3, "ComboBox.Items.Count = 3")

        ' ListBox items
        Dim lb As New ListBox()
        lb.Items.Add("Item A")
        lb.Items.Add("Item B")
        Assert(lb.Items.Count = 2, "ListBox.Items.Count = 2")

        ' NumericUpDown
        Dim nud As New NumericUpDown()
        nud.Minimum = 0
        nud.Maximum = 100
        nud.Value = 50
        Assert(nud.Value = 50, "NumericUpDown.Value = 50")
        Assert(nud.Minimum = 0, "NumericUpDown.Minimum = 0")
        Assert(nud.Maximum = 100, "NumericUpDown.Maximum = 100")

        ' ProgressBar
        Dim pb As New ProgressBar()
        pb.Minimum = 0
        pb.Maximum = 100
        pb.Value = 75
        Assert(pb.Value = 75, "ProgressBar.Value = 75")
        Console.WriteLine()

        ' ── 12. Collection initializer syntax ────────────────────────
        Console.WriteLine("--- 12. Collection initializers ---")
        Dim fruits As New List(Of String) From {"Apple", "Banana", "Cherry"}
        Assert(fruits.Count = 3, "List initializer: Count = 3")
        Assert(fruits.Contains("Banana"), "List initializer contains 'Banana'")
        Console.WriteLine()

        ' ── 13. ArrayList methods: Sort, Reverse, Find ───────────────
        Console.WriteLine("--- 13. ArrayList advanced ---")
        Dim nums As New ArrayList()
        nums.Add(30)
        nums.Add(10)
        nums.Add(20)
        nums.Sort()
        Assert(nums.Item(0) = 10, "After Sort: first = 10")
        Assert(nums.Item(2) = 30, "After Sort: last = 30")
        nums.Reverse()
        Assert(nums.Item(0) = 30, "After Reverse: first = 30")
        Dim idx As Integer = nums.LastIndexOf(10)
        Assert(idx = 2, "LastIndexOf(10) = 2")
        Console.WriteLine()

        ' ── 14. Controls.Add tracking ────────────────────────────────
        Console.WriteLine("--- 14. Controls.Add tracking ---")
        ' Create a container (panel) and add controls to it
        Dim pnl As New Panel()
        pnl.Name = "Panel1"

        Dim addBtn As New Button()
        addBtn.Name = "DynBtn1"
        addBtn.Text = "Dynamic"

        Dim addTxt As New TextBox()
        addTxt.Name = "DynTxt1"
        addTxt.Text = "Dynamic Text"

        pnl.Controls.Add(addBtn)
        pnl.Controls.Add(addTxt)

        ' Now iterate the controls
        Dim dynCount As Integer = 0
        For Each c In pnl.Controls
            dynCount = dynCount + 1
        Next
        Assert(dynCount >= 2, "Panel.Controls has added controls")
        Console.WriteLine()

        ' ── Summary ──────────────────────────────────────────────────
        Console.WriteLine()
        Console.WriteLine("=== Results: " & passed & " passed, " & failed & " failed ===")
        If failed = 0 Then
            Console.WriteLine("=== ALL TESTS PASSED ===")
        End If
    End Sub
End Module
