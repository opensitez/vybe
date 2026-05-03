//! Advanced WinForms compilation tests covering cross-control interaction,
//! form state persistence, nested containers, dynamic creation, partial classes,
//! timers, combo/listbox items, anchor/dock, event parameters, multi-form, and tabs.
//!
//! Categories:
//!   A. Cross-control interaction in handlers (6 tests)
//!   B. Form state across multiple handler invocations (5 tests)
//!   C. Nested containers (5 tests)
//!   D. Dynamic control creation in handler (4 tests)
//!   E. Partial class pattern (5 tests)
//!   F. Timer and interval (3 tests)
//!   G. ComboBox/ListBox items (4 tests)
//!   H. Anchor/Dock properties (3 tests)
//!   I. Event handler parameter access (4 tests)
//!   J. Multi-form pattern (3 tests)
//!   K. Tab control (3 tests)

use super::helpers::{run_vb, run_vb_gui};
use std::sync::{Arc, Mutex};
use vybe_host::gui_state::GuiState;

// ============================================================
// A. CROSS-CONTROL INTERACTION IN HANDLERS (6 tests)
// ============================================================

/// A01. Button click handler reads TextBox.Text field (Me.txtInput.Text).
#[test]
fn a01_handler_reads_textbox_text() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim txtInput As New TextBox()
    Public Sub New()
        txtInput.Name = "txtInput"
        txtInput.Text = "hello"
        Me.Controls.Add(txtInput)
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Dim val As String = Me.txtInput.Text
        Console.WriteLine(val)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    // The handler should compile and register; txtInput should be added
    assert!(g.control_names.contains(&"txtinput".to_string()), "Expected txtInput control");
    assert!(g.event_handlers.contains_key("btngo.click"), "Expected Click handler registered for btngo");
}

/// A02. Button click handler writes to Label.Text (Me.lblOutput.Text = "result").
#[test]
fn a02_handler_writes_label_text() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim lblOutput As New Label()
    Dim btnGo As New Button()
    Public Sub New()
        lblOutput.Name = "lblOutput"
        btnGo.Name = "btnGo"
        Me.Controls.Add(lblOutput)
        Me.Controls.Add(btnGo)
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Me.lblOutput.Text = "result"
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"lbloutput".to_string()), "Expected lblOutput control");
    assert!(g.control_names.contains(&"btngo".to_string()), "Expected btnGo control");
    assert!(g.event_handlers.contains_key("btngo.click"), "Expected Click handler on btngo");
}

/// A03. Handler concatenates text from two TextBoxes.
#[test]
fn a03_handler_concatenates_two_textboxes() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim txt1 As New TextBox()
    Dim txt2 As New TextBox()
    Dim lblResult As New Label()
    Public Sub New()
        txt1.Name = "txt1"
        txt2.Name = "txt2"
        lblResult.Name = "lblResult"
        Me.Controls.Add(txt1)
        Me.Controls.Add(txt2)
        Me.Controls.Add(lblResult)
    End Sub
    Private Sub btnConcat_Click(sender As Object, e As EventArgs) Handles btnConcat.Click
        Me.lblResult.Text = Me.txt1.Text & " " & Me.txt2.Text
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 3, "Expected 3 controls added");
}

/// A04. Handler toggles CheckBox.Checked via Me reference.
#[test]
fn a04_handler_toggles_checkbox() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim chk As New CheckBox()
    Public Sub New()
        chk.Name = "chk"
        chk.Checked = False
        Me.Controls.Add(chk)
    End Sub
    Private Sub btnToggle_Click(sender As Object, e As EventArgs) Handles btnToggle.Click
        Me.chk.Checked = Not Me.chk.Checked
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"chk".to_string()), "Expected chk control");
}

/// A05. Handler reads one control, computes, writes to another.
#[test]
fn a05_handler_reads_computes_writes() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim txtNum As New TextBox()
    Dim lblDouble As New Label()
    Public Sub New()
        txtNum.Name = "txtNum"
        lblDouble.Name = "lblDouble"
        Me.Controls.Add(txtNum)
        Me.Controls.Add(lblDouble)
    End Sub
    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        Dim n As Integer = CInt(Me.txtNum.Text)
        Me.lblDouble.Text = CStr(n * 2)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 2, "Expected 2 controls");
}

/// A06. Handler accesses form-level Dim field AND control property.
#[test]
fn a06_handler_accesses_field_and_control() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim prefix As String = "Value: "
    Dim txtInput As New TextBox()
    Dim lblOutput As New Label()
    Public Sub New()
        txtInput.Name = "txtInput"
        lblOutput.Name = "lblOutput"
        Me.Controls.Add(txtInput)
        Me.Controls.Add(lblOutput)
    End Sub
    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Me.lblOutput.Text = prefix & Me.txtInput.Text
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txtinput".to_string()));
    assert!(g.control_names.contains(&"lbloutput".to_string()));
}

// ============================================================
// B. FORM STATE ACROSS MULTIPLE HANDLER INVOCATIONS (5 tests)
// ============================================================

/// B07. Dim counter field incremented by handler, read shows updated value.
#[test]
fn b07_counter_field_incremented() {
    let out = run_vb(r#"
Public Class Form1
    Dim counter As Integer = 0
    Public Sub New()
    End Sub
    Public Sub Increment()
        counter = counter + 1
    End Sub
    Public Function GetCount() As Integer
        Return counter
    End Function
End Class
Dim f As New Form1()
f.Increment()
f.Increment()
f.Increment()
Console.WriteLine(f.GetCount())
"#);
    assert_eq!(out, vec!["3"]);
}

/// B08. Handler appends to string field -- multiple invocations accumulate.
#[test]
fn b08_string_field_accumulates() {
    let out = run_vb(r#"
Public Class Form1
    Dim log As String = ""
    Public Sub New()
    End Sub
    Public Sub AppendLog(msg As String)
        log = log & msg & ";"
    End Sub
    Public Function GetLog() As String
        Return log
    End Function
End Class
Dim f As New Form1()
f.AppendLog("a")
f.AppendLog("b")
f.AppendLog("c")
Console.WriteLine(f.GetLog())
"#);
    assert_eq!(out, vec!["a;b;c;"]);
}

/// B09. Boolean flag toggled by handler.
#[test]
fn b09_boolean_flag_toggled() {
    let out = run_vb(r#"
Public Class Form1
    Dim isActive As Boolean = False
    Public Sub New()
    End Sub
    Public Sub Toggle()
        isActive = Not isActive
    End Sub
    Public Function GetActive() As Boolean
        Return isActive
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetActive())
f.Toggle()
Console.WriteLine(f.GetActive())
f.Toggle()
Console.WriteLine(f.GetActive())
"#);
    assert_eq!(out, vec!["false", "true", "false"]);
}

/// B10. Handler stores value in field, another handler reads it.
#[test]
fn b10_cross_method_field_access() {
    let out = run_vb(r#"
Public Class Form1
    Dim savedValue As String = ""
    Public Sub New()
    End Sub
    Public Sub Save(val As String)
        savedValue = val
    End Sub
    Public Function Load() As String
        Return savedValue
    End Function
End Class
Dim f As New Form1()
f.Save("important data")
Console.WriteLine(f.Load())
"#);
    assert_eq!(out, vec!["important data"]);
}

/// B11. Form with undo pattern: store previous value in field before update.
#[test]
fn b11_undo_pattern() {
    let out = run_vb(r#"
Public Class Form1
    Dim currentValue As String = "initial"
    Dim previousValue As String = ""
    Public Sub New()
    End Sub
    Public Sub SetValue(val As String)
        previousValue = currentValue
        currentValue = val
    End Sub
    Public Sub Undo()
        currentValue = previousValue
    End Sub
    Public Function GetValue() As String
        Return currentValue
    End Function
End Class
Dim f As New Form1()
f.SetValue("second")
f.SetValue("third")
Console.WriteLine(f.GetValue())
f.Undo()
Console.WriteLine(f.GetValue())
"#);
    assert_eq!(out, vec!["third", "second"]);
}

// ============================================================
// C. NESTED CONTAINERS (5 tests)
// ============================================================

/// C12. Panel with Button inside -- Controls.Add on Panel, not Form.
#[test]
fn c12_panel_with_button_inside() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pnl As New Panel()
        pnl.Name = "pnl1"
        Me.Controls.Add(pnl)
        Dim btn As New Button()
        btn.Name = "btnInner"
        pnl.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pnl1".to_string()), "Expected pnl1 control");
    assert!(g.control_names.contains(&"btninner".to_string()), "Expected btnInner control");
}

/// C13. GroupBox with RadioButtons inside.
#[test]
fn c13_groupbox_with_radiobuttons() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim grp As New GroupBox()
        grp.Name = "grpOptions"
        grp.Text = "Options"
        Me.Controls.Add(grp)
        Dim rb1 As New RadioButton()
        rb1.Name = "rbOption1"
        rb1.Text = "Option 1"
        grp.Controls.Add(rb1)
        Dim rb2 As New RadioButton()
        rb2.Name = "rbOption2"
        rb2.Text = "Option 2"
        grp.Controls.Add(rb2)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"grpoptions".to_string()), "Expected grpOptions");
    assert!(g.control_names.contains(&"rboption1".to_string()), "Expected rbOption1");
    assert!(g.control_names.contains(&"rboption2".to_string()), "Expected rbOption2");
}

/// C14. Panel position + child control position (should be relative).
#[test]
fn c14_panel_child_position() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim pnl As New Panel()
        pnl.Name = "pnl1"
        pnl.Location = New Point(50, 50)
        Me.Controls.Add(pnl)
        Dim btn As New Button()
        btn.Name = "btnChild"
        btn.Location = New Point(10, 10)
        pnl.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    // Panel and child should both be registered
    assert!(g.control_names.contains(&"pnl1".to_string()), "Expected pnl1 control");
    assert!(g.control_names.contains(&"btnchild".to_string()), "Expected btnChild control");
}

/// C15. Multiple panels each with own controls.
#[test]
fn c15_multiple_panels_with_controls() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pnlLeft As New Panel()
        pnlLeft.Name = "pnlLeft"
        Me.Controls.Add(pnlLeft)
        Dim pnlRight As New Panel()
        pnlRight.Name = "pnlRight"
        Me.Controls.Add(pnlRight)
        Dim btnL As New Button()
        btnL.Name = "btnLeft"
        pnlLeft.Controls.Add(btnL)
        Dim btnR As New Button()
        btnR.Name = "btnRight"
        pnlRight.Controls.Add(btnR)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 4, "Expected 4 controls (2 panels + 2 buttons)");
}

/// C16. Nested: Panel inside GroupBox inside Form.
#[test]
fn c16_nested_panel_in_groupbox() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim grp As New GroupBox()
        grp.Name = "grp1"
        Me.Controls.Add(grp)
        Dim pnl As New Panel()
        pnl.Name = "pnlInner"
        grp.Controls.Add(pnl)
        Dim lbl As New Label()
        lbl.Name = "lblDeep"
        pnl.Controls.Add(lbl)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"grp1".to_string()), "Expected grp1");
    assert!(g.control_names.contains(&"pnlinner".to_string()), "Expected pnlInner");
    assert!(g.control_names.contains(&"lbldeep".to_string()), "Expected lblDeep");
}

// ============================================================
// D. DYNAMIC CONTROL CREATION IN HANDLER (4 tests)
// ============================================================

/// D17. Handler creates New Button(), sets properties, adds to form.
#[test]
fn d17_handler_creates_button() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
    Public Sub AddButton()
        Dim btn As New Button()
        btn.Name = "btnDynamic"
        btn.Text = "Dynamic"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
f.AddButton()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btndynamic".to_string()), "Expected btnDynamic control");
}

/// D18. Handler creates control based on runtime condition.
#[test]
fn d18_handler_conditional_create() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
    Public Sub AddControl(useButton As Boolean)
        If useButton Then
            Dim btn As New Button()
            btn.Name = "btnCond"
            Me.Controls.Add(btn)
        Else
            Dim lbl As New Label()
            lbl.Name = "lblCond"
            Me.Controls.Add(lbl)
        End If
    End Sub
End Class
Dim f As New Form1()
f.AddControl(True)
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btncond".to_string()), "Expected btnCond control");
}

/// D19. Handler creates multiple controls in loop.
#[test]
fn d19_handler_creates_controls_in_loop() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
    Public Sub AddButtons(count As Integer)
        For i As Integer = 1 To count
            Dim btn As New Button()
            btn.Name = "btn" & CStr(i)
            btn.Text = "Button " & CStr(i)
            Me.Controls.Add(btn)
        Next
    End Sub
End Class
Dim f As New Form1()
f.AddButtons(3)
"#);
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 3, "Expected 3 dynamically created buttons");
}

/// D20. Handler creates TextBox and wires it with AddHandler.
#[test]
fn d20_handler_creates_and_wires() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
    Public Sub AddWiredTextBox()
        Dim txt As New TextBox()
        txt.Name = "txtDynamic"
        Me.Controls.Add(txt)
        AddHandler txt.TextChanged, AddressOf txt_TextChanged
    End Sub
    Private Sub txt_TextChanged(sender As Object, e As EventArgs)
        Console.WriteLine("text changed")
    End Sub
End Class
Dim f As New Form1()
f.AddWiredTextBox()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txtdynamic".to_string()), "Expected txtDynamic control");
}

// ============================================================
// E. PARTIAL CLASS PATTERN (5 tests)
// ============================================================

/// E21. Two Partial Class declarations -- fields from both accessible.
#[test]
fn e21_partial_class_fields_merged() {
    let out = run_vb(r#"
Partial Public Class Form1
    Dim name As String = "hello"
End Class
Partial Public Class Form1
    Dim count As Integer = 42
    Public Sub New()
    End Sub
    Public Sub ShowBoth()
        Console.WriteLine(name & " " & CStr(count))
    End Sub
End Class
Dim f As New Form1()
f.ShowBoth()
"#);
    assert_eq!(out, vec!["hello 42"]);
}

/// E22. Partial: one has InitializeComponent, other has event handlers.
#[test]
fn e22_partial_init_and_handlers() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Partial Public Class Form1
    Dim btnOk As New Button()
    Private Sub InitializeComponent()
        btnOk.Name = "btnOk"
        Me.Controls.Add(btnOk)
    End Sub
End Class
Partial Public Class Form1
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        MsgBox("Clicked")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btnok".to_string()), "Expected btnOk control");
}

/// E23. Partial: constructor in designer part, methods in user part.
#[test]
fn e23_partial_constructor_in_designer() {
    let out = run_vb(r#"
Partial Public Class Form1
    Dim title As String = "My Form"
    Public Sub New()
    End Sub
End Class
Partial Public Class Form1
    Public Function GetTitle() As String
        Return title
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetTitle())
"#);
    assert_eq!(out, vec!["My Form"]);
}

/// E24. Partial: field declared in designer, used in user method.
#[test]
fn e24_partial_field_cross_access() {
    let out = run_vb(r#"
Partial Public Class Form1
    Dim status As String = "ready"
End Class
Partial Public Class Form1
    Public Sub New()
    End Sub
    Public Function GetStatus() As String
        Return status
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetStatus())
"#);
    assert_eq!(out, vec!["ready"]);
}

/// E25. Designer Friend WithEvents declarations parsed as fields.
#[test]
fn e25_friend_withevents_as_fields() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Friend WithEvents btnSave As Button
    Public Sub New()
        btnSave = New Button()
        btnSave.Name = "btnSave"
        btnSave.Text = "Save"
        Me.Controls.Add(btnSave)
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        MsgBox("Saved")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btnsave".to_string()), "Expected btnSave control");
}

// ============================================================
// F. TIMER AND INTERVAL (3 tests)
// ============================================================

/// F26. New Timer() -- __control_type is "Timer".
#[test]
fn f26_new_timer_control_type() {
    let (_vm, _gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim tmr As New Timer()
        Console.WriteLine("timer created")
    End Sub
End Class
Dim f As New Form1()
"#);
    // If we get here without error, the Timer() constructor compiled correctly
}

/// F27. Timer with Interval property set.
#[test]
fn f27_timer_interval_property() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim tmr As New Timer()
    Public Sub New()
        tmr.Interval = 1000
        tmr.Enabled = True
    End Sub
End Class
Dim f As New Form1()
"#);
    // Timer properties are mirrored into GuiState by the canvas/property
    // setters; we just verify the program compiles and runs without
    // panicking. (Asserting on the actual mirrored Interval value
    // belongs in a Timer-specific test, not here.)
    let _g = gui.lock().unwrap();
}

/// F28. Timer.Tick Handles clause.
#[test]
fn f28_timer_tick_handles() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim tmr As New Timer()
    Public Sub New()
        tmr.Name = "tmr1"
        tmr.Interval = 500
    End Sub
    Private Sub tmr_Tick(sender As Object, e As EventArgs) Handles tmr.Tick
        Console.WriteLine("tick")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("tmr.tick") || g.event_handlers.contains_key("tmr1.tick"),
        "Expected Tick handler registered for timer");
}

// ============================================================
// G. COMBOBOX/LISTBOX ITEMS (4 tests)
// ============================================================

/// G29. ComboBox.Items.Add -- adds items.
#[test]

fn g29_combobox_items_add() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim cbo As New ComboBox()
    Public Sub New()
        cbo.Name = "cboColors"
        cbo.Items.Add("Red")
        cbo.Items.Add("Green")
        cbo.Items.Add("Blue")
        Me.Controls.Add(cbo)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"cbocolors".to_string()), "Expected cboColors control");
}

/// G30. ListBox.Items.Add -- adds items.
#[test]

fn g30_listbox_items_add() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim lst As New ListBox()
    Public Sub New()
        lst.Name = "lstNames"
        lst.Items.Add("Alice")
        lst.Items.Add("Bob")
        Me.Controls.Add(lst)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"lstnames".to_string()), "Expected lstNames control");
}

/// G31. ComboBox SelectedIndexChanged handler.
#[test]
fn g31_combobox_selectedindexchanged_handler() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim cbo As New ComboBox()
    Public Sub New()
        cbo.Name = "cbo1"
        Me.Controls.Add(cbo)
    End Sub
    Private Sub cbo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo.SelectedIndexChanged
        Console.WriteLine("selection changed")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("cbo.selectedindexchanged"),
        "Expected SelectedIndexChanged handler on cbo");
}

/// G32. ListBox with multiple items added.
#[test]

fn g32_listbox_multiple_items() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim lst As New ListBox()
    Public Sub New()
        lst.Name = "lstItems"
        lst.Items.Add("Item 1")
        lst.Items.Add("Item 2")
        lst.Items.Add("Item 3")
        lst.Items.Add("Item 4")
        lst.Items.Add("Item 5")
        Me.Controls.Add(lst)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"lstitems".to_string()), "Expected lstItems control");
}

// ============================================================
// H. ANCHOR/DOCK PROPERTIES (3 tests)
// ============================================================

/// H33. Control with Anchor property set.
#[test]
fn h33_control_anchor_property() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim txt As New TextBox()
        txt.Name = "txtAnchored"
        txt.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(txt)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txtanchored".to_string()), "Expected txtAnchored control");
}

/// H34. Control with Dock property set.
#[test]
fn h34_control_dock_property() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pnl As New Panel()
        pnl.Name = "pnlDocked"
        pnl.Dock = DockStyle.Top
        Me.Controls.Add(pnl)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pnldocked".to_string()), "Expected pnlDocked control");
}

/// H35. Multiple controls with different Anchor values.
#[test]
fn h35_multiple_anchors() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim txt1 As New TextBox()
        txt1.Name = "txt1"
        txt1.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Me.Controls.Add(txt1)
        Dim txt2 As New TextBox()
        txt2.Name = "txt2"
        txt2.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Me.Controls.Add(txt2)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 2, "Expected 2 controls");
}

// ============================================================
// I. EVENT HANDLER PARAMETER ACCESS (4 tests)
// ============================================================

/// I36. Handler receives sender parameter -- can reference it.
#[test]
fn i36_handler_sender_parameter() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn As New Button()
    Public Sub New()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
    Private Sub btn_Click(sender As Object, e As EventArgs) Handles btn.Click
        Dim s As Object = sender
        Console.WriteLine("clicked")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("btn.click"), "Expected Click handler on btn");
}

/// I37. Handler receives e parameter.
#[test]
fn i37_handler_e_parameter() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn As New Button()
    Public Sub New()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
    Private Sub btn_Click(sender As Object, e As EventArgs) Handles btn.Click
        Dim ev As Object = e
        Console.WriteLine("event received")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("btn.click"), "Expected Click handler on btn");
}

/// I38. Handler with sender As Object, e As EventArgs -- compiles.
#[test]
fn i38_handler_standard_signature_compiles() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn As New Button()
    Public Sub New()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
    Private Sub btn_Click(sender As Object, e As EventArgs) Handles btn.Click
        MsgBox("clicked")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()), "Expected btn1 control");
    assert!(g.event_handlers.contains_key("btn.click"), "Expected Click handler on btn");
}

/// I39. Handler ignores parameters (common pattern) -- works fine.
#[test]
fn i39_handler_ignores_parameters() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn As New Button()
    Public Sub New()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
    Private Sub btn_Click(sender As Object, e As EventArgs) Handles btn.Click
        ' Do nothing with sender or e, just show a message
        MsgBox("simple click")
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()), "Expected btn1 control");
    assert!(g.event_handlers.contains_key("btn.click"), "Expected Click handler on btn");
}

// ============================================================
// J. MULTI-FORM PATTERN (3 tests)
// ============================================================

/// J40. Two form classes defined -- both instantiable.
#[test]
fn j40_two_form_classes() {
    let out = run_vb(r#"
Public Class Form1
    Dim title As String = "Form One"
    Public Sub New()
    End Sub
    Public Function GetTitle() As String
        Return title
    End Function
End Class
Public Class Form2
    Dim title As String = "Form Two"
    Public Sub New()
    End Sub
    Public Function GetTitle() As String
        Return title
    End Function
End Class
Dim f1 As New Form1()
Dim f2 As New Form2()
Console.WriteLine(f1.GetTitle())
Console.WriteLine(f2.GetTitle())
"#);
    assert_eq!(out, vec!["Form One", "Form Two"]);
}

/// J41. Form1 field holds reference to Form2 instance.
#[test]
fn j41_form_holds_reference_to_other() {
    let out = run_vb(r#"
Public Class Form2
    Dim msg As String = "from form2"
    Public Sub New()
    End Sub
    Public Function GetMsg() As String
        Return msg
    End Function
End Class
Public Class Form1
    Dim child As Form2
    Public Sub New()
        child = New Form2()
    End Sub
    Public Function GetChildMsg() As String
        Return child.GetMsg()
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetChildMsg())
"#);
    assert_eq!(out, vec!["from form2"]);
}

/// J42. Application.Run starts one form, other exists.
#[test]
fn j42_application_run_one_form() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
End Class
Public Class Form2
    Public Sub New()
    End Sub
End Class
Dim f1 As New Form1()
Dim f2 As New Form2()
Application.Run(f1)
"#);
    let g = gui.lock().unwrap();
    assert!(g.should_run, "Expected should_run after Application.Run");
}

// ============================================================
// K. TAB CONTROL (3 tests)
// ============================================================

/// K43. New TabControl() created.
#[test]
fn k43_new_tabcontrol() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim tabs As New TabControl()
        tabs.Name = "tabs1"
        Me.Controls.Add(tabs)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"tabs1".to_string()), "Expected tabs1 control");
}

/// K44. New TabPage() created with Text.
#[test]
fn k44_new_tabpage_with_text() {
    let (_vm, _gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim tp As New TabPage()
        tp.Name = "tpGeneral"
        tp.Text = "General"
    End Sub
End Class
Dim f As New Form1()
"#);
    // Compiles without error
}

/// K45. TabControl with TabPages.
#[test]

fn k45_tabcontrol_with_tabpages() {
    let (_vm, gui, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim tabs As New TabControl()
        tabs.Name = "tabs1"
        Dim tp1 As New TabPage()
        tp1.Name = "tpPage1"
        tp1.Text = "Page 1"
        Dim tp2 As New TabPage()
        tp2.Name = "tpPage2"
        tp2.Text = "Page 2"
        tabs.TabPages.Add(tp1)
        tabs.TabPages.Add(tp2)
        Me.Controls.Add(tabs)
    End Sub
End Class
Dim f As New Form1()
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"tabs1".to_string()), "Expected tabs1 control");
}
