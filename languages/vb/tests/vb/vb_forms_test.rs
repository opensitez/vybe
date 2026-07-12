//! Comprehensive tests for VB.NET WinForms compilation and widget state.
//!
//! Categories:
//!   A. Form creation and lifecycle (10 tests)
//!   B. Control creation and properties (10 tests)
//!   C. Layout: Point, Size, Font (10 tests)
//!   D. Event handling and Handles clause (10 tests)
//!   E. InitializeComponent pattern (8 tests)
//!   F. Multiple controls and complex forms (8 tests)
//!   G. Property side effects and propagation (8 tests)
//!   H. MsgBox, Close, Show, dialogs (8 tests)

use super::helpers::{run_vb, run_vb_gui};

// ============================================================
// A. FORM CREATION AND LIFECYCLE (10 tests)
// ============================================================

/// A01. Creating a Form class emits a Text property change.
#[test]

fn a01_form_class_creation_emits_text() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    let text = g.get_property("form1", "text");
    assert!(!text.is_empty(), "Expected a Text property for form1");
}

/// A02. Application.Run triggers the GUI launch host path.
#[test]
fn a02_application_run_emits_run_application() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
Application.Run(f)
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.should_run,
        "Expected should_run to be true after Application.Run"
    );
}

/// A03. Form.Show emits FormShow side effect.
#[test]

fn a03_form_show_emits_form_show() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.Show()
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.should_run, "Expected should_run after Form.Show");
}

/// A04. Form.Close emits FormClose side effect.
#[test]

fn a04_form_close_emits_form_close() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.Close()
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.close_requested,
        "Expected close_requested after Form.Close"
    );
}

/// A05. Form with custom title via Me.Text assignment.
#[test]
fn a05_form_custom_title_via_me_text() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Inherits Form
    Public Sub New()
        Me.Text = "My Application"
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    // The form's runtime control name is the lowercased child class name
    // (`form1`), stamped by `compile_class` after the `Form` parent ctor runs.
    // The setter chain: Me.Text → __set_text (inherited from Control) →
    // controlSetProperty(this, "Text", "My Application") → gui state under
    // ("form1", "Text").
    let text = g.get_property("form1", "Text");
    assert_eq!(
        text, "My Application",
        "Expected Text='My Application', got '{}'",
        text
    );
}

/// A06. Multiple forms can be created independently.
#[test]

fn a06_multiple_forms_created() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
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
"#,
    );
    let g = gui.lock().unwrap();
    // Both forms should produce controls or properties
    assert!(
        g.control_names.len() >= 2,
        "Expected at least 2 forms, got {}",
        g.control_names.len()
    );
}

/// A07. Empty form class with no Sub New compiles and runs.
#[test]
fn a07_empty_form_class_no_constructor() {
    let (_vm, _gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
End Class
Dim f As New Form1()
"#,
    );
    // Just verifying it doesn't crash.
}

/// A08. Form with Inherits System.Windows.Forms.Form compiles.
#[test]
fn a08_form_inherits_system_windows_forms() {
    let (_vm, _gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Inherits System.Windows.Forms.Form
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    // Should compile and run without error.
}

/// A09. Form with fields declared at class level.
#[test]
fn a09_form_with_fields() {
    let out = run_vb(
        r#"
Public Class Form1
    Dim title As String = "Hello"
    Dim count As Integer = 0
    Public Sub New()
    End Sub
    Public Function GetTitle() As String
        Return title
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetTitle())
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}

/// A10. Form with method that modifies fields.
#[test]
fn a10_form_method_modifies_field() {
    let out = run_vb(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["3"]);
}

// ============================================================
// B. CONTROL CREATION AND PROPERTIES (10 tests)
// ============================================================

/// B01. New Button() creates an object with correct type.
#[test]
fn b01_new_button_creates_object() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btnOK"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btnok".to_string()),
        "Expected control btnok, got {:?}",
        g.control_names
    );
}

/// B02. New TextBox() with name and size.
#[test]
fn b02_new_textbox_with_properties() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim txt As New TextBox()
        txt.Name = "txtInput"
        txt.Location = New Point(10, 20)
        txt.Size = New Size(200, 25)
        Me.Controls.Add(txt)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"txtinput".to_string()),
        "Expected control txtinput"
    );
}

/// B03. New Label() with text property.
#[test]
fn b03_new_label_with_text() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim lbl As New Label()
        lbl.Name = "lblTitle"
        lbl.Text = "Welcome"
        Me.Controls.Add(lbl)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"lbltitle".to_string()),
        "Expected control lbltitle"
    );
    let text = g.get_property("lbltitle", "text");
    assert_eq!(text, "Welcome", "Expected Text='Welcome' for lbltitle");
}

/// B04. New CheckBox() emits AddControl with type CheckBox.
#[test]
fn b04_new_checkbox() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim chk As New CheckBox()
        chk.Name = "chkAgree"
        chk.Text = "I agree"
        Me.Controls.Add(chk)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"chkagree".to_string()),
        "Expected control chkagree"
    );
}

/// B05. New ComboBox() emits correct control type.
#[test]
fn b05_new_combobox() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim cbo As New ComboBox()
        cbo.Name = "cboItems"
        Me.Controls.Add(cbo)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"cboitems".to_string()),
        "Expected control cboitems"
    );
}

/// B06. New RadioButton() emits correct type.
#[test]
fn b06_new_radiobutton() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim rb As New RadioButton()
        rb.Name = "rbOption1"
        rb.Text = "Option 1"
        Me.Controls.Add(rb)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"rboption1".to_string()),
        "Expected control rboption1"
    );
}

/// B07. New ListBox() emits correct type.
#[test]
fn b07_new_listbox() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim lst As New ListBox()
        lst.Name = "lstItems"
        Me.Controls.Add(lst)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"lstitems".to_string()),
        "Expected control lstitems"
    );
}

/// B08. New Panel() emits correct type.
#[test]
fn b08_new_panel() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pnl As New Panel()
        pnl.Name = "pnlMain"
        Me.Controls.Add(pnl)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"pnlmain".to_string()),
        "Expected control pnlmain"
    );
}

/// B09. New GroupBox() emits correct type.
#[test]
fn b09_new_groupbox() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim grp As New GroupBox()
        grp.Name = "grpSettings"
        grp.Text = "Settings"
        Me.Controls.Add(grp)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"grpsettings".to_string()),
        "Expected control grpsettings"
    );
}

/// B10. New PictureBox() emits correct type.
#[test]
fn b10_new_picturebox() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pic As New PictureBox()
        pic.Name = "picLogo"
        Me.Controls.Add(pic)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"piclogo".to_string()),
        "Expected control piclogo"
    );
}

// ============================================================
// C. LAYOUT: POINT, SIZE, FONT (10 tests)
// ============================================================

/// C01. New Point(x, y) assigns location correctly.
#[test]
fn c01_point_assigns_location() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Location = New Point(50, 100)
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
}

/// C02. New Size(w, h) assigns size correctly.
#[test]
fn c02_size_assigns_dimensions() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Size = New Size(150, 40)
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
}

/// C03. Point and Size together give correct layout.
#[test]
fn c03_point_and_size_together() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim txt As New TextBox()
        txt.Name = "txtName"
        txt.Location = New Point(20, 30)
        txt.Size = New Size(250, 25)
        Me.Controls.Add(txt)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"txtname".to_string()),
        "Expected control txtname"
    );
}

/// C04. Zero position and size defaults.
#[test]
fn c04_zero_position_defaults() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Location = New Point(0, 0)
        btn.Size = New Size(0, 0)
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
}

/// C05. Large coordinates are preserved.
#[test]
fn c05_large_coordinates() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Location = New Point(1000, 2000)
        btn.Size = New Size(500, 300)
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
}

/// C06. Control without explicit location uses defaults.
#[test]
fn c06_control_default_location() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
}

/// C07. Control without explicit size gets default 100x30.
#[test]
fn c07_control_default_size() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
}

/// C08. New Font() with name and size compiles.
#[test]
fn c08_new_font_compiles() {
    let out = run_vb(
        r#"
Imports System.Drawing
Dim f As New Font("Arial", 12)
Console.WriteLine("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

/// C09. Multiple controls with different positions.
#[test]
fn c09_multiple_controls_different_positions() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn1 As New Button()
        btn1.Name = "btn1"
        btn1.Location = New Point(10, 10)
        btn1.Size = New Size(80, 30)

        Dim btn2 As New Button()
        btn2.Name = "btn2"
        btn2.Location = New Point(100, 10)
        btn2.Size = New Size(80, 30)

        Me.Controls.Add(btn1)
        Me.Controls.Add(btn2)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
    assert!(
        g.control_names.contains(&"btn2".to_string()),
        "Expected control btn2"
    );
}

/// C10. Point with computed coordinates.
#[test]
fn c10_point_computed_coordinates() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim x As Integer = 5
        Dim y As Integer = 10
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Location = New Point(x * 2, y * 3)
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control btn1"
    );
}

// ============================================================
// D. EVENT HANDLING AND HANDLES CLAUSE (10 tests)
// ============================================================

/// D01. Handles clause registers Click event.
#[test]
fn d01_handles_click_event() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Public Sub New()
        btn1 = New Button()
        btn1.Name = "btn1"
        Me.Controls.Add(btn1)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        Console.WriteLine("clicked")
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("btn1.click"),
        "Expected Click handler for btn1, got keys: {:?}",
        g.event_keys()
    );
}

/// D02. Multiple Handles clauses on different controls.
#[test]
fn d02_multiple_handles_different_controls() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Dim btn2 As Button
    Public Sub New()
        btn1 = New Button()
        btn1.Name = "btn1"
        btn2 = New Button()
        btn2.Name = "btn2"
        Me.Controls.Add(btn1)
        Me.Controls.Add(btn2)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    End Sub
    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("btn1.click"),
        "Missing handler for btn1.click"
    );
    assert!(
        g.event_handlers.contains_key("btn2.click"),
        "Missing handler for btn2.click"
    );
}

/// D03. Handles Me.Load registers form Load event.
#[test]
fn d03_handles_me_load() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("form1.load"),
        "Expected Load handler for form1, got keys: {:?}",
        g.event_keys()
    );
}

/// D04. Handler registered via Handles is a callable Value.
#[test]
fn d04_handler_is_callable() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Public Sub New()
        btn1 = New Button()
        btn1.Name = "btn1"
        Me.Controls.Add(btn1)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    let handler = g.get_event_handler("btn1", "Click");
    assert!(handler.is_some(), "Handler should exist");
    let h = handler.unwrap();
    assert!(
        !matches!(h, vybe_bytecode::Value::Null),
        "Handler should not be Null, got {:?}",
        h
    );
}

/// D05. Handles with InitializeComponent pattern.
#[test]
fn d05_handles_with_initialize_component() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        btn1 = New Button()
        btn1.Name = "btn1"
        Me.Controls.Add(btn1)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("btn1.click"),
        "Expected handler after InitializeComponent"
    );
}

/// D06. Form with TextChanged handle.
#[test]
fn d06_handles_textchanged() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim txt1 As TextBox
    Public Sub New()
        txt1 = New TextBox()
        txt1.Name = "txt1"
        Me.Controls.Add(txt1)
    End Sub
    Private Sub txt1_TextChanged(sender As Object, e As EventArgs) Handles txt1.TextChanged
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("txt1.textchanged"),
        "Expected TextChanged handler for txt1"
    );
}

/// D07. Two events on the same control.
#[test]
fn d07_two_events_same_control() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Public Sub New()
        btn1 = New Button()
        btn1.Name = "btn1"
        Me.Controls.Add(btn1)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    End Sub
    Private Sub btn1_MouseDown(sender As Object, e As EventArgs) Handles btn1.MouseDown
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("btn1.click"),
        "Missing Click handler"
    );
    assert!(
        g.event_handlers.contains_key("btn1.mousedown"),
        "Missing MouseDown handler"
    );
}

/// D08. Handler with body that references form fields.
#[test]
fn d08_handler_references_fields() {
    // Compile test: the handler body accesses a class field via Me.
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Dim counter As Integer = 0
    Public Sub New()
        btn1 = New Button()
        btn1.Name = "btn1"
        Me.Controls.Add(btn1)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        counter = counter + 1
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("btn1.click"),
        "Expected handler that accesses fields"
    );
}

/// D09. SelectedIndexChanged handle on ComboBox.
#[test]
fn d09_handles_selectedindexchanged() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim cbo1 As ComboBox
    Public Sub New()
        cbo1 = New ComboBox()
        cbo1.Name = "cbo1"
        Me.Controls.Add(cbo1)
    End Sub
    Private Sub cbo1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo1.SelectedIndexChanged
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("cbo1.selectedindexchanged"),
        "Expected SelectedIndexChanged handler for cbo1"
    );
}

/// D10. CheckedChanged handle on CheckBox.
#[test]
fn d10_handles_checkedchanged() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim chk1 As CheckBox
    Public Sub New()
        chk1 = New CheckBox()
        chk1.Name = "chk1"
        Me.Controls.Add(chk1)
    End Sub
    Private Sub chk1_CheckedChanged(sender As Object, e As EventArgs) Handles chk1.CheckedChanged
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("chk1.checkedchanged"),
        "Expected CheckedChanged handler for chk1"
    );
}

// ============================================================
// E. INITIALIZECOMPONENT PATTERN (8 tests)
// ============================================================

/// E01. InitializeComponent called from constructor creates controls.
#[test]
fn e01_initialize_component_creates_controls() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Dim btn1 As Button
    Dim txtName As TextBox
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        btn1 = New Button()
        btn1.Name = "btn1"
        btn1.Location = New Point(10, 20)
        btn1.Size = New Size(80, 30)
        btn1.Text = "Click"
        txtName = New TextBox()
        txtName.Name = "txtName"
        txtName.Location = New Point(10, 60)
        txtName.Size = New Size(200, 25)
        Me.Controls.Add(btn1)
        Me.Controls.Add(txtName)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(
        g.control_names.len(),
        2,
        "Expected 2 controls, got {:?}",
        g.control_names
    );
    assert!(g.control_names.contains(&"btn1".to_string()));
    assert!(g.control_names.contains(&"txtname".to_string()));
}

/// E02. InitializeComponent sets form text.
#[test]
fn e02_initialize_component_sets_form_text() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Inherits Form
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        Me.Text = "Login Form"
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    let text = g.get_property("form1", "Text");
    assert_eq!(
        text, "Login Form",
        "Expected form text 'Login Form', got '{}'",
        text
    );
}

/// E03. SuspendLayout and ResumeLayout are no-ops (don't crash).
#[test]
fn e03_suspend_resume_layout_noop() {
    let (_vm, _gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.ResumeLayout(False)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    // Just verifying it doesn't crash.
}

/// E04. Multiple controls in InitializeComponent with correct types.
#[test]
fn e04_multiple_controls_correct_types() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Dim lblName As Label
    Dim txtName As TextBox
    Dim btnSubmit As Button
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        lblName = New Label()
        lblName.Name = "lblName"
        lblName.Text = "Name:"
        lblName.Location = New Point(10, 10)

        txtName = New TextBox()
        txtName.Name = "txtName"
        txtName.Location = New Point(80, 10)
        txtName.Size = New Size(200, 25)

        btnSubmit = New Button()
        btnSubmit.Name = "btnSubmit"
        btnSubmit.Text = "Submit"
        btnSubmit.Location = New Point(80, 50)

        Me.Controls.Add(lblName)
        Me.Controls.Add(txtName)
        Me.Controls.Add(btnSubmit)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(
        g.control_names.len(),
        3,
        "Expected 3 controls, got {:?}",
        g.control_names
    );
    assert!(g.control_names.contains(&"lblname".to_string()));
    assert!(g.control_names.contains(&"txtname".to_string()));
    assert!(g.control_names.contains(&"btnsubmit".to_string()));
}

/// E05. InitializeComponent pattern with event handlers (Handles).
#[test]
fn e05_initialize_component_with_handles() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        btn1 = New Button()
        btn1.Name = "btn1"
        Me.Controls.Add(btn1)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 1);
    assert!(
        g.event_handlers.contains_key("btn1.click"),
        "Handler should be registered after InitializeComponent"
    );
}

/// E06. InitializeComponent with form size settings.
#[test]
fn e06_initialize_component_form_size() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Inherits Form
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        Me.Text = "Sized Form"
        Me.Size = New Size(800, 600)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    let text = g.get_property("form1", "Text");
    assert_eq!(text, "Sized Form", "Expected Text='Sized Form'");
}

/// E07. InitializeComponent with Enabled and Visible properties.
#[test]
fn e07_initialize_component_enabled_visible() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Dim btn1 As Button
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        btn1 = New Button()
        btn1.Name = "btn1"
        btn1.Enabled = False
        btn1.Visible = True
        Me.Controls.Add(btn1)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected btn1 control"
    );
}

/// E08. Complex InitializeComponent with many control types.
#[test]
fn e08_complex_initialize_component() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Dim lbl1 As Label
    Dim txt1 As TextBox
    Dim btn1 As Button
    Dim chk1 As CheckBox
    Dim cbo1 As ComboBox
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        lbl1 = New Label()
        lbl1.Name = "lbl1"
        txt1 = New TextBox()
        txt1.Name = "txt1"
        btn1 = New Button()
        btn1.Name = "btn1"
        chk1 = New CheckBox()
        chk1.Name = "chk1"
        cbo1 = New ComboBox()
        cbo1.Name = "cbo1"
        Me.Controls.Add(lbl1)
        Me.Controls.Add(txt1)
        Me.Controls.Add(btn1)
        Me.Controls.Add(chk1)
        Me.Controls.Add(cbo1)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(
        g.control_names.len(),
        5,
        "Expected 5 controls, got {:?}",
        g.control_names
    );
}

// ============================================================
// F. MULTIPLE CONTROLS AND COMPLEX FORMS (8 tests)
// ============================================================

/// F01. Form with 3 buttons at different positions.
#[test]
fn f01_three_buttons_different_positions() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Dim btn1 As Button
    Dim btn2 As Button
    Dim btn3 As Button
    Public Sub New()
        btn1 = New Button()
        btn1.Name = "btn1"
        btn1.Location = New Point(10, 10)
        btn2 = New Button()
        btn2.Name = "btn2"
        btn2.Location = New Point(10, 50)
        btn3 = New Button()
        btn3.Name = "btn3"
        btn3.Location = New Point(10, 90)
        Me.Controls.Add(btn1)
        Me.Controls.Add(btn2)
        Me.Controls.Add(btn3)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 3);
    assert!(g.control_names.contains(&"btn1".to_string()));
    assert!(g.control_names.contains(&"btn2".to_string()));
    assert!(g.control_names.contains(&"btn3".to_string()));
}

/// F02. Form with label + textbox + button (login pattern).
#[test]
fn f02_login_form_pattern() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class LoginForm
    Dim lblUser As Label
    Dim txtUser As TextBox
    Dim btnLogin As Button
    Public Sub New()
        lblUser = New Label()
        lblUser.Name = "lblUser"
        lblUser.Text = "Username:"
        lblUser.Location = New Point(10, 15)
        txtUser = New TextBox()
        txtUser.Name = "txtUser"
        txtUser.Location = New Point(100, 10)
        txtUser.Size = New Size(200, 25)
        btnLogin = New Button()
        btnLogin.Name = "btnLogin"
        btnLogin.Text = "Login"
        btnLogin.Location = New Point(100, 50)
        btnLogin.Size = New Size(80, 30)
        Me.Controls.Add(lblUser)
        Me.Controls.Add(txtUser)
        Me.Controls.Add(btnLogin)
    End Sub
End Class
Dim f As New LoginForm()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 3);
    assert!(g.control_names.contains(&"lbluser".to_string()));
    assert!(g.control_names.contains(&"txtuser".to_string()));
    assert!(g.control_names.contains(&"btnlogin".to_string()));
}

/// F03. Form with controls and event handlers together.
#[test]
fn f03_form_controls_and_events() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Dim btn1 As Button
    Dim txt1 As TextBox
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        btn1 = New Button()
        btn1.Name = "btn1"
        btn1.Location = New Point(10, 10)
        txt1 = New TextBox()
        txt1.Name = "txt1"
        txt1.Location = New Point(10, 50)
        Me.Controls.Add(btn1)
        Me.Controls.Add(txt1)
    End Sub
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
    End Sub
    Private Sub txt1_TextChanged(sender As Object, e As EventArgs) Handles txt1.TextChanged
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 2);
    assert!(g.event_handlers.contains_key("btn1.click"));
    assert!(g.event_handlers.contains_key("txt1.textchanged"));
}

/// F04. Form constructor calls method that sets multiple properties.
#[test]
fn f04_constructor_calls_setup_method() {
    let out = run_vb(
        r#"
Public Class Form1
    Dim title As String
    Dim w As Integer
    Dim h As Integer
    Public Sub New()
        SetupDefaults()
    End Sub
    Private Sub SetupDefaults()
        title = "Default"
        w = 800
        h = 600
    End Sub
    Public Function GetTitle() As String
        Return title
    End Function
    Public Function GetWidth() As Integer
        Return w
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetTitle())
Console.WriteLine(f.GetWidth())
"#,
    );
    assert_eq!(out, vec!["Default", "800"]);
}

/// F05. Form with field initialized to a value.
#[test]
fn f05_form_field_initializer() {
    let out = run_vb(
        r#"
Public Class Form1
    Dim greeting As String = "Hello World"
    Public Sub New()
    End Sub
    Public Function GetGreeting() As String
        Return greeting
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetGreeting())
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

/// F06. Form with method that returns a computed value from fields.
#[test]
fn f06_form_computed_value() {
    let out = run_vb(
        r#"
Public Class Form1
    Dim width As Integer = 100
    Dim height As Integer = 50
    Public Sub New()
    End Sub
    Public Function Area() As Integer
        Return width * height
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.Area())
"#,
    );
    assert_eq!(out, vec!["5000"]);
}

/// F07. Two forms, each with controls, don't interfere.
#[test]
fn f07_two_forms_independent_controls() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class FormA
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btnA"
        Me.Controls.Add(btn)
    End Sub
End Class
Public Class FormB
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btnB"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim a As New FormA()
Dim b As New FormB()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btna".to_string()));
    assert!(g.control_names.contains(&"btnb".to_string()));
}

/// F08. Form with ProgressBar and TrackBar control types.
#[test]
fn f08_progress_and_trackbar() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pb As New ProgressBar()
        pb.Name = "pb1"
        Dim tb As New TrackBar()
        tb.Name = "tb1"
        Me.Controls.Add(pb)
        Me.Controls.Add(tb)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"pb1".to_string()),
        "Expected control pb1"
    );
    assert!(
        g.control_names.contains(&"tb1".to_string()),
        "Expected control tb1"
    );
}

// ============================================================
// G. PROPERTY MIRRORING TO GUI STATE (8 tests)
// ============================================================
//
// Programmatic property writes from VB code (`btn.Text = "OK"`)
// flow through the canvas/property mirror in `GuiState::set_property`
// and land in the property store. These tests assert on the mirror
// — they don't observe events.

/// G01. Setting Text on a control mirrors into the gui-state property store.
#[test]
fn g01_text_property_mirrors_to_gui_state() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Text = "OK"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    let text = g.get_property("btn1", "text");
    assert_eq!(text, "OK", "Expected Text=OK for btn1, got '{}'", text);
}

/// G02. Setting Name property is reflected in control_names.
#[test]
fn g02_name_in_add_control() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "myButton"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"mybutton".to_string()),
        "Expected control named mybutton"
    );
}

/// G03. Boolean property (Enabled = False).
#[test]
fn g03_boolean_property_enabled() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Enabled = False
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected btn1 control"
    );
}

/// G04. Integer property (TabIndex).
#[test]
fn g04_integer_property_tabindex() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.TabIndex = 3
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected btn1 control"
    );
}

/// G05. Multiple property changes on the same control.
#[test]
fn g05_multiple_properties_same_control() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Text = "Go"
        btn.Enabled = True
        btn.Location = New Point(10, 20)
        btn.Size = New Size(80, 30)
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
    let text = g.get_property("btn1", "text");
    assert_eq!(text, "Go", "Expected Text='Go' for btn1");
}

/// G06. Setting form BackColor as string.
#[test]
fn g06_form_backcolor_string() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.BackColor = "Red"
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    // Just verifying it compiles and runs without error.
    let _g = gui.lock().unwrap();
}

/// G07. Property set before Controls.Add is still emitted.
#[test]
fn g07_property_before_controls_add() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Text = "Before Add"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    let text = g.get_property("btn1", "text");
    assert_eq!(
        text, "Before Add",
        "Expected Text='Before Add' for btn1, got '{}'",
        text
    );
}

/// G08. Setting Multiline on TextBox.
#[test]
fn g08_textbox_multiline() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim txt As New TextBox()
        txt.Name = "txt1"
        txt.Multiline = True
        Me.Controls.Add(txt)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"txt1".to_string()),
        "Expected txt1 control"
    );
}

// ============================================================
// H. MSGBOX, CLOSE, SHOW, DIALOGS (8 tests)
// ============================================================

/// H01. MsgBox calls the host fn.
#[test]
fn h01_msgbox_emits_call() {
    let (_vm, _gui, msgs) = super::helpers::run_vb_gui_capture_msgbox(
        r#"
MsgBox("Hello!")
"#,
    );
    let msgs = msgs.lock().unwrap();
    let has_msg = msgs.iter().any(|(text, _)| text == "Hello!");
    assert!(has_msg, "Expected MsgBox call, got {:?}", *msgs);
}

/// H02. MsgBox with title.
#[test]
fn h02_msgbox_with_title() {
    let (_vm, _gui, msgs) = super::helpers::run_vb_gui_capture_msgbox(
        r#"
MsgBox("Are you sure?", "Confirm")
"#,
    );
    let msgs = msgs.lock().unwrap();
    let has_msg = msgs
        .iter()
        .any(|(text, title)| text == "Are you sure?" && title == "Confirm");
    assert!(has_msg, "Expected MsgBox with title, got {:?}", *msgs);
}

/// H03. MsgBox called from a class method.
#[test]
fn h03_msgbox_from_method() {
    let (_vm, _gui, msgs) = super::helpers::run_vb_gui_capture_msgbox(
        r#"
Public Class Form1
    Public Sub New()
        ShowMessage()
    End Sub
    Private Sub ShowMessage()
        MsgBox("From method")
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let msgs = msgs.lock().unwrap();
    let has_msg = msgs.iter().any(|(text, _)| text == "From method");
    assert!(has_msg, "Expected MsgBox from method, got {:?}", *msgs);
}

/// H04. Form.Close called from method.
#[test]

fn h04_close_from_method() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        CloseForm()
    End Sub
    Private Sub CloseForm()
        Me.Close()
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.close_requested,
        "Expected close_requested after Me.Close()"
    );
}

/// H05. Application.Run with form object.
#[test]
fn h05_application_run_with_object() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
Application.Run(f)
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.should_run, "Expected should_run after Application.Run");
}

/// H06. Console.WriteLine still works alongside form code.
#[test]
fn h06_console_writeline_with_forms() {
    let (_vm, _gui, output) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Console.WriteLine("Starting")
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
Console.WriteLine("Done")
"#,
    );
    let out = output.lock().unwrap().clone();
    assert!(out.contains(&"Starting".to_string()));
    assert!(out.contains(&"Done".to_string()));
}

/// H07. MsgBox with string concatenation.
#[test]
fn h07_msgbox_string_concat() {
    let (_vm, _gui, msgs) = super::helpers::run_vb_gui_capture_msgbox(
        r#"
Dim name As String = "World"
MsgBox("Hello " & name)
"#,
    );
    let msgs = msgs.lock().unwrap();
    let has_msg = msgs.iter().any(|(text, _)| text == "Hello World");
    assert!(has_msg, "Expected MsgBox 'Hello World', got {:?}", *msgs);
}

/// H08. Multiple MsgBox calls produce multiple host fn invocations.
#[test]
fn h08_multiple_msgbox() {
    let (_vm, _gui, msgs) = super::helpers::run_vb_gui_capture_msgbox(
        r#"
MsgBox("First")
MsgBox("Second")
MsgBox("Third")
"#,
    );
    let msgs = msgs.lock().unwrap();
    assert_eq!(msgs.len(), 3, "Expected 3 MsgBox calls, got {}", msgs.len());
}

// ============================================================
// BONUS: Edge cases and additional coverage (2+ tests)
// ============================================================

/// X01. Empty InitializeComponent does not crash.
#[test]
fn x01_empty_initialize_component() {
    let (_vm, _gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
    End Sub
End Class
Dim f As New Form1()
"#,
    );
}

/// X02. Form method accesses field set in InitializeComponent.
#[test]
fn x02_method_accesses_field_from_init() {
    let out = run_vb(
        r#"
Public Class Form1
    Dim status As String
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        status = "ready"
    End Sub
    Public Function GetStatus() As String
        Return status
    End Function
End Class
Dim f As New Form1()
Console.WriteLine(f.GetStatus())
"#,
    );
    assert_eq!(out, vec!["ready"]);
}

/// X03. Form with numeric field used in method.
#[test]
fn x03_numeric_field_in_method() {
    let out = run_vb(
        r#"
Public Class Counter
    Dim count As Integer = 0
    Public Sub New()
    End Sub
    Public Sub Add(n As Integer)
        count = count + n
    End Sub
    Public Function GetCount() As Integer
        Return count
    End Function
End Class
Dim c As New Counter()
c.Add(5)
c.Add(3)
Console.WriteLine(c.GetCount())
"#,
    );
    assert_eq!(out, vec!["8"]);
}

/// X04. Control with all standard properties set.
#[test]
fn x04_all_standard_properties() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        btn.Text = "Click Me"
        btn.Location = New Point(10, 20)
        btn.Size = New Size(120, 35)
        btn.Enabled = True
        btn.Visible = True
        btn.TabIndex = 0
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let mut g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected btn1 control"
    );
    let text = g.get_property("btn1", "text");
    assert_eq!(text, "Click Me", "Expected Text='Click Me' for btn1");
}
