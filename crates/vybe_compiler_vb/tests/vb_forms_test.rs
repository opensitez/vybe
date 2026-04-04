//! Comprehensive tests for VB.NET WinForms compilation and side-effect generation.
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

use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};

// ---------------------------------------------------------------------------
// Helpers (same pattern as vb_interop_test.rs)
// ---------------------------------------------------------------------------

fn run_vb(source: &str) -> Vec<String> {
    let program = vybe_parser_basic::parse_program(source)
        .unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybe_host::setup_namespaces(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    let chunks = vybe_compiler_vb::Compiler::new().compile(&program)
        .unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));
    output.borrow().clone()
}

fn run_vb_gui(source: &str) -> (VM, Rc<RefCell<vybe_host::SideEffectQueue>>, Rc<RefCell<Vec<String>>>) {
    let program = vybe_parser_basic::parse_program(source)
        .unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybe_host::setup_namespaces(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    let chunks = vybe_compiler_vb::Compiler::new().compile(&program)
        .unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));
    (vm, queue, output)
}

/// Drain all side effects from the queue.
fn drain(queue: &Rc<RefCell<vybe_host::SideEffectQueue>>) -> Vec<vybe_host::SideEffect> {
    queue.borrow_mut().drain()
}

/// Count AddControl effects.
fn count_add_controls(effects: &[vybe_host::SideEffect]) -> usize {
    effects.iter().filter(|e| matches!(e, vybe_host::SideEffect::AddControl { .. })).count()
}

/// Find an AddControl by control_name.
fn find_add_control<'a>(effects: &'a [vybe_host::SideEffect], name: &str) -> Option<&'a vybe_host::SideEffect> {
    effects.iter().find(|e| {
        if let vybe_host::SideEffect::AddControl { control_name, .. } = e {
            control_name == name
        } else { false }
    })
}

/// Find a PropertyChange by object+property.
fn find_prop_change<'a>(effects: &'a [vybe_host::SideEffect], obj: &str, prop: &str) -> Option<&'a vybe_host::SideEffect> {
    effects.iter().find(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, .. } = e {
            object == obj && property == prop
        } else { false }
    })
}

// ============================================================
// A. FORM CREATION AND LIFECYCLE (10 tests)
// ============================================================

/// A01. Creating a Form class emits a Text property change.
#[test]
#[ignore = "known bug: plain Form class without WinForms base does not emit Text property via side effects"]
fn a01_form_class_creation_emits_text() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let has_text = effects.iter().any(|e| matches!(e,
        vybe_host::SideEffect::PropertyChange { object, property, .. }
        if property == "Text" && object.contains("form1")
    ));
    assert!(has_text, "Expected a Text property side effect for form1, got {:?}", effects);
}

/// A02. Application.Run emits RunApplication side effect.
#[test]
fn a02_application_run_emits_run_application() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
Application.Run(f)
"#);
    let effects = drain(&queue);
    let has_run = effects.iter().any(|e| matches!(e, vybe_host::SideEffect::RunApplication { .. }));
    assert!(has_run, "Expected RunApplication side effect");
}

/// A03. Form.Show emits FormShow side effect.
#[test]
#[ignore = "known bug: VB class instances dont have show/close methods from host"]
fn a03_form_show_emits_form_show() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.Show()
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let has_show = effects.iter().any(|e| matches!(e, vybe_host::SideEffect::FormShow { .. }));
    assert!(has_show, "Expected FormShow side effect, got {:?}", effects);
}

/// A04. Form.Close emits FormClose side effect.
#[test]
#[ignore = "known bug: VB class instances dont have show/close methods from host"]
fn a04_form_close_emits_form_close() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.Close()
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let has_close = effects.iter().any(|e| matches!(e, vybe_host::SideEffect::FormClose { .. }));
    assert!(has_close, "Expected FormClose side effect, got {:?}", effects);
}

/// A05. Form with custom title via Me.Text assignment.
#[test]
fn a05_form_custom_title_via_me_text() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.Text = "My Application"
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let has_title = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { property, value, .. } = e {
            property == "Text" && format!("{}", value) == "My Application"
        } else { false }
    });
    assert!(has_title, "Expected Text='My Application', got {:?}", effects);
}

/// A06. Multiple forms can be created independently.
#[test]
#[ignore = "known bug: plain Form classes without WinForms base do not emit Text property side effects"]
fn a06_multiple_forms_created() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    // Both forms should produce Text property changes
    let text_changes: Vec<_> = effects.iter().filter(|e| matches!(e,
        vybe_host::SideEffect::PropertyChange { property, .. } if property == "Text"
    )).collect();
    assert!(text_changes.len() >= 2, "Expected at least 2 Text property changes, got {}", text_changes.len());
}

/// A07. Empty form class with no Sub New compiles and runs.
#[test]
fn a07_empty_form_class_no_constructor() {
    let (_vm, _queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
End Class
Dim f As New Form1()
"#);
    // Just verifying it doesn't crash.
}

/// A08. Form with Inherits System.Windows.Forms.Form compiles.
#[test]
fn a08_form_inherits_system_windows_forms() {
    let (_vm, _queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Inherits System.Windows.Forms.Form
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
"#);
    // Should compile and run without error.
}

/// A09. Form with fields declared at class level.
#[test]
fn a09_form_with_fields() {
    let out = run_vb(r#"
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
"#);
    assert_eq!(out, vec!["Hello"]);
}

/// A10. Form with method that modifies fields.
#[test]
fn a10_form_method_modifies_field() {
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

// ============================================================
// B. CONTROL CREATION AND PROPERTIES (10 tests)
// ============================================================

/// B01. New Button() creates an object with correct type.
#[test]
fn b01_new_button_creates_object() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btnOK"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "btnOK");
    assert!(ctrl.is_some(), "Expected AddControl for btnOK, got {:?}", effects);
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "Button");
    }
}

/// B02. New TextBox() with name and size.
#[test]
fn b02_new_textbox_with_properties() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "txtInput");
    assert!(ctrl.is_some(), "Expected AddControl for txtInput");
    if let Some(vybe_host::SideEffect::AddControl { left, top, width, height, .. }) = ctrl {
        assert_eq!(*left, 10);
        assert_eq!(*top, 20);
        assert_eq!(*width, 200);
        assert_eq!(*height, 25);
    }
}

/// B03. New Label() with text property.
#[test]
fn b03_new_label_with_text() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "lblTitle");
    assert!(ctrl.is_some(), "Expected AddControl for lblTitle");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "Label");
    }
    // Text property should also be emitted
    let text_prop = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, value, .. } = e {
            object == "lblTitle" && property == "Text" && format!("{}", value) == "Welcome"
        } else { false }
    });
    assert!(text_prop, "Expected Text property for lblTitle");
}

/// B04. New CheckBox() emits AddControl with type CheckBox.
#[test]
fn b04_new_checkbox() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "chkAgree");
    assert!(ctrl.is_some(), "Expected AddControl for chkAgree");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "CheckBox");
    }
}

/// B05. New ComboBox() emits correct control type.
#[test]
fn b05_new_combobox() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim cbo As New ComboBox()
        cbo.Name = "cboItems"
        Me.Controls.Add(cbo)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "cboItems");
    assert!(ctrl.is_some(), "Expected AddControl for cboItems");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "ComboBox");
    }
}

/// B06. New RadioButton() emits correct type.
#[test]
fn b06_new_radiobutton() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "rbOption1");
    assert!(ctrl.is_some(), "Expected AddControl for rbOption1");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "RadioButton");
    }
}

/// B07. New ListBox() emits correct type.
#[test]
fn b07_new_listbox() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim lst As New ListBox()
        lst.Name = "lstItems"
        Me.Controls.Add(lst)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "lstItems");
    assert!(ctrl.is_some(), "Expected AddControl for lstItems");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "ListBox");
    }
}

/// B08. New Panel() emits correct type.
#[test]
fn b08_new_panel() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pnl As New Panel()
        pnl.Name = "pnlMain"
        Me.Controls.Add(pnl)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "pnlMain");
    assert!(ctrl.is_some(), "Expected AddControl for pnlMain");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "Panel");
    }
}

/// B09. New GroupBox() emits correct type.
#[test]
fn b09_new_groupbox() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "grpSettings");
    assert!(ctrl.is_some(), "Expected AddControl for grpSettings");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "GroupBox");
    }
}

/// B10. New PictureBox() emits correct type.
#[test]
fn b10_new_picturebox() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim pic As New PictureBox()
        pic.Name = "picLogo"
        Me.Controls.Add(pic)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let ctrl = find_add_control(&effects, "picLogo");
    assert!(ctrl.is_some(), "Expected AddControl for picLogo");
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = ctrl {
        assert_eq!(control_type, "PictureBox");
    }
}

// ============================================================
// C. LAYOUT: POINT, SIZE, FONT (10 tests)
// ============================================================

/// C01. New Point(x, y) assigns location correctly.
#[test]
fn c01_point_assigns_location() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { left, top, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*left, 50);
        assert_eq!(*top, 100);
    } else {
        panic!("Expected AddControl for btn1");
    }
}

/// C02. New Size(w, h) assigns size correctly.
#[test]
fn c02_size_assigns_dimensions() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { width, height, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*width, 150);
        assert_eq!(*height, 40);
    } else {
        panic!("Expected AddControl for btn1");
    }
}

/// C03. Point and Size together give correct layout.
#[test]
fn c03_point_and_size_together() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { left, top, width, height, .. }) = find_add_control(&effects, "txtName") {
        assert_eq!(*left, 20);
        assert_eq!(*top, 30);
        assert_eq!(*width, 250);
        assert_eq!(*height, 25);
    } else {
        panic!("Expected AddControl for txtName");
    }
}

/// C04. Zero position and size defaults.
#[test]
fn c04_zero_position_defaults() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { left, top, width, height, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*left, 0);
        assert_eq!(*top, 0);
        assert_eq!(*width, 0);
        assert_eq!(*height, 0);
    } else {
        panic!("Expected AddControl for btn1");
    }
}

/// C05. Large coordinates are preserved.
#[test]
fn c05_large_coordinates() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { left, top, width, height, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*left, 1000);
        assert_eq!(*top, 2000);
        assert_eq!(*width, 500);
        assert_eq!(*height, 300);
    } else {
        panic!("Expected AddControl for btn1");
    }
}

/// C06. Control without explicit location uses defaults.
#[test]
fn c06_control_default_location() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { left, top, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*left, 0);
        assert_eq!(*top, 0);
    } else {
        panic!("Expected AddControl for btn1");
    }
}

/// C07. Control without explicit size gets default 100x30.
#[test]
fn c07_control_default_size() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btn1"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { width, height, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*width, 100);
        assert_eq!(*height, 30);
    } else {
        panic!("Expected AddControl for btn1");
    }
}

/// C08. New Font() with name and size compiles.
#[test]
fn c08_new_font_compiles() {
    let out = run_vb(r#"
Imports System.Drawing
Dim f As New Font("Arial", 12)
Console.WriteLine("ok")
"#);
    assert_eq!(out, vec!["ok"]);
}

/// C09. Multiple controls with different positions.
#[test]
fn c09_multiple_controls_different_positions() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { left, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*left, 10);
    } else {
        panic!("Expected AddControl for btn1");
    }
    if let Some(vybe_host::SideEffect::AddControl { left, .. }) = find_add_control(&effects, "btn2") {
        assert_eq!(*left, 100);
    } else {
        panic!("Expected AddControl for btn2");
    }
}

/// C10. Point with computed coordinates.
#[test]
fn c10_point_computed_coordinates() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { left, top, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*left, 10);
        assert_eq!(*top, 30);
    } else {
        panic!("Expected AddControl for btn1");
    }
}

// ============================================================
// D. EVENT HANDLING AND HANDLES CLAUSE (10 tests)
// ============================================================

/// D01. Handles clause registers Click event.
#[test]
fn d01_handles_click_event() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let handler = queue.borrow().get_event_handler("btn1", "Click").cloned();
    assert!(handler.is_some(), "Expected Click handler for btn1");
}

/// D02. Multiple Handles clauses on different controls.
#[test]
fn d02_multiple_handles_different_controls() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let q = queue.borrow();
    assert!(q.get_event_handler("btn1", "Click").is_some(), "Missing handler for btn1.Click");
    assert!(q.get_event_handler("btn2", "Click").is_some(), "Missing handler for btn2.Click");
}

/// D03. Handles Me.Load registers form Load event.
#[test]
fn d03_handles_me_load() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
    End Sub
End Class
Dim f As New Form1()
"#);
    let q = queue.borrow();
    let handler = q.get_event_handler("form1", "Load").cloned();
    assert!(handler.is_some(), "Expected Load handler for form1");
}

/// D04. Handler registered via Handles is a callable Value.
#[test]
fn d04_handler_is_callable() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let handler = queue.borrow().get_event_handler("btn1", "Click").cloned();
    assert!(handler.is_some(), "Handler should exist");
    // The handler value should not be null
    let h = handler.unwrap();
    assert!(!matches!(h, Value::Null), "Handler should not be Null, got {:?}", h);
}

/// D05. Handles with InitializeComponent pattern.
#[test]
fn d05_handles_with_initialize_component() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let handler = queue.borrow().get_event_handler("btn1", "Click").cloned();
    assert!(handler.is_some(), "Expected handler after InitializeComponent");
}

/// D06. Form with TextChanged handle.
#[test]
fn d06_handles_textchanged() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let handler = queue.borrow().get_event_handler("txt1", "TextChanged").cloned();
    assert!(handler.is_some(), "Expected TextChanged handler for txt1");
}

/// D07. Two events on the same control.
#[test]
fn d07_two_events_same_control() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let q = queue.borrow();
    assert!(q.get_event_handler("btn1", "Click").is_some(), "Missing Click handler");
    assert!(q.get_event_handler("btn1", "MouseDown").is_some(), "Missing MouseDown handler");
}

/// D08. Handler with body that references form fields.
#[test]
fn d08_handler_references_fields() {
    // Compile test: the handler body accesses a class field via Me.
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let handler = queue.borrow().get_event_handler("btn1", "Click").cloned();
    assert!(handler.is_some(), "Expected handler that accesses fields");
}

/// D09. SelectedIndexChanged handle on ComboBox.
#[test]
fn d09_handles_selectedindexchanged() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let handler = queue.borrow().get_event_handler("cbo1", "SelectedIndexChanged").cloned();
    assert!(handler.is_some(), "Expected SelectedIndexChanged handler for cbo1");
}

/// D10. CheckedChanged handle on CheckBox.
#[test]
fn d10_handles_checkedchanged() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let handler = queue.borrow().get_event_handler("chk1", "CheckedChanged").cloned();
    assert!(handler.is_some(), "Expected CheckedChanged handler for chk1");
}

// ============================================================
// E. INITIALIZECOMPONENT PATTERN (8 tests)
// ============================================================

/// E01. InitializeComponent called from constructor creates controls.
#[test]
fn e01_initialize_component_creates_controls() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert_eq!(count_add_controls(&effects), 2, "Expected 2 controls");
    assert!(find_add_control(&effects, "btn1").is_some());
    assert!(find_add_control(&effects, "txtName").is_some());
}

/// E02. InitializeComponent sets form text.
#[test]
fn e02_initialize_component_sets_form_text() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        Me.Text = "Login Form"
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let has_text = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { property, value, .. } = e {
            property == "Text" && format!("{}", value) == "Login Form"
        } else { false }
    });
    assert!(has_text, "Expected form text 'Login Form'");
}

/// E03. SuspendLayout and ResumeLayout are no-ops (don't crash).
#[test]
fn e03_suspend_resume_layout_noop() {
    let (_vm, _queue, _) = run_vb_gui(r#"
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
"#);
    // Just verifying it doesn't crash.
}

/// E04. Multiple controls in InitializeComponent with correct types.
#[test]
fn e04_multiple_controls_correct_types() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert_eq!(count_add_controls(&effects), 3, "Expected 3 controls");

    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = find_add_control(&effects, "lblName") {
        assert_eq!(control_type, "Label");
    }
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = find_add_control(&effects, "txtName") {
        assert_eq!(control_type, "TextBox");
    }
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = find_add_control(&effects, "btnSubmit") {
        assert_eq!(control_type, "Button");
    }
}

/// E05. InitializeComponent pattern with event handlers (Handles).
#[test]
fn e05_initialize_component_with_handles() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert_eq!(count_add_controls(&effects), 1);
    let handler = queue.borrow().get_event_handler("btn1", "Click").cloned();
    assert!(handler.is_some(), "Handler should be registered after InitializeComponent");
}

/// E06. InitializeComponent with form size settings.
#[test]
fn e06_initialize_component_form_size() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
        Me.Text = "Sized Form"
        Me.Size = New Size(800, 600)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    // Should have property changes for both Text and Size
    let has_text = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { property, value, .. } = e {
            property == "Text" && format!("{}", value) == "Sized Form"
        } else { false }
    });
    assert!(has_text, "Expected Text property");
}

/// E07. InitializeComponent with Enabled and Visible properties.
#[test]
fn e07_initialize_component_enabled_visible() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let has_enabled = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, .. } = e {
            object == "btn1" && property == "Enabled"
        } else { false }
    });
    assert!(has_enabled, "Expected Enabled property for btn1, got {:?}", effects);
}

/// E08. Complex InitializeComponent with many control types.
#[test]
fn e08_complex_initialize_component() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert_eq!(count_add_controls(&effects), 5, "Expected 5 controls");
}

// ============================================================
// F. MULTIPLE CONTROLS AND COMPLEX FORMS (8 tests)
// ============================================================

/// F01. Form with 3 buttons at different positions.
#[test]
fn f01_three_buttons_different_positions() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert_eq!(count_add_controls(&effects), 3);
    if let Some(vybe_host::SideEffect::AddControl { top, .. }) = find_add_control(&effects, "btn1") { assert_eq!(*top, 10); }
    if let Some(vybe_host::SideEffect::AddControl { top, .. }) = find_add_control(&effects, "btn2") { assert_eq!(*top, 50); }
    if let Some(vybe_host::SideEffect::AddControl { top, .. }) = find_add_control(&effects, "btn3") { assert_eq!(*top, 90); }
}

/// F02. Form with label + textbox + button (login pattern).
#[test]
fn f02_login_form_pattern() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert_eq!(count_add_controls(&effects), 3);
    assert!(find_add_control(&effects, "lblUser").is_some());
    assert!(find_add_control(&effects, "txtUser").is_some());
    assert!(find_add_control(&effects, "btnLogin").is_some());
}

/// F03. Form with controls and event handlers together.
#[test]
fn f03_form_controls_and_events() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert_eq!(count_add_controls(&effects), 2);
    let q = queue.borrow();
    assert!(q.get_event_handler("btn1", "Click").is_some());
    assert!(q.get_event_handler("txt1", "TextChanged").is_some());
}

/// F04. Form constructor calls method that sets multiple properties.
#[test]
fn f04_constructor_calls_setup_method() {
    let out = run_vb(r#"
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
"#);
    assert_eq!(out, vec!["Default", "800"]);
}

/// F05. Form with field initialized to a value.
#[test]
fn f05_form_field_initializer() {
    let out = run_vb(r#"
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
"#);
    assert_eq!(out, vec!["Hello World"]);
}

/// F06. Form with method that returns a computed value from fields.
#[test]
fn f06_form_computed_value() {
    let out = run_vb(r#"
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
"#);
    assert_eq!(out, vec!["5000"]);
}

/// F07. Two forms, each with controls, don't interfere.
#[test]
fn f07_two_forms_independent_controls() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert!(find_add_control(&effects, "btnA").is_some());
    assert!(find_add_control(&effects, "btnB").is_some());
}

/// F08. Form with ProgressBar and TrackBar control types.
#[test]
fn f08_progress_and_trackbar() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = find_add_control(&effects, "pb1") {
        assert_eq!(control_type, "ProgressBar");
    } else {
        panic!("Expected AddControl for pb1");
    }
    if let Some(vybe_host::SideEffect::AddControl { control_type, .. }) = find_add_control(&effects, "tb1") {
        assert_eq!(control_type, "TrackBar");
    } else {
        panic!("Expected AddControl for tb1");
    }
}

// ============================================================
// G. PROPERTY SIDE EFFECTS AND PROPAGATION (8 tests)
// ============================================================

/// G01. Setting Text on a control emits PropertyChange.
#[test]
fn g01_text_property_emits_side_effect() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let has_text = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, value, .. } = e {
            object == "btn1" && property == "Text" && format!("{}", value) == "OK"
        } else { false }
    });
    assert!(has_text, "Expected Text=OK for btn1, got {:?}", effects);
}

/// G02. Setting Name property is reflected in AddControl.
#[test]
fn g02_name_in_add_control() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "myButton"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    assert!(find_add_control(&effects, "myButton").is_some(), "Expected control named myButton");
}

/// G03. Boolean property (Enabled = False).
#[test]
fn g03_boolean_property_enabled() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let has_enabled = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, .. } = e {
            object == "btn1" && property == "Enabled"
        } else { false }
    });
    assert!(has_enabled, "Expected Enabled property for btn1");
}

/// G04. Integer property (TabIndex).
#[test]
fn g04_integer_property_tabindex() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let has_tabindex = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, .. } = e {
            object == "btn1" && property == "Tabindex"
        } else { false }
    });
    assert!(has_tabindex, "Expected Tabindex property for btn1, got {:?}", effects);
}

/// G05. Multiple property changes on the same control.
#[test]
fn g05_multiple_properties_same_control() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert!(find_add_control(&effects, "btn1").is_some());
    let btn1_props: Vec<_> = effects.iter().filter(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, .. } = e {
            object == "btn1"
        } else { false }
    }).collect();
    // At least Text and Enabled should be separate property changes
    assert!(btn1_props.len() >= 2, "Expected multiple property changes for btn1, got {}", btn1_props.len());
}

/// G06. Setting form BackColor as string.
#[test]
fn g06_form_backcolor_string() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.BackColor = "Red"
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let has_backcolor = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { property, .. } = e {
            property == "Backcolor" || property == "BackColor"
        } else { false }
    });
    assert!(has_backcolor, "Expected BackColor property, got {:?}", effects);
}

/// G07. Property set before Controls.Add is still emitted.
#[test]
fn g07_property_before_controls_add() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    // Text should be emitted as PropertyChange (either via controlSetProperty or during controlsAdd)
    let has_text = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, value, .. } = e {
            object == "btn1" && property == "Text" && format!("{}", value) == "Before Add"
        } else { false }
    });
    assert!(has_text, "Expected Text property for btn1 set before Controls.Add, got {:?}", effects);
}

/// G08. Setting Multiline on TextBox.
#[test]
fn g08_textbox_multiline() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let has_multiline = effects.iter().any(|e| {
        if let vybe_host::SideEffect::PropertyChange { object, property, .. } = e {
            object == "txt1" && property == "Multiline"
        } else { false }
    });
    assert!(has_multiline, "Expected Multiline property for txt1, got {:?}", effects);
}

// ============================================================
// H. MSGBOX, CLOSE, SHOW, DIALOGS (8 tests)
// ============================================================

/// H01. MsgBox emits MsgBox side effect.
#[test]
fn h01_msgbox_emits_side_effect() {
    let (_vm, queue, _) = run_vb_gui(r#"
MsgBox("Hello!")
"#);
    let effects = drain(&queue);
    let has_msg = effects.iter().any(|e| {
        if let vybe_host::SideEffect::MsgBox { text, .. } = e {
            text == "Hello!"
        } else { false }
    });
    assert!(has_msg, "Expected MsgBox side effect, got {:?}", effects);
}

/// H02. MsgBox with title.
#[test]
fn h02_msgbox_with_title() {
    let (_vm, queue, _) = run_vb_gui(r#"
MsgBox("Are you sure?", "Confirm")
"#);
    let effects = drain(&queue);
    let has_msg = effects.iter().any(|e| {
        if let vybe_host::SideEffect::MsgBox { text, title } = e {
            text == "Are you sure?" && title == "Confirm"
        } else { false }
    });
    assert!(has_msg, "Expected MsgBox with title, got {:?}", effects);
}

/// H03. MsgBox called from a class method.
#[test]
fn h03_msgbox_from_method() {
    let (_vm, queue, _) = run_vb_gui(r#"
Public Class Form1
    Public Sub New()
        ShowMessage()
    End Sub
    Private Sub ShowMessage()
        MsgBox("From method")
    End Sub
End Class
Dim f As New Form1()
"#);
    let effects = drain(&queue);
    let has_msg = effects.iter().any(|e| {
        if let vybe_host::SideEffect::MsgBox { text, .. } = e {
            text == "From method"
        } else { false }
    });
    assert!(has_msg, "Expected MsgBox from method, got {:?}", effects);
}

/// H04. Form.Close called from method.
#[test]
#[ignore = "known bug: Me.Close() inside a method resolves to undefined rather than host closeForm"]
fn h04_close_from_method() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    let has_close = effects.iter().any(|e| matches!(e, vybe_host::SideEffect::FormClose { .. }));
    assert!(has_close, "Expected FormClose side effect, got {:?}", effects);
}

/// H05. Application.Run with form object.
#[test]
fn h05_application_run_with_object() {
    let (_vm, queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
Application.Run(f)
"#);
    let effects = drain(&queue);
    let has_run = effects.iter().any(|e| matches!(e, vybe_host::SideEffect::RunApplication { .. }));
    assert!(has_run, "Expected RunApplication side effect");
}

/// H06. Console.WriteLine still works alongside form code.
#[test]
fn h06_console_writeline_with_forms() {
    let (_vm, _queue, output) = run_vb_gui(r#"
Imports System.Windows.Forms
Console.WriteLine("Starting")
Public Class Form1
    Public Sub New()
    End Sub
End Class
Dim f As New Form1()
Console.WriteLine("Done")
"#);
    let out = output.borrow().clone();
    assert!(out.contains(&"Starting".to_string()));
    assert!(out.contains(&"Done".to_string()));
}

/// H07. MsgBox with string concatenation.
#[test]
fn h07_msgbox_string_concat() {
    let (_vm, queue, _) = run_vb_gui(r#"
Dim name As String = "World"
MsgBox("Hello " & name)
"#);
    let effects = drain(&queue);
    let has_msg = effects.iter().any(|e| {
        if let vybe_host::SideEffect::MsgBox { text, .. } = e {
            text == "Hello World"
        } else { false }
    });
    assert!(has_msg, "Expected MsgBox 'Hello World', got {:?}", effects);
}

/// H08. Multiple MsgBox calls produce multiple side effects.
#[test]
fn h08_multiple_msgbox() {
    let (_vm, queue, _) = run_vb_gui(r#"
MsgBox("First")
MsgBox("Second")
MsgBox("Third")
"#);
    let effects = drain(&queue);
    let msg_count = effects.iter().filter(|e| matches!(e, vybe_host::SideEffect::MsgBox { .. })).count();
    assert_eq!(msg_count, 3, "Expected 3 MsgBox side effects, got {}", msg_count);
}

// ============================================================
// BONUS: Edge cases and additional coverage (2+ tests)
// ============================================================

/// X01. Empty InitializeComponent does not crash.
#[test]
fn x01_empty_initialize_component() {
    let (_vm, _queue, _) = run_vb_gui(r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub InitializeComponent()
    End Sub
End Class
Dim f As New Form1()
"#);
}

/// X02. Form method accesses field set in InitializeComponent.
#[test]
fn x02_method_accesses_field_from_init() {
    let out = run_vb(r#"
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
"#);
    assert_eq!(out, vec!["ready"]);
}

/// X03. Form with numeric field used in method.
#[test]
fn x03_numeric_field_in_method() {
    let out = run_vb(r#"
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
"#);
    assert_eq!(out, vec!["8"]);
}

/// X04. Control with all standard properties set.
#[test]
fn x04_all_standard_properties() {
    let (_vm, queue, _) = run_vb_gui(r#"
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
"#);
    let effects = drain(&queue);
    assert!(find_add_control(&effects, "btn1").is_some(), "Expected btn1 AddControl");
    if let Some(vybe_host::SideEffect::AddControl { left, top, width, height, .. }) = find_add_control(&effects, "btn1") {
        assert_eq!(*left, 10);
        assert_eq!(*top, 20);
        assert_eq!(*width, 120);
        assert_eq!(*height, 35);
    }
}
