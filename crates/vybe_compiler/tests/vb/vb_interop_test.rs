//! Comprehensive end-to-end tests for the VB compiler + bytecode VM interop.
//!
//! Categories:
//!   A. Class method resolution (10 tests)
//!   B. Parameter passing (8 tests)
//!   C. Namespace resolution (10 tests)
//!   D. Form lifecycle (8 tests)
//!   E. Event dispatch via invoke (8 tests)
//!   F. Object property access (8 tests)
//!   G. WinForms control properties (6 tests)
//!   H. Conversions and edge cases (8 tests)

use super::helpers::{run_vb, run_vb_gui, run_vb_vm};
use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value};

// ============================================================
// A. CLASS METHOD RESOLUTION (10 tests)
// ============================================================

/// A1. Constructor calls another method (InitializeComponent pattern)
#[test]
fn a01_constructor_calls_another_method() {
    let out = run_vb(
        r#"
Public Class Foo
    Dim value As String
    Public Sub New()
        Setup()
    End Sub
    Private Sub Setup()
        value = "initialized"
    End Sub
    Public Function GetValue() As String
        Return value
    End Function
End Class
Dim f As New Foo()
Console.WriteLine(f.GetValue())
"#,
    );
    assert_eq!(out, vec!["initialized"]);
}

/// A2. Method calls another method on same instance (bare name -> Me.method)
#[test]
fn a02_method_calls_another_method_on_same_instance() {
    let out = run_vb(
        r#"
Public Class Calc
    Dim total As Double
    Public Sub New()
        total = 0
    End Sub
    Public Sub Add(n As Double)
        total = total + n
    End Sub
    Public Sub AddTwice(n As Double)
        Add(n)
        Add(n)
    End Sub
    Public Function GetTotal() As Double
        Return total
    End Function
End Class
Dim c As New Calc()
c.AddTwice(5)
Console.WriteLine(c.GetTotal())
"#,
    );
    assert_eq!(out, vec!["10"]);
}

/// A3. Method accesses field without Me prefix
#[test]
fn a03_method_accesses_field_without_me_prefix() {
    let out = run_vb(
        r#"
Public Class Person
    Dim name As String
    Dim age As Integer
    Public Sub New(n As String, a As Integer)
        name = n
        age = a
    End Sub
    Public Function Describe() As String
        Return name & " is " & CStr(age)
    End Function
End Class
Dim p As New Person("Alice", 30)
Console.WriteLine(p.Describe())
"#,
    );
    assert_eq!(out, vec!["Alice is 30"]);
}

/// A4. Method with parameters called from another method
#[test]
fn a04_method_with_params_called_from_another_method() {
    let out = run_vb(
        r#"
Public Class Formatter
    Public Function Wrap(s As String, prefix As String, suffix As String) As String
        Return prefix & s & suffix
    End Function
    Public Function WrapBrackets(s As String) As String
        Return Wrap(s, "[", "]")
    End Function
End Class
Dim f As New Formatter()
Console.WriteLine(f.WrapBrackets("hello"))
"#,
    );
    assert_eq!(out, vec!["[hello]"]);
}

/// A5. Recursive method call on same instance
#[test]
fn a05_recursive_method_call() {
    let out = run_vb(
        r#"
Public Class MathHelper
    Public Function Factorial(n As Integer) As Integer
        If n <= 1 Then
            Return 1
        End If
        Return n * Factorial(n - 1)
    End Function
End Class
Dim m As New MathHelper()
Console.WriteLine(m.Factorial(5))
"#,
    );
    assert_eq!(out, vec!["120"]);
}

/// A6. Multiple methods calling each other in chain
#[test]
fn a06_method_chain() {
    let out = run_vb(
        r#"
Public Class Pipeline
    Public Function Step1(x As Integer) As Integer
        Return x + 1
    End Function
    Public Function Step2(x As Integer) As Integer
        Return Step1(x) * 2
    End Function
    Public Function Step3(x As Integer) As Integer
        Return Step2(x) + 10
    End Function
End Class
Dim p As New Pipeline()
Console.WriteLine(p.Step3(5))
"#,
    );
    // Step1(5) = 6, Step2(5) = 6*2 = 12, Step3(5) = 12 + 10 = 22
    assert_eq!(out, vec!["22"]);
}

/// A7. Constructor initializes fields, methods read them
#[test]
fn a07_constructor_initializes_fields_methods_read() {
    let out = run_vb(
        r#"
Public Class Config
    Dim host As String
    Dim port As Integer
    Dim secure As Boolean
    Public Sub New()
        host = "localhost"
        port = 8080
        secure = True
    End Sub
    Public Function GetUrl() As String
        If secure Then
            Return "https://" & host & ":" & CStr(port)
        Else
            Return "http://" & host & ":" & CStr(port)
        End If
    End Function
End Class
Dim cfg As New Config()
Console.WriteLine(cfg.GetUrl())
"#,
    );
    assert_eq!(out, vec!["https://localhost:8080"]);
}

/// A8. Method calls inherited method
/// KNOWN BUG: Inherited method calls via bare name fail with "undefined is not callable"
/// because the compiler doesn't resolve bare GetSpecies() to the parent class method
/// when called from a derived class method.
#[test]
fn a08_method_calls_inherited_method() {
    let out = run_vb(
        r#"
Public Class Animal
    Dim species As String
    Public Sub New(s As String)
        species = s
    End Sub
    Public Function GetSpecies() As String
        Return species
    End Function
End Class
Public Class Dog
    Inherits Animal
    Public Sub New()
        MyBase.New("Canine")
    End Sub
    Public Function Describe() As String
        Return "Dog: " & GetSpecies()
    End Function
End Class
Dim d As New Dog()
Console.WriteLine(d.Describe())
"#,
    );
    assert_eq!(out, vec!["Dog: Canine"]);
}

/// A9. Overridden method replaces parent
#[test]
fn a09_overridden_method_replaces_parent() {
    let out = run_vb(
        r#"
Public Class Base
    Public Function Name() As String
        Return "Base"
    End Function
End Class
Public Class Derived
    Inherits Base
    Public Overrides Function Name() As String
        Return "Derived"
    End Function
End Class
Dim d As New Derived()
Console.WriteLine(d.Name())
"#,
    );
    assert_eq!(out, vec!["Derived"]);
}

/// A10. Shared (static) method called without instance
#[test]
fn a10_shared_static_method() {
    let out = run_vb(
        r#"
Public Class MathUtils
    Public Shared Function Double(n As Integer) As Integer
        Return n * 2
    End Function
End Class
Console.WriteLine(MathUtils.Double(21))
"#,
    );
    assert_eq!(out, vec!["42"]);
}

// ============================================================
// B. PARAMETER PASSING (8 tests)
// ============================================================

/// B11. Zero args function
#[test]
fn b11_zero_args_function() {
    let out = run_vb(
        r#"
Function Hello() As String
    Return "world"
End Function
Console.WriteLine(Hello())
"#,
    );
    assert_eq!(out, vec!["world"]);
}

/// B12. Multiple args in correct order (a - b to verify order)
#[test]
fn b12_multiple_args_correct_order() {
    let out = run_vb(
        r#"
Function Subtract(a As Double, b As Double) As Double
    Return a - b
End Function
Console.WriteLine(Subtract(10, 3))
"#,
    );
    assert_eq!(out, vec!["7"]);
}

/// B13. Class method with multiple args
#[test]
fn b13_class_method_with_multiple_args() {
    let out = run_vb(
        r#"
Public Class Math2
    Public Function Add(a As Double, b As Double) As Double
        Return a + b
    End Function
End Class
Dim m As New Math2()
Console.WriteLine(m.Add(3, 4))
"#,
    );
    assert_eq!(out, vec!["7"]);
}

/// B14. Optional args (missing -> default/Null)
#[test]
fn b14_optional_args() {
    let out = run_vb(
        r#"
Function Greet(name As String, greeting As String) As String
    If greeting = "" Then
        Return "Hello " & name
    Else
        Return greeting & " " & name
    End If
End Function
Console.WriteLine(Greet("Alice", "Hi"))
Console.WriteLine(Greet("Bob", ""))
"#,
    );
    assert_eq!(out, vec!["Hi Alice", "Hello Bob"]);
}

/// B15. String args with special chars
#[test]
fn b15_string_args_with_special_chars() {
    let out = run_vb(
        r#"
Function Echo(s As String) As String
    Return s
End Function
Console.WriteLine(Echo("hello world"))
Console.WriteLine(Echo("it's"))
Console.WriteLine(Echo("a&b"))
"#,
    );
    assert_eq!(out, vec!["hello world", "it's", "a&b"]);
}

/// B16. Passing object as argument
/// KNOWN BUG: When an object is passed as a function argument and a method is called
/// on it inside the function, it resolves to null because the parameter binding
/// doesn't propagate the object's method table correctly.
#[test]
fn b16_passing_object_as_argument() {
    let out = run_vb(
        r#"
Public Class Item
    Dim name As String
    Public Sub New(n As String)
        name = n
    End Sub
    Public Function GetName() As String
        Return name
    End Function
End Class
Function Describe(item As Object) As String
    Return "Item: " & item.GetName()
End Function
Dim it As New Item("Widget")
Console.WriteLine(Describe(it))
"#,
    );
    assert_eq!(out, vec!["Item: Widget"]);
}

/// B17. Passing function result as argument
#[test]
fn b17_passing_function_result_as_argument() {
    let out = run_vb(
        r#"
Function Double(n As Integer) As Integer
    Return n * 2
End Function
Function AddOne(n As Integer) As Integer
    Return n + 1
End Function
Console.WriteLine(AddOne(Double(5)))
"#,
    );
    assert_eq!(out, vec!["11"]);
}

/// B18. Nested function calls as arguments
#[test]
fn b18_nested_function_calls_as_arguments() {
    let out = run_vb(
        r#"
Function Add(a As Integer, b As Integer) As Integer
    Return a + b
End Function
Function Mul(a As Integer, b As Integer) As Integer
    Return a * b
End Function
Console.WriteLine(Add(Mul(2, 3), Mul(4, 5)))
"#,
    );
    // 2*3 + 4*5 = 6 + 20 = 26
    assert_eq!(out, vec!["26"]);
}

// ============================================================
// C. NAMESPACE RESOLUTION (10 tests)
// ============================================================

/// C19. Math.Floor, Math.Ceiling, Math.Abs, Math.Sqrt
#[test]
fn c19_math_functions() {
    let out = run_vb(
        r#"
Console.WriteLine(Math.Floor(3.7))
Console.WriteLine(Math.Ceiling(3.2))
Console.WriteLine(Math.Abs(-5))
Console.WriteLine(Math.Sqrt(16))
"#,
    );
    assert_eq!(out, vec!["3", "4", "5", "4"]);
}

/// C20. Console.WriteLine with various types
#[test]
fn c20_console_writeline_various_types() {
    let out = run_vb(
        r#"
Console.WriteLine("hello")
Console.WriteLine(42)
Console.WriteLine(3.14)
Console.WriteLine(True)
Console.WriteLine(False)
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["hello", "42", "3.14", "true", "false"])
    );
}

/// C21. New System.Windows.Forms.Button()
#[test]
fn c21_new_system_windows_forms_button() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Dim btn As New Button()
Console.WriteLine(btn.__control_type)
"#,
    );
    assert_eq!(out, vec!["Button"]);
}

/// C22. New System.Drawing.Point(x, y) — verify x,y properties
#[test]
fn c22_new_system_drawing_point() {
    let out = run_vb(
        r#"
Imports System.Drawing
Dim p As New Point(10, 20)
Console.WriteLine(p.x)
Console.WriteLine(p.y)
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

/// C23. New System.Drawing.Size(w, h) — verify width,height
#[test]
fn c23_new_system_drawing_size() {
    let out = run_vb(
        r#"
Imports System.Drawing
Dim s As New Size(100, 50)
Console.WriteLine(s.width)
Console.WriteLine(s.height)
"#,
    );
    assert_eq!(out, vec!["100", "50"]);
}

/// C24. String.IsNullOrEmpty
#[test]
fn c24_string_is_null_or_empty() {
    let out = run_vb(
        r#"
Console.WriteLine(String.IsNullOrEmpty(""))
Console.WriteLine(String.IsNullOrEmpty("hello"))
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "false"])
    );
}

/// C25. Convert.ToString, Convert.ToInt32
/// KNOWN BUG: Convert.ToString/ToInt32 triggers "Unresolved import: vybe:math floor"
/// because the namespace resolution path incorrectly routes through math intrinsics.
#[test]
fn c25_convert_tostring_toint32() {
    let out = run_vb(
        r#"
Console.WriteLine(Convert.ToString(42))
Console.WriteLine(Convert.ToInt32("123"))
"#,
    );
    assert_eq!(out, vec!["42", "123"]);
}

/// C26. System.Math.Floor (fully qualified)
/// KNOWN BUG: System.Math.Floor (4-part dotted name) triggers "Unresolved import: vybe:math floor"
/// — the fully qualified path doesn't resolve to the WASM intrinsic correctly.
#[test]
fn c26_system_math_floor_fully_qualified() {
    let out = run_vb(
        r#"
Console.WriteLine(System.Math.Floor(9.9))
"#,
    );
    assert_eq!(out, vec!["9"]);
}

/// C27. Multiple namespace imports
#[test]
fn c27_multiple_namespace_imports() {
    let out = run_vb(
        r#"
Imports System.Drawing
Imports System.Windows.Forms
Dim pt As New Point(5, 10)
Dim btn As New Button()
btn.Location = pt
Console.WriteLine(btn.location.x)
Console.WriteLine(btn.location.y)
"#,
    );
    assert_eq!(out, vec!["5", "10"]);
}

/// C28. Math.Round (another common System.Math function)
#[test]
fn c28_math_round() {
    let out = run_vb(
        r#"
Console.WriteLine(Math.Round(3.6))
Console.WriteLine(Math.Round(3.4))
"#,
    );
    assert_eq!(out, vec!["4", "3"]);
}

// ============================================================
// D. FORM LIFECYCLE (8 tests)
// ============================================================

/// D29. Class with InitializeComponent — creates controls, sets properties
#[test]
fn d29_form_class_with_initialize_component() {
    let (_vm, gui, _output) = run_vb_gui(
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
Application.Run(f)
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(
        g.form.control_count(),
        2,
        "Expected 2 controls, got {}",
        g.form.control_count()
    );
    assert!(
        g.should_run,
        "Expected should_run to be true after Application.Run"
    );
}

/// D30. Controls.Add emits AddControl side effect with correct name/position/size
#[test]
fn d30_controls_add_emits_correct_properties() {
    let (_vm, gui, _) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Dim btn As Button
    Public Sub New()
        btn = New Button()
        btn.Name = "btn1"
        btn.Location = New Point(10, 20)
        btn.Size = New Size(80, 30)
        btn.Text = "OK"
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.form.control_count(), 1);
    assert!(
        g.control_names.contains(&"btn1".to_string()),
        "Expected control 'btn1'"
    );
}

/// D31. Handles clause registers event handler in queue
#[test]
fn d31_handles_clause_registers_event_handler() {
    let (_vm, gui, _output) = run_vb_gui(
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
        Console.WriteLine("clicked")
    End Sub
End Class

Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    let handler = g.get_event_handler("btn1", "Click");
    assert!(
        handler.is_some(),
        "Expected Click handler registered for btn1"
    );
}

/// D32. Multiple Handles on different controls
#[test]
fn d32_multiple_handles_on_different_controls() {
    let (_vm, gui, _output) = run_vb_gui(
        r#"
Imports System.Windows.Forms

Public Class Form1
    Dim btn1 As Button
    Dim btn2 As Button

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        btn1 = New Button()
        btn1.Name = "btn1"
        btn2 = New Button()
        btn2.Name = "btn2"
        Me.Controls.Add(btn1)
        Me.Controls.Add(btn2)
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        Console.WriteLine("btn1 clicked")
    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        Console.WriteLine("btn2 clicked")
    End Sub
End Class

Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.get_event_handler("btn1", "Click").is_some(),
        "btn1.Click handler"
    );
    assert!(
        g.get_event_handler("btn2", "Click").is_some(),
        "btn2.Click handler"
    );
}

/// D33. SuspendLayout/ResumeLayout don't crash
#[test]
fn d33_suspend_resume_layout_noop() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Public Class Form1
    Public Sub New()
        Me.SuspendLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()
        Console.WriteLine("ok")
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

/// D34. Application.Run triggers the GUI launch host path.
#[test]
fn d34_application_run_emits_run_application() {
    let (_vm, gui, _output) = run_vb_gui(
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

/// D35. Form with TextBox and Button — both controls added
#[test]
fn d35_form_with_textbox_and_button() {
    let (_vm, gui, _output) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Dim btn As Button
    Dim txt As TextBox
    Public Sub New()
        btn = New Button()
        btn.Name = "btnOK"
        btn.Location = New Point(10, 10)
        btn.Size = New Size(75, 23)
        txt = New TextBox()
        txt.Name = "txtInput"
        txt.Location = New Point(10, 40)
        txt.Size = New Size(200, 23)
        Me.Controls.Add(btn)
        Me.Controls.Add(txt)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.control_names.contains(&"btnok".to_string()),
        "Expected btnOK"
    );
    assert!(
        g.control_names.contains(&"txtinput".to_string()),
        "Expected txtInput"
    );
}

/// D36. Control properties set in InitializeComponent visible in gui state
#[test]
fn d36_control_properties_in_gui_state() {
    let (_vm, gui, _output) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim lbl As New Label()
        lbl.Name = "lblTitle"
        lbl.Location = New Point(50, 100)
        lbl.Size = New Size(150, 25)
        lbl.Text = "Welcome"
        Me.Controls.Add(lbl)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.form.control_count(), 1, "Expected 1 control");
    assert!(
        g.control_names.contains(&"lbltitle".to_string()),
        "Expected lblTitle control"
    );
}

// ============================================================
// E. EVENT DISPATCH VIA INVOKE (8 tests)
// ============================================================

/// E37. invoke() class method with Me — accesses fields correctly
#[test]
fn e37_invoke_class_method_with_me() {
    let (mut vm, output) = run_vb_vm(
        r#"
Public Class Counter
    Dim count As Integer
    Public Sub New()
        count = 0
    End Sub
    Public Sub Increment()
        count = count + 1
        Console.WriteLine(count)
    End Sub
End Class
Dim c As New Counter()
"#,
    );
    let instance = vm
        .globals
        .get("c")
        .cloned()
        .expect("Global 'c' should exist");
    let method = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("increment").cloned()
    } else {
        None
    }
    .expect("Instance should have 'increment' method");

    let result = vm.invoke(&method, &[instance.clone()]);
    assert!(result.is_ok(), "invoke should succeed: {:?}", result.err());

    let out = output.lock().unwrap();
    assert_eq!(out.last().map(|s| s.as_str()), Some("1"));
}

/// E38. invoke() without Me — fields are Null (documents the bug)
#[test]
fn e38_invoke_without_me_documents_bug() {
    let (mut vm, _output) = run_vb_vm(
        r#"
Public Class Greeter
    Dim name As String
    Public Sub New(n As String)
        name = n
    End Sub
    Public Function Greet() As String
        Return "Hello " & name
    End Function
End Class
Dim g As New Greeter("World")
"#,
    );
    let instance = vm.globals.get("g").cloned().unwrap();
    let method = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("greet").cloned()
    } else {
        None
    }
    .unwrap();

    // Invoke WITHOUT Me — method tries to access Me.name but Me is Null
    // This should either error or return wrong result — it should NOT crash
    let result = vm.invoke(&method, &[]);
    assert!(result.is_ok() || result.is_err(), "Should not panic");
}

/// E39. invoke() after vm.run — globals preserved
#[test]
fn e39_invoke_after_vm_run_globals_preserved() {
    let (mut vm, _output) = run_vb_vm(
        r#"
Dim x As Integer = 42
Function GetX() As Integer
    Return x
End Function
"#,
    );
    let x = vm.globals.get("x").cloned();
    assert!(x.is_some(), "Global x should exist after vm.run");

    // Invoke the function
    let func = vm.globals.get("getx").cloned();
    if let Some(f) = func {
        let result = vm.invoke(&f, &[]);
        assert!(result.is_ok(), "invoke should succeed");
    }
}

/// E40. invoke() class method that calls another method
#[test]
fn e40_invoke_class_method_that_calls_another_method() {
    let (mut vm, output) = run_vb_vm(
        r#"
Public Class Calculator
    Dim result As Integer
    Public Sub New()
        result = 0
    End Sub
    Public Sub AddTo(n As Integer)
        result = result + n
    End Sub
    Public Sub AddTwoNumbers(a As Integer, b As Integer)
        AddTo(a)
        AddTo(b)
        Console.WriteLine(result)
    End Sub
End Class
Dim calc As New Calculator()
"#,
    );
    let instance = vm.globals.get("calc").cloned().expect("Global 'calc'");
    let method = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("addtwonumbers").cloned()
    } else {
        None
    }
    .expect("addtwonumbers method");

    let result = vm.invoke(
        &method,
        &[instance.clone(), Value::F64(3.0), Value::F64(7.0)],
    );
    assert!(result.is_ok(), "invoke should succeed: {:?}", result.err());

    let out = output.lock().unwrap();
    assert_eq!(out.last().map(|s| s.as_str()), Some("10"));
}

/// E41. invoke() class method that modifies field, then read field
#[test]
fn e41_invoke_modifies_field_then_read() {
    let (mut vm, output) = run_vb_vm(
        r#"
Public Class Accum
    Dim val As Integer
    Public Sub New()
        val = 0
    End Sub
    Public Sub Bump()
        val = val + 10
    End Sub
    Public Sub Show()
        Console.WriteLine(val)
    End Sub
End Class
Dim a As New Accum()
"#,
    );
    let instance = vm.globals.get("a").cloned().expect("Global 'a'");

    // Invoke Bump
    let bump = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("bump").cloned()
    } else {
        None
    }
    .expect("bump method");
    vm.invoke(&bump, &[instance.clone()]).unwrap();

    // Invoke Show
    let show = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("show").cloned()
    } else {
        None
    }
    .expect("show method");
    vm.invoke(&show, &[instance.clone()]).unwrap();

    let out = output.lock().unwrap();
    assert_eq!(out.last().map(|s| s.as_str()), Some("10"));
}

/// E42. Top-level Dim creates global (accessible after vm.run)
#[test]
fn e42_top_level_dim_creates_global() {
    let (vm, _output) = run_vb_vm(
        r#"
Dim x As Integer = 42
Dim name As String = "test"
Dim flag As Boolean = True
"#,
    );
    assert!(vm.globals.get("x").is_some(), "x should be global");
    assert!(vm.globals.get("name").is_some(), "name should be global");
    assert!(vm.globals.get("flag").is_some(), "flag should be global");
}

/// E43. Class instance Dim creates global object
#[test]
fn e43_class_instance_dim_creates_global_object() {
    let (vm, _output) = run_vb_vm(
        r#"
Public Class Foo
    Dim x As Integer
    Public Sub New()
        x = 99
    End Sub
End Class
Dim f As New Foo()
"#,
    );
    let f = vm.globals.get("f").cloned();
    assert!(f.is_some(), "Top-level Dim f should be a global");
    assert!(
        matches!(f.unwrap(), Value::Object(_)),
        "f should be an Object"
    );
}

/// E44. invoke() method on object retrieved from global
#[test]
fn e44_invoke_method_on_global_object() {
    let (mut vm, output) = run_vb_vm(
        r#"
Public Class Greeter
    Dim prefix As String
    Public Sub New(p As String)
        prefix = p
    End Sub
    Public Sub SayHello(name As String)
        Console.WriteLine(prefix & " " & name)
    End Sub
End Class
Dim g As New Greeter("Hi")
"#,
    );
    let instance = vm.globals.get("g").cloned().expect("Global 'g'");
    let method = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("sayhello").cloned()
    } else {
        None
    }
    .expect("sayhello method");

    let result = vm.invoke(
        &method,
        &[instance.clone(), Value::String(Arc::from("World"))],
    );
    assert!(result.is_ok(), "invoke should succeed: {:?}", result.err());

    let out = output.lock().unwrap();
    assert_eq!(out.last().map(|s| s.as_str()), Some("Hi World"));
}

// ============================================================
// F. OBJECT PROPERTY ACCESS (8 tests)
// ============================================================

/// F45. Set and get property on object
#[test]
fn f45_set_and_get_property() {
    let out = run_vb(
        r#"
Public Class Box
    Dim content As String
    Public Sub New()
        content = "empty"
    End Sub
End Class
Dim b As New Box()
Console.WriteLine(b.content)
b.content = "full"
Console.WriteLine(b.content)
"#,
    );
    assert_eq!(out, vec!["empty", "full"]);
}

/// F46. Nested property chain (a.b.c)
#[test]
fn f46_nested_property_chain() {
    let out = run_vb(
        r#"
Public Class Inner
    Dim value As String
    Public Sub New(v As String)
        value = v
    End Sub
End Class
Public Class Outer
    Dim inner As Inner
    Public Sub New()
        inner = New Inner("deep")
    End Sub
End Class
Dim o As New Outer()
Console.WriteLine(o.inner.value)
"#,
    );
    assert_eq!(out, vec!["deep"]);
}

/// F47. Property set from outside class
#[test]
fn f47_property_set_from_outside() {
    let out = run_vb(
        r#"
Public Class Holder
    Dim data As String
    Public Sub New()
        data = ""
    End Sub
    Public Function GetData() As String
        Return data
    End Function
End Class
Dim h As New Holder()
h.data = "external"
Console.WriteLine(h.GetData())
"#,
    );
    assert_eq!(out, vec!["external"]);
}

/// F48. Method returns object, access its properties
#[test]
fn f48_method_returns_object_access_properties() {
    let out = run_vb(
        r#"
Public Class Pair
    Dim first As String
    Dim second As String
    Public Sub New(a As String, b As String)
        first = a
        second = b
    End Sub
End Class
Public Class Factory
    Public Function MakePair() As Pair
        Return New Pair("hello", "world")
    End Function
End Class
Dim f As New Factory()
Dim p As Pair = f.MakePair()
Console.WriteLine(p.first)
Console.WriteLine(p.second)
"#,
    );
    assert_eq!(out, vec!["hello", "world"]);
}

/// F49. Object stored in array, access via index
/// KNOWN BUG: Array indexed assignment fails with "null is not callable"
/// (same root cause as b74_array_operations).
#[test]
fn f49_object_in_array_access_via_index() {
    let out = run_vb(
        r#"
Public Class Item
    Dim name As String
    Public Sub New(n As String)
        name = n
    End Sub
End Class
Dim items(2) As Item
items(0) = New Item("first")
items(1) = New Item("second")
items(2) = New Item("third")
Console.WriteLine(items(0).name)
Console.WriteLine(items(2).name)
"#,
    );
    assert_eq!(out, vec!["first", "third"]);
}

/// F50. Multiple instances have independent state
#[test]
fn f50_multiple_instances_independent_state() {
    let out = run_vb(
        r#"
Public Class Counter
    Dim count As Integer
    Public Sub New(start As Integer)
        count = start
    End Sub
    Public Sub Inc()
        count = count + 1
    End Sub
    Public Function GetCount() As Integer
        Return count
    End Function
End Class
Dim a As New Counter(0)
Dim b As New Counter(100)
a.Inc()
a.Inc()
b.Inc()
Console.WriteLine(a.GetCount())
Console.WriteLine(b.GetCount())
"#,
    );
    assert_eq!(out, vec!["2", "101"]);
}

/// F51. Property Get/Set (VB property with getter/setter)
/// KNOWN BUG: The parser does not support Property with explicit Get/Set blocks.
/// It only supports auto-property syntax (Public Property Name As String).
#[test]
fn f51_property_get_set() {
    let out = run_vb(
        r#"
Public Class Person
    Dim _name As String
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
    Public Sub New()
        _name = "unknown"
    End Sub
End Class
Dim p As New Person()
Console.WriteLine(p.Name)
p.Name = "Alice"
Console.WriteLine(p.Name)
"#,
    );
    assert_eq!(out, vec!["unknown", "Alice"]);
}

/// F52. Property with backing field pattern
/// KNOWN BUG: Same as F51 — parser does not support Property with Get/Set blocks.
#[test]
fn f52_property_with_backing_field() {
    let out = run_vb(
        r#"
Public Class Temperature
    Dim _celsius As Double
    Public Property Celsius As Double
        Get
            Return _celsius
        End Get
        Set(value As Double)
            _celsius = value
        End Set
    End Property
    Public Function GetFahrenheit() As Double
        Return _celsius * 9 / 5 + 32
    End Function
    Public Sub New(c As Double)
        _celsius = c
    End Sub
End Class
Dim t As New Temperature(100)
Console.WriteLine(t.Celsius)
Console.WriteLine(t.GetFahrenheit())
t.Celsius = 0
Console.WriteLine(t.GetFahrenheit())
"#,
    );
    assert_eq!(out, vec!["100", "212", "32"]);
}

// ============================================================
// G. WINFORMS CONTROL PROPERTIES (6 tests)
// ============================================================

/// G53. Button Location = New Point(x,y) — read back x,y
#[test]
fn g53_button_location_readback() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Dim btn As New Button()
btn.Location = New Point(40, 100)
Console.WriteLine(btn.location.x)
Console.WriteLine(btn.location.y)
"#,
    );
    assert_eq!(out, vec!["40", "100"]);
}

/// G54. Button Size = New Size(w,h) — read back w,h
#[test]
fn g54_button_size_readback() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Dim btn As New Button()
btn.Size = New Size(60, 30)
Console.WriteLine(btn.size.width)
Console.WriteLine(btn.size.height)
"#,
    );
    assert_eq!(out, vec!["60", "30"]);
}

/// G55. Button Text, Name properties
#[test]
fn g55_button_text_name_properties() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Dim btn As New Button()
btn.Text = "Click Me"
btn.Name = "btnSubmit"
Console.WriteLine(btn.text)
Console.WriteLine(btn.name)
"#,
    );
    assert_eq!(out, vec!["Click Me", "btnSubmit"]);
}

/// G56. TextBox with Location, Size, Name
#[test]
fn g56_textbox_properties() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Dim txt As New TextBox()
txt.Name = "txtInput"
txt.Location = New Point(15, 25)
txt.Size = New Size(180, 22)
Console.WriteLine(txt.name)
Console.WriteLine(txt.location.x)
Console.WriteLine(txt.location.y)
Console.WriteLine(txt.size.width)
Console.WriteLine(txt.size.height)
"#,
    );
    assert_eq!(out, vec!["txtInput", "15", "25", "180", "22"]);
}

/// G57. Control created then properties set then Controls.Add — all correct in gui state
#[test]
fn g57_control_created_properties_set_then_add() {
    let (_vm, gui, _output) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn As New Button()
        btn.Name = "btnTest"
        btn.Text = "Test"
        btn.Location = New Point(25, 35)
        btn.Size = New Size(100, 40)
        Me.Controls.Add(btn)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.form.control_count(), 1);
    assert!(
        g.control_names.contains(&"btntest".to_string()),
        "Expected control 'btnTest'"
    );
}

/// G58. Multiple controls added — all appear in gui state
#[test]
fn g58_multiple_controls_all_appear() {
    let (_vm, gui, _output) = run_vb_gui(
        r#"
Imports System.Windows.Forms
Imports System.Drawing
Public Class Form1
    Public Sub New()
        Dim btn1 As New Button()
        btn1.Name = "btn1"
        btn1.Location = New Point(10, 10)
        btn1.Size = New Size(75, 23)

        Dim btn2 As New Button()
        btn2.Name = "btn2"
        btn2.Location = New Point(90, 10)
        btn2.Size = New Size(75, 23)

        Dim txt1 As New TextBox()
        txt1.Name = "txt1"
        txt1.Location = New Point(10, 40)
        txt1.Size = New Size(155, 23)

        Me.Controls.Add(btn1)
        Me.Controls.Add(btn2)
        Me.Controls.Add(txt1)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 3);
    assert!(g.control_names.contains(&"btn1".to_string()));
    assert!(g.control_names.contains(&"btn2".to_string()));
    assert!(g.control_names.contains(&"txt1".to_string()));
}

// ============================================================
// H. CONVERSIONS AND EDGE CASES (8 tests)
// ============================================================

/// H59. Val("42"), Val("abc"), Val("")
#[test]
fn h59_val_conversions() {
    let out = run_vb(
        r#"
Console.WriteLine(Val("42"))
Console.WriteLine(Val("3.14"))
Console.WriteLine(Val("abc"))
Console.WriteLine(Val(""))
"#,
    );
    assert_eq!(out, vec!["42", "3.14", "0", "0"]);
}

/// H60. CStr(42), CStr(3.14), CStr(True)
#[test]
fn h60_cstr_conversions() {
    let out = run_vb(
        r#"
Console.WriteLine(CStr(42))
Console.WriteLine(CStr(3.14))
Console.WriteLine(CStr(True))
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["42", "3.14", "true"])
    );
}

/// H61. CInt("42"), CDbl("3.14")
/// KNOWN BUG: CInt("42") returns NaN instead of 42 — string-to-integer conversion
/// doesn't parse the string correctly, treating it as a direct numeric cast.
#[test]
fn h61_cint_cdbl_conversions() {
    let out = run_vb(
        r#"
Console.WriteLine(CInt("42"))
Console.WriteLine(CDbl("3.14"))
"#,
    );
    assert_eq!(out, vec!["42", "3.14"]);
}

/// H62. String comparison <=, >=
#[test]
fn h62_string_comparison_le_ge() {
    let out = run_vb(
        r#"
Console.WriteLine("a" <= "b")
Console.WriteLine("b" >= "a")
Console.WriteLine("abc" <= "abc")
Console.WriteLine("z" <= "a")
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "true", "true", "false"])
    );
}

/// H63. Boolean And, Or, Not
#[test]
fn h63_boolean_and_or_not() {
    let out = run_vb(
        r#"
Console.WriteLine(True And True)
Console.WriteLine(True And False)
Console.WriteLine(False Or True)
Console.WriteLine(False Or False)
Console.WriteLine(Not True)
Console.WriteLine(Not False)
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "false", "true", "false", "false", "true"])
    );
}

/// H64. AndAlso short-circuit (second not evaluated)
#[test]
fn h64_andalso_short_circuit() {
    // If AndAlso short-circuits, the second function should not be called
    let out = run_vb(
        r#"
Dim called As Boolean = False
Function SideEffect() As Boolean
    called = True
    Return True
End Function
Dim result As Boolean = False AndAlso SideEffect()
Console.WriteLine(called)
Console.WriteLine(result)
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["false", "false"])
    );
}

/// H65. String concatenation & with mixed types
#[test]
fn h65_string_concatenation_mixed_types() {
    let out = run_vb(
        r#"
Console.WriteLine("Count: " & 42)
Console.WriteLine("Pi: " & 3.14)
Console.WriteLine("Active: " & True)
Console.WriteLine("Hello" & " " & "World")
"#,
    );
    assert_eq!(
        out,
        vec!["Count: 42", "Pi: 3.14", "Active: true", "Hello World"]
    );
}

/// H66. Integer division \
#[test]
fn h66_integer_division() {
    let out = run_vb(
        r#"
Console.WriteLine(7 \ 2)
Console.WriteLine(10 \ 3)
Console.WriteLine(100 \ 7)
"#,
    );
    assert_eq!(out, vec!["3", "3", "14"]);
}

// ============================================================
// BONUS: Additional edge cases and patterns
// ============================================================

/// B67. Three-level deep method call chain across classes
/// KNOWN BUG: Calling methods on fields that are objects of other classes fails
/// with "null is not callable" because the field's method table isn't properly
/// resolved when calling a.GetVal() inside B's method.
#[test]

fn b67_cross_class_method_chain() {
    let out = run_vb(
        r#"
Public Class A
    Public Function GetVal() As Integer
        Return 5
    End Function
End Class
Public Class B
    Dim a As A
    Public Sub New()
        a = New A()
    End Sub
    Public Function GetDouble() As Integer
        Return a.GetVal() * 2
    End Function
End Class
Public Class C
    Dim b As B
    Public Sub New()
        b = New B()
    End Sub
    Public Function GetTriple() As Integer
        Return b.GetDouble() + b.GetDouble() + b.GetDouble()
    End Function
End Class
Dim c As New C()
Console.WriteLine(c.GetTriple())
"#,
    );
    // 5*2 + 5*2 + 5*2 = 30
    assert_eq!(out, vec!["30"]);
}

/// B68. For loop basic
#[test]
fn b68_for_loop() {
    let out = run_vb(
        r#"
Dim total As Integer = 0
For i As Integer = 1 To 5
    total = total + i
Next
Console.WriteLine(total)
"#,
    );
    assert_eq!(out, vec!["15"]);
}

/// B69. If/ElseIf/Else
#[test]
fn b69_if_elseif_else() {
    let out = run_vb(
        r#"
Function Classify(n As Integer) As String
    If n > 0 Then
        Return "positive"
    ElseIf n < 0 Then
        Return "negative"
    Else
        Return "zero"
    End If
End Function
Console.WriteLine(Classify(5))
Console.WriteLine(Classify(-3))
Console.WriteLine(Classify(0))
"#,
    );
    assert_eq!(out, vec!["positive", "negative", "zero"]);
}

/// B70. While loop
#[test]
fn b70_while_loop() {
    let out = run_vb(
        r#"
Dim n As Integer = 1
Dim result As Integer = 1
While n <= 5
    result = result * n
    n = n + 1
End While
Console.WriteLine(result)
"#,
    );
    // 1*1*2*3*4*5 = 120
    assert_eq!(out, vec!["120"]);
}

/// B71. Select Case
#[test]
fn b71_select_case() {
    let out = run_vb(
        r#"
Function DayName(d As Integer) As String
    Select Case d
        Case 1
            Return "Monday"
        Case 2
            Return "Tuesday"
        Case 3
            Return "Wednesday"
        Case Else
            Return "Other"
    End Select
End Function
Console.WriteLine(DayName(1))
Console.WriteLine(DayName(3))
Console.WriteLine(DayName(7))
"#,
    );
    assert_eq!(out, vec!["Monday", "Wednesday", "Other"]);
}

/// B72. String functions: Len, Left, Right, Mid, UCase, LCase, Trim
#[test]
fn b72_string_functions() {
    let out = run_vb(
        r#"
Console.WriteLine(Len("hello"))
Console.WriteLine(Left("hello", 3))
Console.WriteLine(Right("hello", 2))
Console.WriteLine(Mid("hello", 2, 3))
Console.WriteLine(UCase("hello"))
Console.WriteLine(LCase("HELLO"))
Console.WriteLine(Trim("  hi  "))
"#,
    );
    assert_eq!(out, vec!["5", "hel", "lo", "ell", "HELLO", "hello", "hi"]);
}

/// B73. Math operations: Mod, power (^)
#[test]
fn b73_math_mod_power() {
    let out = run_vb(
        r#"
Console.WriteLine(10 Mod 3)
Console.WriteLine(2 ^ 10)
"#,
    );
    assert_eq!(out, vec!["1", "1024"]);
}

/// B74. Array operations
/// KNOWN BUG: Dim arr(4) As Integer with indexed assignment arr(i) = value
/// fails with "null is not callable" — array index assignment syntax issue.
#[test]
fn b74_array_operations() {
    let out = run_vb(
        r#"
Dim arr(4) As Integer
For i As Integer = 0 To 4
    arr(i) = i * 10
Next
Console.WriteLine(arr(0))
Console.WriteLine(arr(2))
Console.WriteLine(arr(4))
"#,
    );
    assert_eq!(out, vec!["0", "20", "40"]);
}

/// B75. Do While loop
#[test]
fn b75_do_while_loop() {
    let out = run_vb(
        r#"
Dim x As Integer = 0
Do While x < 3
    x = x + 1
Loop
Console.WriteLine(x)
"#,
    );
    assert_eq!(out, vec!["3"]);
}

/// B76. Nested class instantiation in expression
#[test]
fn b76_nested_class_in_expression() {
    let out = run_vb(
        r#"
Public Class Wrapper
    Dim val As Integer
    Public Sub New(v As Integer)
        val = v
    End Sub
    Public Function GetVal() As Integer
        Return val
    End Function
End Class
Console.WriteLine(New Wrapper(42).GetVal())
"#,
    );
    assert_eq!(out, vec!["42"]);
}

/// B77. Complex expression in method arg
#[test]
fn b77_complex_expression_in_method_arg() {
    let out = run_vb(
        r#"
Public Class Formatter
    Public Function Format(n As Double) As String
        Return "Value: " & CStr(n)
    End Function
End Class
Dim f As New Formatter()
Console.WriteLine(f.Format(2 + 3 * 4))
"#,
    );
    assert_eq!(out, vec!["Value: 14"]);
}

/// B78. Math.Max and Math.Min
#[test]
fn b78_math_max_min() {
    let out = run_vb(
        r#"
Console.WriteLine(Math.Max(5, 10))
Console.WriteLine(Math.Min(5, 10))
Console.WriteLine(Math.Max(-1, -5))
"#,
    );
    assert_eq!(out, vec!["10", "5", "-1"]);
}

/// B79. Comparison operators on numbers
#[test]
fn b79_comparison_operators() {
    let out = run_vb(
        r#"
Console.WriteLine(5 > 3)
Console.WriteLine(3 > 5)
Console.WriteLine(5 >= 5)
Console.WriteLine(3 < 5)
Console.WriteLine(5 < 3)
Console.WriteLine(5 <= 5)
Console.WriteLine(5 = 5)
Console.WriteLine(5 <> 3)
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&[
            "true", "false", "true", "true", "false", "true", "true", "true"
        ])
    );
}

/// B80. String equality
#[test]
fn b80_string_equality() {
    let out = run_vb(
        r#"
Console.WriteLine("hello" = "hello")
Console.WriteLine("hello" = "world")
Console.WriteLine("hello" <> "world")
"#,
    );
    assert_eq!(
        out,
        super::helpers::dotnet_expected_lines(&["true", "false", "true"])
    );
}
