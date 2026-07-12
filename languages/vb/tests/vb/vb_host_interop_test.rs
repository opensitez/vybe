/// Tests for VB → host function interop: objects crossing the boundary,
/// namespace resolution, host returning objects to VB, callbacks.
use super::helpers::{run_vb, run_vb_vm};
use std::sync::Arc;
use vybe_bytecode::Value;

// ============================================================
// A. HOST RETURNS OBJECT TO VB
// ============================================================

#[test]
fn host_list_create_add_count() {
    let out = run_vb(
        r#"
Dim list As New List(Of String)
list.Add("a")
list.Add("b")
list.Add("c")
Console.WriteLine(list.Count)
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn host_list_iterate() {
    let out = run_vb(
        r#"
Dim list As New List(Of String)
list.Add("x")
list.Add("y")
Dim total As Integer = 0
For Each item In list
    total = total + 1
Next
Console.WriteLine(total)
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn host_dictionary_create_add_access() {
    let out = run_vb(
        r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("name", "Alice")
Console.WriteLine(dict.Item("name"))
"#,
    );
    assert_eq!(out, vec!["Alice"]);
}

#[test]
fn host_dictionary_count() {
    let out = run_vb(
        r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("a", "1")
dict.Add("b", "2")
Console.WriteLine(dict.Count)
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn host_queue_enqueue_dequeue() {
    let out = run_vb(
        r#"
Dim q As New Queue(Of String)
q.Enqueue("first")
q.Enqueue("second")
Console.WriteLine(q.Dequeue())
Console.WriteLine(q.Dequeue())
"#,
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn host_stack_push_pop() {
    let out = run_vb(
        r#"
Dim s As New Stack(Of String)
s.Push("a")
s.Push("b")
Console.WriteLine(s.Pop())
Console.WriteLine(s.Pop())
"#,
    );
    assert_eq!(out, vec!["b", "a"]);
}

// ============================================================
// B. VB OBJECT PASSED TO HOST AND BACK
// ============================================================

#[test]
fn vb_class_instance_persists_in_global() {
    let (vm, _) = run_vb_vm(
        r#"
Public Class Dog
    Dim name As String
    Public Sub New(n As String)
        name = n
    End Sub
End Class
Dim d As New Dog("Rex")
"#,
    );
    let d = vm.globals.get("d").cloned();
    assert!(d.is_some(), "Dog instance should be in globals");
    if let Some(Value::Object(obj)) = d {
        let o = obj.lock().unwrap();
        let name = o.properties.get("name").map(|v| format!("{}", v));
        assert_eq!(name.as_deref(), Some("Rex"));
    } else {
        panic!("Expected Object");
    }
}

#[test]
fn vb_class_method_invokable_from_rust() {
    let (mut vm, output) = run_vb_vm(
        r#"
Public Class Greeter
    Dim prefix As String
    Public Sub New(p As String)
        prefix = p
    End Sub
    Public Sub Greet(name As String)
        Console.WriteLine(prefix & " " & name)
    End Sub
End Class
Dim g As New Greeter("Hello")
"#,
    );
    let instance = vm.globals.get("g").cloned().unwrap();
    let method = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("greet").cloned()
    } else {
        None
    }
    .unwrap();

    vm.invoke(
        &method,
        &[instance.clone(), Value::String(Arc::from("World"))],
    )
    .unwrap();
    assert_eq!(
        output.lock().unwrap().last().map(|s| s.as_str()),
        Some("Hello World")
    );
}

// ============================================================
// C. NAMESPACE RESOLUTION END-TO-END
// ============================================================

#[test]
fn namespace_math_floor() {
    let out = run_vb("Console.WriteLine(Math.Floor(3.7))");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn namespace_math_ceiling() {
    let out = run_vb("Console.WriteLine(Math.Ceiling(3.2))");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn namespace_math_abs() {
    let out = run_vb("Console.WriteLine(Math.Abs(-42))");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn namespace_math_sqrt() {
    let out = run_vb("Console.WriteLine(Math.Sqrt(144))");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn namespace_math_round() {
    let out = run_vb("Console.WriteLine(Math.Round(3.5))");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn namespace_abs_bare() {
    let out = run_vb("Console.WriteLine(Abs(-7))");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn namespace_new_point() {
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

#[test]
fn namespace_new_size() {
    let out = run_vb(
        r#"
Imports System.Drawing
Dim s As New Size(640, 480)
Console.WriteLine(s.width)
Console.WriteLine(s.height)
"#,
    );
    assert_eq!(out, vec!["640", "480"]);
}

#[test]
fn namespace_new_font() {
    let out = run_vb(
        r#"
Imports System.Drawing
Dim f As New Font("Arial", 12)
Console.WriteLine(f.name)
Console.WriteLine(f.size)
"#,
    );
    assert_eq!(out, vec!["Arial", "12"]);
}

#[test]
fn namespace_new_button() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Dim btn As New Button()
Console.WriteLine(btn.__control_type)
"#,
    );
    assert_eq!(out, vec!["Button"]);
}

#[test]
fn namespace_new_textbox() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Dim txt As New TextBox()
Console.WriteLine(txt.__control_type)
"#,
    );
    assert_eq!(out, vec!["TextBox"]);
}

#[test]
fn namespace_new_label() {
    let out = run_vb(
        r#"
Imports System.Windows.Forms
Dim lbl As New Label()
Console.WriteLine(lbl.__control_type)
"#,
    );
    assert_eq!(out, vec!["Label"]);
}

// ============================================================
// D. STRING HOST FUNCTIONS
// ============================================================

#[test]
fn host_string_len() {
    let out = run_vb(r#"Console.WriteLine(Len("hello"))"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn host_string_ucase() {
    let out = run_vb(r#"Console.WriteLine(UCase("hello"))"#);
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn host_string_lcase() {
    let out = run_vb(r#"Console.WriteLine(LCase("HELLO"))"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn host_string_trim() {
    let out = run_vb(r#"Console.WriteLine(Trim("  hi  "))"#);
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn host_string_mid() {
    let out = run_vb(r#"Console.WriteLine(Mid("hello", 2, 3))"#);
    assert_eq!(out, vec!["ell"]);
}

#[test]
fn host_string_left() {
    let out = run_vb(r#"Console.WriteLine(Left("hello", 3))"#);
    assert_eq!(out, vec!["hel"]);
}

#[test]
fn host_string_right() {
    let out = run_vb(r#"Console.WriteLine(Right("hello", 3))"#);
    assert_eq!(out, vec!["llo"]);
}

#[test]
fn host_string_instr() {
    let out = run_vb(r#"Console.WriteLine(InStr("hello world", "world"))"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn host_string_replace() {
    let out = run_vb(r#"Console.WriteLine(Replace("hello world", "world", "vb"))"#);
    assert_eq!(out, vec!["hello vb"]);
}

#[test]
fn host_string_split_join() {
    let out = run_vb(
        r#"
Dim parts() As String = Split("a,b,c", ",")
Console.WriteLine(Join(parts, "-"))
"#,
    );
    assert_eq!(out, vec!["a-b-c"]);
}

#[test]
fn host_string_chr_asc() {
    let out = run_vb(
        r#"
Console.WriteLine(Chr(65))
Console.WriteLine(Asc("A"))
"#,
    );
    assert_eq!(out, vec!["A", "65"]);
}

// ============================================================
// E. TYPE CONVERSIONS VIA HOST
// ============================================================

#[test]
fn host_val() {
    let out = run_vb(r#"Console.WriteLine(Val("42.5"))"#);
    assert_eq!(out, vec!["42.5"]);
}

#[test]
fn host_cstr() {
    let out = run_vb(r#"Console.WriteLine(CStr(123))"#);
    assert_eq!(out, vec!["123"]);
}

#[test]
fn host_cbool() {
    let out = run_vb(r#"Console.WriteLine(CBool(1))"#);
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["true"]));
}

// ============================================================
// F. INVOKE VB METHOD FROM RUST (event dispatch pattern)
// ============================================================

#[test]
fn invoke_vb_method_with_me() {
    let (mut vm, output) = run_vb_vm(
        r#"
Public Class Counter
    Dim count As Integer
    Public Sub New()
        count = 0
    End Sub
    Public Sub Inc()
        count = count + 1
    End Sub
    Public Sub Report()
        Console.WriteLine(count)
    End Sub
End Class
Dim c As New Counter()
"#,
    );
    let instance = vm.globals.get("c").cloned().unwrap();
    let inc = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("inc").cloned()
    } else {
        None
    }
    .unwrap();
    let report = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("report").cloned()
    } else {
        None
    }
    .unwrap();

    // Simulate 3 button clicks
    vm.invoke(&inc, &[instance.clone()]).unwrap();
    vm.invoke(&inc, &[instance.clone()]).unwrap();
    vm.invoke(&inc, &[instance.clone()]).unwrap();
    vm.invoke(&report, &[instance.clone()]).unwrap();

    assert_eq!(output.lock().unwrap().last().map(|s| s.as_str()), Some("3"));
}

#[test]
fn invoke_vb_method_reads_control_property() {
    // Simulates: handler reads Me.txtName.Text
    let (mut vm, output) = run_vb_vm(
        r#"
Public Class Form1
    Dim txtName As Object
    Public Sub New()
        txtName = New System.Windows.Forms.TextBox()
        txtName.text = "Alice"
    End Sub
    Public Sub ShowName()
        Console.WriteLine(txtName.text)
    End Sub
End Class
Dim f As New Form1()
"#,
    );
    let instance = vm.globals.get("f").cloned().unwrap();
    let method = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("showname").cloned()
    } else {
        None
    }
    .unwrap();

    vm.invoke(&method, &[instance.clone()]).unwrap();
    assert_eq!(
        output.lock().unwrap().last().map(|s| s.as_str()),
        Some("Alice")
    );
}

#[test]
fn invoke_vb_method_modifies_field() {
    let (mut vm, _) = run_vb_vm(
        r#"
Public Class State
    Dim value As String
    Public Sub New()
        value = "initial"
    End Sub
    Public Sub Update(v As String)
        value = v
    End Sub
End Class
Dim s As New State()
"#,
    );
    let instance = vm.globals.get("s").cloned().unwrap();
    let update = if let Value::Object(obj) = &instance {
        obj.lock().unwrap().properties.get("update").cloned()
    } else {
        None
    }
    .unwrap();

    vm.invoke(
        &update,
        &[instance.clone(), Value::String(Arc::from("changed"))],
    )
    .unwrap();

    // Check the field was modified
    if let Value::Object(obj) = &instance {
        let v = obj.lock().unwrap().properties.get("value").cloned();
        assert_eq!(v.map(|v| format!("{}", v)).as_deref(), Some("changed"));
    }
}
