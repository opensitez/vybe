use super::helpers::run_vb;
use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value};
use vybe_host::{SideEffectQueue, SideEffect};

fn compile_vb_gui(source: &str) -> (Vec<String>, Vec<SideEffect>) {
    let program = vybe_parser_basic::parse_program(source)
        .unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    let queue = Rc::new(RefCell::new(SideEffectQueue::new()));

    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    let chunks = vybe_compiler_vb::Compiler::new().compile(&program)
        .unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));

    let effects: Vec<SideEffect> = queue.borrow_mut().drain();
    (output.borrow().clone(), effects)
}

#[test]
fn simple_form_with_button() {
    let (output, effects) = compile_vb_gui(r#"
Module Program
    Sub Main()
        Dim form As Object = Window.Forms.Form("My App")
        Dim btn As Object = Window.Forms.Button()
        btn.text = "Click Me"
        btn.left = 10
        btn.top = 20
        Console.WriteLine("Form created")
    End Sub
End Module
"#);
    assert_eq!(output, vec!["Form created"]);

    // Check that a form was created (PropertyChange for Text)
    let has_form = effects.iter().any(|e| matches!(e, SideEffect::PropertyChange { object, property, .. }
        if object == "My App" && property == "Text"));
    assert!(has_form, "Expected form creation side effect");
}

#[test]
fn button_properties() {
    let (_, effects) = compile_vb_gui(r#"
Module Program
    Sub Main()
        Dim btn As Object = Window.Forms.Button()
        btn.text = "Hello"
        btn.width = 200
        btn.height = 50
    End Sub
End Module
"#);
    // Button was created (it's an object) — no side effects until Controls.Add
    // But properties are stored on the object
    // This test just verifies it compiles and runs without error
    assert!(true);
}

#[test]
fn form_with_controls_add() {
    let (_, effects) = compile_vb_gui(r#"
Module Program
    Sub Main()
        Dim form As Object = Window.Forms.Form("Calculator")
        Dim btn As Object = Window.Forms.Button()
        btn.text = "Press"
        btn.left = 10
        btn.top = 10
        btn.width = 100
        btn.height = 40

        ' Add control to form — this should emit AddControl side effect
        vybe.gui.controlsAdd("Calculator", btn)
    End Sub
End Module
"#);

    // Check that AddControl was emitted
    let has_add = effects.iter().any(|e| matches!(e, SideEffect::AddControl { control_type, .. }
        if control_type == "Button"));
    assert!(has_add, "Expected AddControl side effect for Button, got: {:?}", effects);
}

#[test]
fn calculator_example() {
    let source = std::fs::read_to_string(
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../examples/vb/calculator.vb")
    ).unwrap();
    let (_, effects) = compile_vb_gui(&source);

    // 1 textbox + 16 buttons = 17 controls
    let add_count = effects.iter().filter(|e| matches!(e, SideEffect::AddControl { .. })).count();
    assert_eq!(add_count, 17, "Expected 17 AddControl effects, got {}", add_count);

    // Should have a RunApplication effect
    let has_run = effects.iter().any(|e| matches!(e, SideEffect::RunApplication { .. }));
    assert!(has_run, "Expected RunApplication side effect");
}

#[test]
fn contacts_example() {
    let source = std::fs::read_to_string(
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../examples/vb/form_contacts.vb")
    ).unwrap();
    let (_, effects) = compile_vb_gui(&source);

    // 4 labels + 3 textboxes + 2 buttons + 1 grid + 1 status label = 11 controls
    let add_count = effects.iter().filter(|e| matches!(e, SideEffect::AddControl { .. })).count();
    assert_eq!(add_count, 11, "Expected 11 AddControl effects, got {}", add_count);
}

#[test]
fn multiple_controls() {
    let (output, effects) = compile_vb_gui(r#"
Module Program
    Sub Main()
        Dim form As Object = Window.Forms.Form("Test Form")

        Dim lbl As Object = Window.Forms.Label()
        lbl.text = "Name:"
        lbl.left = 10
        lbl.top = 10
        vybe.gui.controlsAdd("Test Form", lbl)

        Dim txt As Object = Window.Forms.TextBox()
        txt.left = 80
        txt.top = 10
        txt.width = 200
        vybe.gui.controlsAdd("Test Form", txt)

        Dim btn As Object = Window.Forms.Button()
        btn.text = "Submit"
        btn.left = 80
        btn.top = 50
        vybe.gui.controlsAdd("Test Form", btn)

        Console.WriteLine("3 controls added")
    End Sub
End Module
"#);

    assert_eq!(output, vec!["3 controls added"]);

    let add_count = effects.iter().filter(|e| matches!(e, SideEffect::AddControl { .. })).count();
    assert_eq!(add_count, 3, "Expected 3 AddControl effects, got {}", add_count);
}
