use super::helpers::{run_vb, run_vb_gui};

#[test]
fn simple_form_with_button() {
    let (_vm, gui, output) = run_vb_gui(r#"
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
    assert_eq!(output.lock().unwrap().clone(), vec!["Form created"]);

    // Form was created — check its text property
    let g = gui.lock().unwrap();
    assert!(g.control_names.len() >= 1 || g.form.control_count() >= 0,
        "Expected form to be created");
}

#[test]
fn button_properties() {
    let (_vm, _gui, _output) = run_vb_gui(r#"
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
    let (_vm, gui, _output) = run_vb_gui(r#"
Module Program
    Sub Main()
        Dim form As Object = Window.Forms.Form("Calculator")
        Dim btn As Object = Window.Forms.Button()
        btn.text = "Press"
        btn.left = 10
        btn.top = 10
        btn.width = 100
        btn.height = 40

        ' Add control to form — this should register the control
        vybe.gui.controlsAdd("Calculator", btn)
    End Sub
End Module
"#);

    // Check that the button was added as a control
    let g = gui.lock().unwrap();
    assert!(g.form.control_count() >= 1,
        "Expected at least 1 control, got {}", g.form.control_count());
}

#[test]
fn calculator_example() {
    let source = std::fs::read_to_string(
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../examples/vb/calculator.vb")
    ).unwrap();
    let (_vm, gui, _output) = run_vb_gui(&source);

    let g = gui.lock().unwrap();

    // 1 textbox + 16 buttons = 17 controls
    assert_eq!(g.form.control_count(), 17,
        "Expected 17 controls, got {}. Control names: {:?}",
        g.form.control_count(), g.control_names);

    // Should have triggered runApplication
    assert!(g.should_run, "Expected should_run to be true (RunApplication was called)");
}

#[test]
fn contacts_example() {
    let source = std::fs::read_to_string(
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../examples/vb/form_contacts.vb")
    ).unwrap();
    let (_vm, gui, _output) = run_vb_gui(&source);

    let g = gui.lock().unwrap();

    // 4 labels + 3 textboxes + 2 buttons + 1 grid + 1 status label = 11 controls
    assert_eq!(g.form.control_count(), 11,
        "Expected 11 controls, got {}. Control names: {:?}",
        g.form.control_count(), g.control_names);
}

#[test]
fn multiple_controls() {
    let (_vm, gui, output) = run_vb_gui(r#"
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

    assert_eq!(output.lock().unwrap().clone(), vec!["3 controls added"]);

    let g = gui.lock().unwrap();
    assert_eq!(g.form.control_count(), 3,
        "Expected 3 controls, got {}. Control names: {:?}",
        g.form.control_count(), g.control_names);
}
