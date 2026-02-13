
#[cfg(test)]
mod tests {
    use vybe_runtime::{Interpreter, Value};
    use vybe_parser::parse_program;

    #[test]
    fn test_legacy_control_access() {
        // Simulate what RuntimePanel does: 
        // 1. Load form code (as a module/global script)
        // 2. Register control names as global variables (strings)
        // 3. Execute event handler that uses the control

        let code = r#"
        Sub btn1_Click()
            txtCalc.Text = "Hello"
        End Sub
        "#;

        let mut interp = Interpreter::new();
        let prog = parse_program(code).expect("Failed to parse");
        interp.load_module("Form1", &prog).expect("Failed to load");

        // THIS IS THE CRITICAL PART: Registering the control as a string variable
        // This logic was previously inside `if is_vbnet` block in RuntimePanel
        interp.env.define("txtCalc", Value::String("txtCalc".to_string()));
        interp.env.define("btn1", Value::String("btn1".to_string()));

        // Call the event handler
        // The handler uses `txtCalc.Text = ...`
        // `txtCalc` resolves to String("txtCalc")
        // `MemberAssignment` handles String("txtCalc").Text = ... by setting env var "txtCalc.Text"

        interp.call_event_handler("btn1_Click", &[]).expect("Failed to call handler");

        // Verify result
        // The side effect of setting a property on a string proxy is that it sets a variable in the environment
        let val = interp.env.get("txtCalc.Text").expect("Property not set");
        assert_eq!(val, Value::String("Hello".to_string()));
    }
}
