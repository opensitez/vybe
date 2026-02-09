use irys_runtime::{Interpreter, Value};
use irys_parser::parse_program;

#[test]
fn test_runtime_instance_dispatch() {
    let form_code = r#"
    Public Class TestForm
        Public ClickCount As Integer
        
        Public Sub btn_Click()
            ClickCount = ClickCount + 1
        End Sub
    End Class
    "#;

    let mut interp = Interpreter::new();
    
    // 1. Load Form Module
    let prog = parse_program(form_code).expect("Failed to parse form");
    interp.load_module("TestForm", &prog).expect("Failed to load form");

    // 2. Create Global Instance (simulating RuntimePanel logic)
    let global_setup = r#"
    Module RuntimeGlobals
        Public TestFormInstance As New TestForm
    End Module
    "#;
    let setup_prog = parse_program(global_setup).expect("Failed to parse setup");
    interp.load_module("RuntimeGlobals", &setup_prog).expect("Failed to load setup");

    // 3. Dispatch Event
    // NOTE: RuntimeGlobals is flattened, so TestFormInstance is global.
    // The previous RuntimePanel implementation used "Call RuntimeGlobals.TestFormInstance..." which is correct.
    // We test the corrected version: "Call RuntimeGlobals.TestFormInstance..."
    let dispatch_code = "Call RuntimeGlobals.TestFormInstance.btn_Click()";
    let dispatch_prog = parse_program(dispatch_code).expect("Failed to parse dispatch");
    
    // Run dispatch
    interp.load_module("EventRunner_1", &dispatch_prog).expect("Failed to run dispatch 1");
    
    // Check Result directly in env
    // Try both cases just to be safe, but declaration usually preserves case in env
    // Check Result directly in env (keys are module.var)
    let val = interp.env.get("runtimeglobals.testforminstance").expect("Instance not found");
    
    if let Value::Object(obj) = val {
        let fields = &obj.borrow().fields;
        // Object fields are lowercased by collect_fields
        let count = fields.get("clickcount").expect("ClickCount not found");
        assert_eq!(count, &Value::Integer(1));
    } else {
        panic!("TestFormInstance is not an object");
    }

    // Run dispatch again
    interp.load_module("EventRunner_2", &dispatch_prog).expect("Failed to run dispatch 2");

    let val = interp.env.get("runtimeglobals.testforminstance").expect("Instance not found");
    if let Value::Object(obj) = val {
        let fields = &obj.borrow().fields;
        let count = fields.get("clickcount").expect("ClickCount not found");
        assert_eq!(count, &Value::Integer(2));
    }
}
