use vybe_runtime::{Interpreter, Value};
use vybe_parser::parse_program;

#[test]
fn test_object_event_handling() {
    let code = r#"
    Public Class MyForm
        Public Value As Integer
        
        Public Sub Form_Load()
            Value = 100
        End Sub
        
        Public Sub btn_Click()
            Value = Value + 1
        End Sub
    End Class
    "#;

    let mut interp = Interpreter::new();
    let prog = parse_program(code).expect("Failed to parse");
    interp.load_module("MyForm", &prog).expect("Failed to load");

    // confirm class exists (module-qualified key: module.class)
    assert!(interp.classes.contains_key("myform.myform"));

    // 1. Instantiate
    // We can't use "Dim" inside eval_expr. We need to run a statement.
    // Or we can manually construct the object.
    
    // Let's try running a setup script
    let setup = "Dim f As New MyForm";
    let setup_prog = parse_program(setup).expect("Failed to parse setup");
    // Interpreter::run executes statements in the current scope (global)
    // But Interpreter doesn't expose run() directly for public use easily? 
    // It has `run_script`? No, it has `execute`.
    
    // Let's verify what public API we have for execution.
    // We typically use `load_module` to load code, and `call_event_handler` to run it.
    // But we can also use `eval` for expressions.
    
    // If we wrap the instantiation in a Module Main?
    // Module Main
    //   Public f As New MyForm
    // End Module
    
    // But for the RuntimePanel, we want to simulate what we can do there.
    // RuntimePanel uses `interp.load_module`.
    
    // Let's try creating a "RuntimeWrapper" module
    let wrapper = r#"
    Module RuntimeWrapper
        Public App As New MyForm
    End Module
    "#;
    let wrapper_prog = parse_program(wrapper).expect("Wrapper parse failed");
    interp.load_module("RuntimeWrapper", &wrapper_prog).expect("Wrapper load failed");
    
    // Now "RuntimeWrapper.App" should exist?
    // Global variable initialization happens... when?
    // In `vybe_runtime`, module fields are initialized when?
    // They might be lazy or need `Main`.
    
    // Let's force access to initialize?
    // Or just manually run `App = New MyForm`?
    
    // Check if we can call "App.Form_Load"
    // `call_event_handler` expects a Sub name.
    // It calls `call_user_sub`.
    
    // If we ask for "RuntimeWrapper.App.Form_Load"?
    // The interpreter split logic might not handle 3 parts.
    
    // Let's try executing a proxy sub
    let proxy = r#"
    Module Wrapper
        Public Instance As New MyForm
        
        Sub TriggerLoad()
            Instance.Form_Load()
        End Sub
        
        Sub TriggerClick()
            Instance.btn_Click()
        End Sub
    End Module
    "#;
    
    let proxy_prog = parse_program(proxy).expect("Proxy parse failed");
    interp.load_module("Wrapper", &proxy_prog).unwrap();
    
    // Call TriggerLoad
    interp.call_event_handler("TriggerLoad", &[]).expect("TriggerLoad failed");
    
    // Check Value
    let val = eval_str(&mut interp, "Wrapper.Instance.Value").unwrap();
    assert_eq!(val, Value::Integer(100));
    
    // Call TriggerClick
    interp.call_event_handler("TriggerClick", &[]).expect("TriggerClick failed");
    
    let val = eval_str(&mut interp, "Wrapper.Instance.Value").unwrap();
    assert_eq!(val, Value::Integer(101));
}

fn eval_str(interp: &mut Interpreter, expr_str: &str) -> Result<Value, vybe_runtime::RuntimeError> {
    let expr = vybe_parser::parse_expression_str(expr_str)
        .map_err(|e| vybe_runtime::RuntimeError::Custom(format!("Parse error: {}", e)))?;
    interp.evaluate_expr(&expr)
}
