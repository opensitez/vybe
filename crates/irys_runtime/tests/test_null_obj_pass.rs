use irys_runtime::{Interpreter, Value};
use irys_parser::parse_program;

#[test]
fn test_null_object_passing() {
    let form_code = r#"
    Public Class Form1
        Public btn0 As Object 
        Public result As String
        
        Public Sub InitAndCall()
             ' btn0 is Nothing
             ' But we simulate RuntimePanel settings:
             ' RuntimeGlobals.Form1Instance.btn0.Caption = "Button 0"
             
             HandleClick(btn0)
        End Sub
        
        Public Sub HandleClick(btn As Object)
            If btn Is Nothing Then
                result = "Nothing"
                
                ' Even if Nothing, try property access via fallback?
                ' interpreter fallback uses expr_to_string.
                ' expr_to_string(btn) -> "btn".
                ' So it looks for "btn.Caption" in Env.
                
                ' But "btn0.Caption" is in Env. 
                ' So access should FAIL/Return Empty unless we somehow pass the name "btn0".
                
                Dim cap
                cap = btn.Caption
                if cap = "" Then
                    result = result & " + EmptyCaption"
                Else
                    result = result & " + " & cap
                End If
            Else
                result = "NotNothing"
            End If
        End Sub
    End Class
    "#;

    let mut interp = Interpreter::new();
    let prog = parse_program(form_code).expect("Failed to parse form");
    interp.load_module("Form1", &prog).expect("Failed to load module");
    
    let global_setup = r#"
    Module RuntimeGlobals
        Public Form1Instance As New Form1
    End Module
    "#;
    let setup_prog = parse_program(global_setup).expect("Failed to parse setup");
    interp.load_module("RuntimeGlobals", &setup_prog).expect("Failed to load setup");
    
    // Simulate RuntimePanel sync
    interp.env.define("RuntimeGlobals.Form1Instance.btn0.Caption", Value::String("Button 0".to_string()));
    
    // Call InitAndCall
    interp.call_instance_method("RuntimeGlobals.Form1Instance", "InitAndCall", &[]).expect("Failed to call");
    
    // Check result
    let check_expr = irys_parser::parse_expression_str("RuntimeGlobals.Form1Instance.result").unwrap();
    let val = interp.evaluate_expr(&check_expr).expect("Failed evaluation");
    
    match val {
        Value::String(s) => {
            println!("Result: {}", s);
            // If passed ByVal/ByRef Nothing, it is Nothing inside.
            // Accessing .Caption on Nothing uses fallback "btn.Caption", which is undefined.
            // So result should be "Nothing + EmptyCaption"
            assert_eq!(s, "Nothing + EmptyCaption", "Expected nothing passing to result in default behavior");
        },
        _ => panic!("Expected string"),
    }
}
