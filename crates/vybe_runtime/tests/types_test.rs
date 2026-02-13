use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::parse_program;

#[test]
fn test_types() {
    let source = std::fs::read_to_string("../../tests/test_types.vb").expect("Failed to read test file");
    let program = parse_program(&source).expect("Parse error");
    
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");

    // Call Main manually
    use vybe_parser::ast::Identifier;
    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

    let output_str = interp.side_effects.iter().map(|e| {
         if let RuntimeSideEffect::ConsoleOutput(msg) = e {
             msg.clone()
         } else {
             "".to_string()
         }
    }).collect::<Vec<_>>().join("\n");
    
    println!("Output:\n{}", output_str);

    assert!(output_str.contains("CByte(255) = 255"));
    assert!(output_str.contains("TypeName(b) = Byte"));
    assert!(output_str.contains("CChar('A') = A"));
    assert!(output_str.contains("TypeName(c) = Char"));
    assert!(output_str.contains("CChar(65) = A"));
    assert!(output_str.contains("CInt('&HFF') = 255"));
    assert!(output_str.contains("CLng('&H100') = 256"));
    assert!(output_str.contains("CInt('&O10') = 8"));
    
    assert!(output_str.contains("Type Test Completed"));
}
