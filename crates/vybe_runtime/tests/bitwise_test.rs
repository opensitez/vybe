use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::parse_program;

#[test]
fn test_bitwise_ops() {
    let source = std::fs::read_to_string("../../tests/test_bitwise.vb").expect("Failed to read test file");
    let program = parse_program(&source).expect("Parse error");
    
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");

    // Call Main manually if not called automatically by run() (it performs module-level execution)
    // test_bitwise.vb has Sub Main.
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

    assert!(output_str.contains("True And True = True"));
    assert!(output_str.contains("True Or False = True"));
    assert!(output_str.contains("True Xor False = True"));
    assert!(output_str.contains("False Xor False = False"));
    
    assert!(output_str.contains("5 And 3 = 1"));
    assert!(output_str.contains("5 Or 3 = 7"));
    assert!(output_str.contains("5 Xor 3 = 6"));
    assert!(output_str.contains("Not 5 = -6"));
    
    assert!(output_str.contains("1 << 1 = 2"));
    assert!(output_str.contains("1 << 2 = 4"));
    assert!(output_str.contains("8 >> 1 = 4"));
    assert!(output_str.contains("-8 >> 1 = -4"));
    
    assert!(output_str.contains("Precedence: Or > Xor (Confirmed)"));
    assert!(output_str.contains("Precedence: And > Or (Confirmed)"));
    
    assert!(output_str.contains("Bitwise Test Completed"));
}
