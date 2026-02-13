use vybe_runtime::{Interpreter, Value, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;
use std::rc::Rc;
use std::cell::RefCell;
// use std::collections::VecDeque;

#[test]
fn test_date_functionality() {
    let source = std::fs::read_to_string("../../tests/test_date.vb").expect("Failed to read test file");
    let program = parse_program(&source).expect("Failed to parse program");

    let mut interp = Interpreter::new();
    
    // Capture stdout
    // Output is captured in side_effects automatically by Console.WriteLine
    
    interp.run(&program).expect("Runtime error");
    
    // Explicitly call Main
    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

    let has_success = interp.side_effects.iter().any(|effect| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
            println!("Console Output: {}", msg);
            msg.contains("SUCCESS")
        } else {
            false
        }
    });

    let has_failure = interp.side_effects.iter().any(|effect| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
            msg.contains("FAILURE")
        } else {
            false
        }
    });
    
    assert!(has_success, "Date test failed - did not see SUCCESS");
    assert!(!has_failure, "Date test failed - saw FAILURE");
}
