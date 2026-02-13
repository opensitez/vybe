use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_collections_functionality() {
    let source = std::fs::read_to_string("../../tests/test_collections.vb").expect("Failed to read test file");
    let program = parse_program(&source).expect("Failed to parse program");

    let mut interp = Interpreter::new();
    
    interp.run(&program).expect("Runtime error");
    
    // Explicitly call Main
    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

    let has_success = interp.side_effects.iter().any(|effect| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
            println!("Output: {}", msg);
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
    
    assert!(has_success, "Collections test failed - did not see SUCCESS");
    assert!(!has_failure, "Collections test failed - saw FAILURE");
}
