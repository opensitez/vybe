use irys_runtime::{Interpreter, RuntimeSideEffect};
use irys_parser::ast::Identifier;
use irys_parser::parse_program;

#[test]
fn test_select_functionality() {
    let source = std::fs::read_to_string("../../tests/test_select_case.vb").expect("Failed to read test file");
    match parse_program(&source) {
        Ok(program) => {
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
            
            assert!(has_success, "Select Case test failed - did not see SUCCESS");
            assert!(!has_failure, "Select Case test failed - saw FAILURE");
        },
        Err(e) => {
            panic!("Parse error: {:?}", e);
        }
    }
}
