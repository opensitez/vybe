use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_new_functions_functionality() {
    let source = std::fs::read_to_string("../../tests/test_new_functions.vb").expect("Failed to read test file");
    match parse_program(&source) {
        Ok(program) => {
            let mut interp = Interpreter::new();
            interp.run(&program).expect("Runtime error");
            
            let main_ident = Identifier::new("Main");
            interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

            for effect in &interp.side_effects {
                if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
                    println!("VB Output: {}", msg);
                }
            }

            let has_success = interp.side_effects.iter().any(|effect| {
                if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
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
            
            assert!(has_success, "New functions test failed - did not see SUCCESS");
            assert!(!has_failure, "New functions test failed - saw FAILURE");
        },
        Err(e) => {
            panic!("Parse error: {:?}", e);
        }
    }
}
