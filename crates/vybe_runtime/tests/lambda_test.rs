use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_lambda_functionality() {
    let source = std::fs::read_to_string("../../tests/test_lambda.vb").expect("Failed to read test file");
    match parse_program(&source) {
        Ok(program) => {
            let mut interp = Interpreter::new();
            
            interp.run(&program).expect("Runtime error");
            
            // Explicitly call Main
            let main_ident = Identifier::new("Main");
            interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

            let has_success_square = interp.side_effects.iter().any(|effect| {
                if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
                    msg.contains("SUCCESS: Square lambda")
                } else {
                    false
                }
            });
            
            let has_success_closure = interp.side_effects.iter().any(|effect| {
                if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
                    msg.contains("SUCCESS: Closure capture")
                } else {
                    false
                }
            });

            assert!(has_success_square, "Lambda function test failed");
            assert!(has_success_closure, "Lambda closure test failed");
        },
        Err(e) => {
            panic!("Parse error: {:?}", e);
        }
    }
}
