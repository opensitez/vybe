use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_foreach_iteration() {
    let source = std::fs::read_to_string("../../tests/test_foreach_iteration.vb")
        .expect("Failed to read test file");
    let program = parse_program(&source).expect("Failed to parse program");

    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error during declarations");

    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

    // Print all output for debugging
    for effect in &interp.side_effects {
        if let RuntimeSideEffect::ConsoleOutput(msg) = effect {
            print!("{}", msg);
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

    assert!(has_success, "ForEach iteration test did not produce SUCCESS");
    assert!(!has_failure, "ForEach iteration test produced FAILURE");
}
