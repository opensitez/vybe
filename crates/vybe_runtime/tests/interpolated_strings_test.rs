use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_interpolated_strings() {
    let source = std::fs::read_to_string("../../tests/test_interpolated_strings.vb")
        .expect("Failed to read test file");
    let program = parse_program(&source).expect("Failed to parse program");

    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");

    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

    let output: Vec<String> = interp.side_effects.iter().filter_map(|e| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = e {
            Some(msg.clone())
        } else {
            None
        }
    }).collect();

    for line in &output {
        println!("{}", line);
    }

    let has_failure = output.iter().any(|msg| msg.contains("FAIL"));
    assert!(!has_failure, "Interpolated string test had failures");

    let has_success = output.iter().any(|msg| msg.contains("SUCCESS"));
    assert!(has_success, "Interpolated string test did not complete");
}
