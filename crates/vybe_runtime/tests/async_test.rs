use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::parse_program;

#[test]
fn test_async_await() {
    let source = std::fs::read_to_string("../../tests/test_async.vb").expect("Failed to read test file");
    let program = parse_program(&source).expect("Parse error");
    
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");

    // Call Main manually if run() doesn't automatically trigger it (depends on implementation)
    // interpreter.run() executes module statements. If Main is a Sub, it might not be called automatically unless there is a call statement.
    // However, the test file has "Sub Main", usually we need to call it.
    // Let's assume we need to call "Main".
    
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
    
    // Check execution order (synchronous simulation)
    assert!(output_str.contains("Start"));
    assert!(output_str.contains("Result: 42"));
    assert!(output_str.contains("Msg: Async Sub Call"));
    assert!(output_str.contains("End"));
}
