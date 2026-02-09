use irys_runtime::{Interpreter, Value};
use irys_parser::ast::Identifier;
use irys_parser::parse_program;
use std::rc::Rc;
use std::cell::RefCell;

#[test]
fn test_byref_behavior() {
    let source = std::fs::read_to_string("../../tests/repro_byref.vb").expect("Failed to read repro_byref.vb");
    let program = parse_program(&source).expect("Failed to parse");
    
    let mut interp = Interpreter::new();
    
    // Capture stdout
    let output = Rc::new(RefCell::new(Vec::<String>::new()));
    println!("DEBUG: AST: {:#?}", program);
    let output_clone = output.clone();
    
    // Mock Console.WriteLine to capture output
    // We can't easily mock "Console.WriteLine" inside the interpreter without modifying it or using the side-effect queue.
    // But `repro_byref.vb` uses Console.WriteLine.
    // The current interpreter implementation pushes to `side_effects`.
    
    interp.run(&program).expect("Runtime error");
    
    // Explicitly call Main since it's in a module
    // The parser/interpreter might not auto-run Main for module-based code unless we tell it to
    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");

    // Check side effects
    let mut success = false;
    while let Some(effect) = interp.side_effects.pop_front() {
        if let irys_runtime::RuntimeSideEffect::ConsoleOutput(msg) = effect {
            println!("Output: {}", msg.trim());
            if msg.contains("SUCCESS: ByRef works") {
                success = true;
            }
        }
    }
    
    if !success {
        panic!("ByRef test failed - did not see SUCCESS message");
    }
}
