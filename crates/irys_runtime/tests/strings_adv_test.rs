use irys_runtime::{Interpreter, RuntimeSideEffect};
use irys_parser::parse_program;

#[test]
fn test_strings_adv() {
    let source = std::fs::read_to_string("../../tests/test_strings_adv.vb").expect("Failed to read test file");
    let program = parse_program(&source).expect("Parse error");
    
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");

    // Call Main manually
    use irys_parser::ast::Identifier;
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

    assert!(output_str.contains("Replace: Hello Universe"));
    assert!(output_str.contains("Split(0): A"));
    assert!(output_str.contains("Split(1): B"));
    assert!(output_str.contains("Join: A-B-C"));
    assert!(output_str.contains("StrReverse: CBA"));
    assert!(output_str.contains("InStrRev: 10"));
    assert!(output_str.contains("Space: '   '"));
    assert!(output_str.contains("String: ***"));
    assert!(output_str.contains("Asc: 65"));
    assert!(output_str.contains("Chr: B"));
    
    assert!(output_str.contains("String Test Completed"));
}
