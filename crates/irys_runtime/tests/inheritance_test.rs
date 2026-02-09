use irys_runtime::interpreter::Interpreter;
use irys_parser::parse_program;
use std::fs;
use std::path::PathBuf;

#[test]
fn test_inheritance_and_partials() {
    let mut d = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    // Go up from crates/irys_runtime to root, then to tests
    d.push("../../tests/test_inheritance.vb");
    
    let code = fs::read_to_string(&d).expect("Failed to read test file");
    
    // Append call to Main
    let full_code = format!("{}\nCall Main()", code);
    
    let program = parse_program(&full_code).expect("Failed to parse program");
    
    let mut interpreter = Interpreter::new();
    interpreter.run(&program).expect("Failed to run program");
}
