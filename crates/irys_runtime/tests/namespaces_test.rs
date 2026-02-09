use irys_runtime::Interpreter;
use irys_parser::parser::parse_program;
use std::fs;
use std::path::PathBuf;

#[test]
fn test_namespaces_execution() {
    let manifest_dir = std::env::var("CARGO_MANIFEST_DIR").unwrap();
    let root_dir = PathBuf::from(manifest_dir).parent().unwrap().parent().unwrap().to_path_buf();
    let vb_path = root_dir.join("tests/test_namespaces.vb");
    
    let source = fs::read_to_string(&vb_path).expect("Failed to read test_namespaces.vb");
    
    let program = parse_program(&source).expect("Failed to parse VB script");

    let mut interpreter = Interpreter::new();
    
    // Capture console output
    let mut outputs = Vec::new();

    // Run the program
    let result = interpreter.run(&program);
    
    if let Err(e) = &result {
        println!("Runtime error: {:?}", e);
    }
    assert!(result.is_ok(), "Runtime failed");

    // Check side effects
    while let Some(effect) = interpreter.side_effects.pop_front() {
        if let irys_runtime::RuntimeSideEffect::MsgBox(msg) = effect {
            outputs.push(msg);
        }
    }

    // Verify outputs
    let output_str = outputs.join("\n");
    println!("Interpreter Output:\n{}", output_str);

    assert!(output_str.contains("[Console] Hello from System.Console"));
    assert!(output_str.contains("[Console] Max(10, 20) = 20"));
    assert!(output_str.contains("[Console] Sqrt(16) = 4"));
    assert!(output_str.contains("[Console] myMath.Min(10, 20) = 10"));
    assert!(output_str.contains("[Console] Hello from myConsole"));
}
