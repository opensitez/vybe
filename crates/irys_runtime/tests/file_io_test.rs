use irys_runtime::Interpreter;
use irys_parser::parser::parse_program;
use std::fs;
use std::path::PathBuf;

#[test]
fn test_file_io_execution() {
    let manifest_dir = std::env::var("CARGO_MANIFEST_DIR").unwrap();
    let root_dir = PathBuf::from(manifest_dir).parent().unwrap().parent().unwrap().to_path_buf();
    let vb_path = root_dir.join("tests/test_file_io.vb");
    
    let source = fs::read_to_string(&vb_path).expect("Failed to read test_file_io.vb");
    
    let program = parse_program(&source).expect("Failed to parse VB script");

    let mut interpreter = Interpreter::new();
    let temp_dir = std::env::temp_dir();
    let temp_file = temp_dir.join("test_io_net.txt");
    let temp_legacy = temp_dir.join("test_io_legacy.txt");
    
    interpreter.env.define_const("TestFilePath", irys_runtime::Value::String(temp_file.to_string_lossy().to_string()));
    interpreter.env.define_const("TestLegacyPath", irys_runtime::Value::String(temp_legacy.to_string_lossy().to_string()));
    
    // Capture console output
    let mut outputs = Vec::new();

    // Run the program
    let result = interpreter.run(&program);
    
    if let Err(e) = &result {
        println!("Runtime error: {:?}", e);
    }
    assert!(result.is_ok(), "Runtime failed");

    // Check side effects for Console output
    while let Some(effect) = interpreter.side_effects.pop_front() {
        if let irys_runtime::RuntimeSideEffect::MsgBox(msg) = effect {
            outputs.push(msg);
        }
    }

    // Verify outputs
    let output_str = outputs.join("\n");
    println!("Interpreter Output:\n{}", output_str);

    assert!(output_str.contains("NetRead: Hello from .NET I/O"));
    assert!(output_str.contains("NetExists: True"));
    assert!(output_str.contains("PathExt: txt")); // GetExtension returns "txt" or ".txt"? Rust returns "txt"
    assert!(output_str.contains("LegacyRead1: Line 1"));
    assert!(output_str.contains("LegacyRead2: Line 2"));
}
