use std::path::PathBuf;

#[test]
fn test_misc_gaps() {
    let test_file = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent().unwrap().parent().unwrap()
        .join("tests/test_misc_gaps.vb");
    
    let source = std::fs::read_to_string(&test_file)
        .expect("Failed to read test_misc_gaps.vb");
    
    let parsed = vybe_parser::parse_program(&source)
        .expect("Failed to parse test_misc_gaps.vb");
    
    let mut interp = vybe_runtime::Interpreter::new();
    
    interp.run(&parsed).expect("Runtime error in test_misc_gaps.vb");
    
    let output = interp.side_effects.iter()
        .filter_map(|se| match se {
            vybe_runtime::RuntimeSideEffect::ConsoleOutput(s) => Some(s.as_str()),
            _ => None,
        })
        .collect::<Vec<_>>()
        .join("");
    println!("{}", output);
    
    assert!(output.contains("SUCCESS"), "Misc gaps test did not pass:\n{}", output);
    assert!(!output.contains("FAIL:"), "Misc gaps test had failures:\n{}", output);
}
