use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

#[test]
fn test_stringbuilder() {
    let source = std::fs::read_to_string("../../tests/test_stringbuilder.vb")
        .expect("Failed to read test file");
    let program = parse_program(&source).expect("Parse error");

    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");

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

    // Markers
    assert!(output_str.contains("=== StringBuilder Test Start ==="));
    assert!(output_str.contains("=== StringBuilder Test End ==="));

    // Basic Append
    assert!(output_str.contains("Append: Hello World"));

    // Constructors
    assert!(output_str.contains("InitCtor: Init"));
    assert!(output_str.contains("CapCtor: cap"));

    // AppendLine
    assert!(output_str.contains("AppendLine Contains Line1: True"));
    assert!(output_str.contains("AppendLine Contains Line2: True"));

    // AppendFormat
    assert!(output_str.contains("AppendFormat: Name: Alice, Age: 30"));

    // Insert
    assert!(output_str.contains("Insert: Hello World"));

    // Remove
    assert!(output_str.contains("Remove: Hello"));

    // Replace
    assert!(output_str.contains("Replace: Hello VB"));

    // Clear
    assert!(output_str.contains("Clear Length: 0"));
    assert!(output_str.contains("Clear ToString: ''"));

    // Length read
    assert!(output_str.contains("Length: 5"));

    // Length set (truncate)
    assert!(output_str.contains("Truncate: Hello"));

    // Length set (pad)
    assert!(output_str.contains("PadLength: 5"));

    // Method chaining
    assert!(output_str.contains("Chain: ABC"));

    // EnsureCapacity
    assert!(output_str.contains("EnsureCapacity >= 100: True"));

    // Chars indexer
    assert!(output_str.contains("Chars(2): C"));

    // ToString
    assert!(output_str.contains("ToString: Final"));

    // Equals
    assert!(output_str.contains("Equals Same: True"));
    assert!(output_str.contains("Equals Diff: False"));

    // CopyTo
    assert!(output_str.contains("CopyTo(0): H"));
    assert!(output_str.contains("CopyTo(4): o"));

    // CSV chaining
    assert!(output_str.contains("CSV: Name,Age,City"));

    // Replace chained
    assert!(output_str.contains("ReplaceChain: a-b-c"));

    // Full qualified name
    assert!(output_str.contains("FullQualified: FullQualified"));
}
