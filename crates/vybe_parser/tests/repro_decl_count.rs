use vybe_parser::parse_program;
use vybe_parser::ast::Declaration;

#[test]
fn test_program_declarations_count() {
    let code = "Sub Foo()\nEnd Sub";
    let program = parse_program(code).expect("Failed to parse");
    
    println!("Declarations: {}", program.declarations.len());
    println!("Statements: {}", program.statements.len());
    
    // Check if it's in declarations
    assert_eq!(program.declarations.len(), 1, "Expected 1 declaration, found {}", program.declarations.len());
    
    match &program.declarations[0] {
        Declaration::Sub(sub) => assert_eq!(sub.name.as_str(), "Foo"),
        _ => panic!("Expected SubDecl"),
    }
}
