#[test]
fn test_parse_with_imports() {
    let code = r#"Imports System
Imports System.Net

Module Program
    Sub Main()
        Console.WriteLine("hello")
    End Sub
End Module
"#;
    let prog = vybe_parser::parse_program(code).expect("parse with imports failed");
    println!("[with imports] Declarations: {}", prog.declarations.len());
    for d in &prog.declarations {
        match d {
            vybe_parser::ast::Declaration::Sub(s) => println!("  Sub: {}", s.name.as_str()),
            vybe_parser::ast::Declaration::Function(f) => println!("  Function: {}", f.name.as_str()),
            _ => println!("  (other kind)"),
        }
    }
    println!("[with imports] Statements: {}", prog.statements.len());
    assert!(prog.declarations.iter().any(|d| matches!(d, vybe_parser::ast::Declaration::Sub(s) if s.name.as_str().eq_ignore_ascii_case("Main"))),
        "Sub Main should be in top-level declarations when Imports present");
}

#[test]
fn test_parse_without_imports() {
    let code = r#"Module CollectionsDemo
    Sub Main()
        Console.WriteLine("hello")
    End Sub
End Module
"#;
    let prog = vybe_parser::parse_program(code).expect("parse without imports failed");
    println!("[no imports] Declarations: {}", prog.declarations.len());
    for d in &prog.declarations {
        match d {
            vybe_parser::ast::Declaration::Sub(s) => println!("  Sub: {}", s.name.as_str()),
            _ => println!("  (other kind)"),
        }
    }
    assert!(prog.declarations.iter().any(|d| matches!(d, vybe_parser::ast::Declaration::Sub(s) if s.name.as_str().eq_ignore_ascii_case("Main"))),
        "Sub Main should be in top-level declarations without imports");
}
