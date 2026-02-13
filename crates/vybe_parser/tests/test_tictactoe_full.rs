use vybe_parser::parse_program;

#[test]
fn test_full_tictactoe() {
    let designer = include_str!("../../../examples/TicTacToe/Form1.Designer.vb");
    let user_code = include_str!("../../../examples/TicTacToe/Form1.vb");
    let combined = format!("{}\n{}", designer, user_code);
    
    let result = parse_program(&combined);
    match &result {
        Ok(prog) => {
            println!("Parsed OK: {} declarations", prog.declarations.len());
            for d in &prog.declarations {
                match d {
                    vybe_parser::Declaration::Class(cls) => {
                        println!("  Class: {} (methods: {}, fields: {})", 
                            cls.name.as_str(), cls.methods.len(), cls.fields.len());
                        for m in &cls.methods {
                            match m {
                                vybe_parser::ast::decl::MethodDecl::Sub(s) => {
                                    println!("    Sub: {} ({} stmts)", s.name.as_str(), s.body.len());
                                }
                                vybe_parser::ast::decl::MethodDecl::Function(f) => {
                                    println!("    Function: {} ({} stmts)", f.name.as_str(), f.body.len());
                                }
                            }
                        }
                    }
                    _ => println!("  other decl"),
                }
            }
        }
        Err(e) => {
            println!("Parse ERROR: {}", e);
        }
    }
    assert!(result.is_ok(), "TicTacToe should parse: {:?}", result.err());
}
