use vybe_runtime::{Interpreter, Value};
use vybe_parser::parse_program;

#[test]
fn test_not_unary_parsing() {
    let code = r#"
    Public Class Form1
        Private gameOver As Boolean
        
        Public Sub Test()
            gameOver = False
            If Not gameOver Then
                Console.WriteLine("Not gameOver works")
            End If
        End Sub
    End Class
    "#;

    let prog = parse_program(code).expect("Failed to parse Not unary");
    println!("Parsed successfully");
}
