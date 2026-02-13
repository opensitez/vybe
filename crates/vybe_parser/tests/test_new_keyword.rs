use vybe_parser::parse_program;

#[test]
fn test_new_expression_parses() {
    let code = r#"
Class Form1
    Inherits System.Windows.Forms.Form
    Private Sub InitializeComponent()
        Me.lstTasks = New System.Windows.Forms.ListBox()
        Me.ClientSize = New System.Drawing.Size(285, 270)
        Me.Text = "Todo List"
    End Sub
End Class
"#;
    let result = parse_program(code);
    match &result {
        Ok(prog) => {
            println!("Parsed OK: {} declarations", prog.declarations.len());
            for d in &prog.declarations {
                println!("  decl: {:?}", std::mem::discriminant(d));
            }
        }
        Err(e) => {
            println!("Parse ERROR: {}", e);
        }
    }
    assert!(result.is_ok(), "Designer code should parse: {:?}", result.err());
}

#[test]
fn test_new_task_identifier_not_new_expression() {
    let code = r#"
Module Test
    Sub Main()
        Dim newTask As String
        newTask = "hello"
    End Sub
End Module
"#;
    let result = parse_program(code);
    match &result {
        Ok(prog) => println!("Parsed OK: {} declarations", prog.declarations.len()),
        Err(e) => println!("Parse ERROR: {}", e),
    }
    assert!(result.is_ok(), "newTask should parse as identifier: {:?}", result.err());
}
