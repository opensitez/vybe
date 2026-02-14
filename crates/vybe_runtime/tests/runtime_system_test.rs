use vybe_runtime::interpreter::Interpreter;
use vybe_runtime::value::Value;
use vybe_parser::parse_program;

#[test]
fn test_system_drawing_primitives() {
    let mut interpreter = Interpreter::new();
    
    // Test Color.Red
    let code = r#"
        Dim c = System.Drawing.Color.Red
        Dim r = c.R
    "#;
    let program = parse_program(code).expect("Failed to parse");
    interpreter.run(&program).expect("Failed to run Color test");
    let c = interpreter.env.get("c").expect("c not found");
    if let Value::Object(obj) = c {
        let b = obj.borrow();
        assert_eq!(b.class_name, "System.Drawing.Color");
        // Verify R channel (Red is 255, 0, 0)
        let r = b.fields.get("r").expect("r field missing").as_integer().unwrap();
        assert_eq!(r, 255);
    } else {
        panic!("c is not an object");
    }

    // Test New Pen(Color.Blue, 2.0)
    let code_pen = r#"
        Dim p = New System.Drawing.Pen(System.Drawing.Color.Blue, 2.0)
    "#;
    let program_pen = parse_program(code_pen).expect("Failed to parse pen");
    interpreter.run(&program_pen).expect("Failed to run Pen test");
    let p = interpreter.env.get("p").expect("p not found");
    if let Value::Object(obj) = p {
        let b = obj.borrow();
        assert_eq!(b.class_name, "System.Drawing.Pen");
        let w = b.fields.get("width").expect("width missing").as_double().unwrap();
        assert_eq!(w, 2.0);
    } else {
        panic!("p is not an object");
    }
}

#[test]
fn test_system_io_path() {
    let mut interpreter = Interpreter::new();
    
    let code = r#"
        Dim p = System.IO.Path.Combine("folder", "file.txt")
        Dim ext = System.IO.Path.GetExtension(p)
    "#;
    let program = parse_program(code).expect("Failed to parse IO");
    interpreter.run(&program).expect("Failed to run Path test");
    
    let p = interpreter.env.get("p").expect("p not found").as_string();
    assert!(p.ends_with("file.txt"));
    assert!(p.contains("folder"));
    
    let ext = interpreter.env.get("ext").expect("ext not found").as_string();
    assert_eq!(ext, ".txt");
}

#[test]
fn test_messagebox_stub() {
    // MessageBox.Show should return MsgBoxResult (Integer 1 for OK)
    // We can't easily test visual output, but we can verify it doesn't crash
    let mut interpreter = Interpreter::new();
    let code = r#"
        Dim result = System.Windows.Forms.MessageBox.Show("Test")
    "#;
    // Note: MsgBox might block or print to stdout depending on implementation.
    // In test env, it usually just returns 1 (OK).
    let program = parse_program(code).expect("Failed to parse MsgBox");
    interpreter.run(&program).expect("Failed to run MessageBox test");
    // let res = interpreter.env.get("result").expect("result not found").as_integer().unwrap();
    // assert_eq!(res, 1);
}
