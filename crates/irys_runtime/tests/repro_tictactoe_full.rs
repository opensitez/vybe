use irys_runtime::{Interpreter, Value};
use irys_parser::parse_program;
use std::fs;
use std::path::PathBuf;

#[test]
fn test_tictactoe_logic_full() {
    let project_root = PathBuf::from("../../examples/TicTacToe");
    let form_path = project_root.join("Form1.vb");

    let form_code = std::fs::read_to_string(&form_path).expect("Failed to read Form1.vb");
    
    let mut interp = Interpreter::new();
    let prog = parse_program(&form_code).expect("Failed to parse Form1");
    interp.load_module("Form1", &prog).expect("Failed to load Form1 module");
    
    // Globals
    let global_setup = r#"
    Module RuntimeGlobals
        Public Form1Instance As New Form1
    End Module
    "#;
    let setup_prog = parse_program(global_setup).expect("Failed to parse setup");
    interp.load_module("RuntimeGlobals", &setup_prog).expect("Failed to load setup");
    interp.init_namespaces();

    // Mock Controls Classes
    let mock_controls = r#"
    Public Class StdButton
        Public Name As String
        Public Caption As String
        Public Enabled As Boolean
        Public Sub Click()
            ' No-op
        End Sub
    End Class
    
    Public Class StdLabel
        Public Caption As String
    End Class
    "#;
    let mock_prog = parse_program(mock_controls).expect("Failed to parse mocks");
    interp.load_module("Mocks", &mock_prog).expect("Failed to load mocks");

    // Setup Script to Initialize Controls on the Instance
    let setup_script = r#"
    Module Setup
        Public Sub InitControls()
            ' Dynamic generation modeled after RuntimePanel.rs
            RuntimeGlobals.Form1Instance.btn0 = New StdButton
            RuntimeGlobals.Form1Instance.btn0.Name = "btn0"
            RuntimeGlobals.Form1Instance.btn0.Caption = ""
            RuntimeGlobals.Form1Instance.btn0.Enabled = True
            
            RuntimeGlobals.Form1Instance.btn1 = New StdButton
            RuntimeGlobals.Form1Instance.btn1.Name = "btn1"
            RuntimeGlobals.Form1Instance.btn1.Caption = ""
            RuntimeGlobals.Form1Instance.btn1.Enabled = True
            
            ' ... (In a real reproduction we would use the iterator, but for this test file 
            ' we are hardcoding. The key is to match the order and names exactly as RuntimePanel would)
            
            RuntimeGlobals.Form1Instance.btnReset = New StdButton
            RuntimeGlobals.Form1Instance.btnReset.Name = "btnReset"
             RuntimeGlobals.Form1Instance.btnReset.Caption = "Reset Game"
             RuntimeGlobals.Form1Instance.btnReset.Enabled = True
        End Sub
    End Module
    "#;
    
    // BETTER APPROACH: Use the actual Form module to generate the initialization code
    // This requires extracting the form from the parsed project, which we can do via finding the class
    // But since we don't have the full Project struct here easily, let's simulates the error condition:
    // The user says "btn0" resolves to "btnReset". 
    // This happens if "btn0" field in Form1Instance points to the btnReset object.
    
            RuntimeGlobals.Form1Instance.btn2 = New StdButton
            RuntimeGlobals.Form1Instance.btn2.Name = "btn2"
            RuntimeGlobals.Form1Instance.btn2.Caption = ""
            RuntimeGlobals.Form1Instance.btn2.Enabled = True
             
            RuntimeGlobals.Form1Instance.btn3 = New StdButton
            RuntimeGlobals.Form1Instance.btn3.Name = "btn3"
            RuntimeGlobals.Form1Instance.btn3.Caption = ""
            RuntimeGlobals.Form1Instance.btn3.Enabled = True
             
            RuntimeGlobals.Form1Instance.btn4 = New StdButton
            RuntimeGlobals.Form1Instance.btn4.Name = "btn4"
            RuntimeGlobals.Form1Instance.btn4.Caption = ""
            RuntimeGlobals.Form1Instance.btn4.Enabled = True
             
            RuntimeGlobals.Form1Instance.btn5 = New StdButton
            RuntimeGlobals.Form1Instance.btn5.Name = "btn5"
            RuntimeGlobals.Form1Instance.btn5.Caption = ""
            RuntimeGlobals.Form1Instance.btn5.Enabled = True
             
            RuntimeGlobals.Form1Instance.btn6 = New StdButton
            RuntimeGlobals.Form1Instance.btn6.Name = "btn6"
            RuntimeGlobals.Form1Instance.btn6.Caption = ""
            RuntimeGlobals.Form1Instance.btn6.Enabled = True
             
            RuntimeGlobals.Form1Instance.btn7 = New StdButton
            RuntimeGlobals.Form1Instance.btn7.Name = "btn7"
            RuntimeGlobals.Form1Instance.btn7.Caption = ""
            RuntimeGlobals.Form1Instance.btn7.Enabled = True
             
            RuntimeGlobals.Form1Instance.btn8 = New StdButton
            RuntimeGlobals.Form1Instance.btn8.Name = "btn8"
            RuntimeGlobals.Form1Instance.btn8.Caption = ""
            RuntimeGlobals.Form1Instance.btn8.Enabled = True
             
            RuntimeGlobals.Form1Instance.btnReset = New StdButton
            RuntimeGlobals.Form1Instance.btnReset.Name = "btnReset"
            RuntimeGlobals.Form1Instance.btnReset.Caption = "Reset Game"
            RuntimeGlobals.Form1Instance.btnReset.Enabled = True
            
            RuntimeGlobals.Form1Instance.lblStatus = New StdLabel
            RuntimeGlobals.Form1Instance.lblStatus.Caption = "Ready"
        End Sub
    End Module
    "#;
    let setup_script_prog = parse_program(setup_script).expect("Failed to parse setup script");
    interp.load_module("Setup", &setup_script_prog).expect("Failed to load setup script");

    // Run Initialization
    interp.call_procedure(&irys_parser::ast::Identifier::new("InitControls"), &[]).expect("InitControls failed");
    
    // Call Form_Load
    interp.call_instance_method("RuntimeGlobals.Form1Instance", "Form1_Load", &[]).expect("Form_Load failed");

    // Verify turn initialized
    let turn = interp.evaluate_expr(&irys_parser::parse_expression_str("RuntimeGlobals.Form1Instance.turn").unwrap()).unwrap();
     if let Value::String(s) = turn {
        assert_eq!(s, "X", "Turn should be X after load");
    } else {
        panic!("Turn is not a string");
    }

    // Call btn0_Click
    interp.call_instance_method("RuntimeGlobals.Form1Instance", "btn0_Click", &[]).expect("btn0_Click failed");

    // Verify Caption
    let cap = interp.evaluate_expr(&irys_parser::parse_expression_str("RuntimeGlobals.Form1Instance.btn0.Caption").unwrap()).unwrap();
    if let Value::String(s) = cap {
        println!("Checked Caption: '{}'", s);
        assert_eq!(s, "X", "btn0.Caption should be X");
    } else {
        panic!("Caption is not string");
    }
}
