use vybe_parser::parse_program;
use vybe_runtime::interpreter::Interpreter;

#[test]
fn test_runtime_lookup() {
    let code = "Sub Form_Load()\nEnd Sub";
    let program = parse_program(code).expect("Failed to parse");
    
    let mut interp = Interpreter::new();
    let module_name = "Form1";
    
    interp.load_module(module_name, &program).expect("Failed to load module");
    
    //Check internal state
    println!("Subs: {:?}", interp.subs.keys());
    
    // Check if key exists
    let expected_key = "form1.form_load";
    assert!(interp.subs.contains_key(expected_key), "Key '{}' not found in subs", expected_key);
    
    // Call event handler
    interp.call_event_handler("Form_Load", &[]).expect("Failed to call Form_Load");
}

#[test]
fn test_runtime_lookup_case_insensitive() {
    let code = "Sub FORM_LOAD()\nEnd Sub";
    let program = parse_program(code).expect("Failed to parse");
    
    let mut interp = Interpreter::new();
    let module_name = "Form1";
    
    interp.load_module(module_name, &program).expect("Failed to load module");
    
    //Check internal state
    println!("Subs: {:?}", interp.subs.keys());
    
    // Check if key exists
    let expected_key = "form1.form_load";
    assert!(interp.subs.contains_key(expected_key), "Key '{}' not found in subs", expected_key);
    
    // Call event handler
    interp.call_event_handler("Form_Load", &[]).expect("Failed to call Form_Load");
}
