/// Cross-language tests: JS class used from VB, VB class used from JS.

use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value};

fn setup_vm() -> (VM, Rc<RefCell<Vec<String>>>) {
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    (vm, output)
}

#[test]
fn js_class_used_from_vb() {
    let (mut vm, output) = setup_vm();

    // Step 1: Compile and run JS that defines a class
    // Note: JS stores as "Counter" but VB lowercases to "counter".
    // For cross-language, we also store lowercase alias.
    let js_code = r#"
        class counter {
            constructor(start) {
                this.count = start;
            }
            inc() {
                this.count = this.count + 1;
            }
            get() {
                return this.count;
            }
        }
    "#;
    let js_program = vybe_parser_js::parse(js_code).expect("JS parse failed");
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_program).expect("JS compile failed");
    vm.run(js_chunks).expect("JS runtime error");

    // Verify JS class is in globals
    assert!(vm.globals.contains_key("Counter") || vm.globals.contains_key("counter"),
        "JS Counter class should be in globals");

    // Step 2: Compile and run VB that uses the JS class
    let vb_code = r#"
Dim c As New Counter(10)
c.inc()
c.inc()
c.inc()
Console.WriteLine(c.get())
"#;
    let vb_program = vybe_parser_basic::parse_program(vb_code).expect("VB parse failed");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_program).expect("VB compile failed");
    vm.run(vb_chunks).expect("VB runtime error");

    assert_eq!(output.borrow().as_slice(), &["13"]);
}

#[test]
fn vb_class_used_from_js() {
    let (mut vm, output) = setup_vm();

    // Step 1: Compile and run VB that defines a class
    let vb_code = r#"
Public Class Greeter
    Dim prefix As String
    Public Sub New(p As String)
        prefix = p
    End Sub
    Public Function Greet(name As String) As String
        Return prefix & " " & name & "!"
    End Function
End Class
"#;
    let vb_program = vybe_parser_basic::parse_program(vb_code).expect("VB parse failed");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_program).expect("VB compile failed");
    vm.run(vb_chunks).expect("VB runtime error");

    // Verify VB class is in globals
    assert!(vm.globals.contains_key("greeter"),
        "VB Greeter class should be in globals");

    // Step 2: Compile and run JS that uses the VB class
    let js_code = r#"
        var g = new greeter("Hello");
        console.log(g.greet("World"));
    "#;
    let js_program = vybe_parser_js::parse(js_code).expect("JS parse failed");
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_program).expect("JS compile failed");
    vm.run(js_chunks).expect("JS runtime error");

    assert_eq!(output.borrow().as_slice(), &["Hello World!"]);
}

#[test]
fn shared_global_between_languages() {
    let (mut vm, output) = setup_vm();

    // JS sets a global (function — JS stores lowercase)
    let js_code = "function getvalue() { return 42; }";
    let js_program = vybe_parser_js::parse(js_code).expect("parse");
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_program).expect("compile");
    vm.run(js_chunks).expect("run");

    // VB calls the JS function (VB lowercases names — matches JS)
    let vb_code = "Console.WriteLine(getvalue())";
    let vb_program = vybe_parser_basic::parse_program(vb_code).expect("parse");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_program).expect("compile");
    vm.run(vb_chunks).expect("run");

    assert_eq!(output.borrow().as_slice(), &["42"]);
}

#[test]
fn js_function_called_from_vb() {
    let (mut vm, output) = setup_vm();

    // JS defines a function
    let js_code = "function double(x) { return x * 2; }";
    let js_program = vybe_parser_js::parse(js_code).expect("parse");
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_program).expect("compile");
    vm.run(js_chunks).expect("run");

    // VB calls it
    let vb_code = "Console.WriteLine(double(21))";
    let vb_program = vybe_parser_basic::parse_program(vb_code).expect("parse");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_program).expect("compile");
    vm.run(vb_chunks).expect("run");

    assert_eq!(output.borrow().as_slice(), &["42"]);
}
