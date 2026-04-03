use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value};

fn run_vb(source: &str) -> Vec<String> {
    let program = vybe_parser_basic::parse_program(source)
        .unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();

    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    // Re-setup namespaces after our log override
    vybe_host::setup_namespaces(&mut vm);

    let chunks = vybe_compiler_vb::Compiler::new().compile(&program)
        .unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));
    output.borrow().clone()
}

// ============================================================
// Namespace object access (struct_get chains)
// ============================================================

#[test]
fn math_namespace_object() {
    // Math.Floor accessed via namespace object (struct_get on global "math")
    // The compiler currently uses call_import for Math.Floor,
    // but the namespace object also works if accessed dynamically
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Math.Floor(3.7))
        Console.WriteLine(Math.Abs(-5))
        Console.WriteLine(Math.Sqrt(9))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3", "5", "3"]);
}

#[test]
fn console_namespace_object() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine("hello from namespace")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["hello from namespace"]);
}

#[test]
fn convert_namespace_object() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Convert.ToString(42))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn math_pi_constant() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Math.Floor(Math.PI))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3"]);
}
