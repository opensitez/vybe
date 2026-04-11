/// Cross-language tests: JS class used from VB, VB class used from JS.
///
/// These tests require compiling two languages into the same VM with different
/// profiles, which is not yet supported in the vybex unified pipeline.

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value};

fn setup_vm() -> (VM, Arc<Mutex<Vec<String>>>) {
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    (vm, output)
}

#[test]
#[ignore = "cross-language requires shared VM multi-profile support"]
fn js_class_used_from_vb() {
    let (mut vm, output) = setup_vm();

    // Step 1: Compile and run JS that defines a class
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
    let js_module = vybex::languages::js::parse(js_code).expect("JS parse failed");
    let js_profile = load_js_profile();
    let js_chunks = vybex::compiler::Compiler::with_profile(js_profile)
        .compile(&js_module).expect("JS compile failed");
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
    let vb_module = vybex::languages::vb::parse(vb_code).expect("VB parse failed");
    let vb_profile = super::helpers::load_vb_profile();
    let vb_chunks = vybex::compiler::Compiler::with_profile(vb_profile)
        .compile(&vb_module).expect("VB compile failed");
    vm.run(vb_chunks).expect("VB runtime error");

    assert_eq!(output.lock().unwrap().as_slice(), &["13"]);
}

#[test]
#[ignore = "cross-language requires shared VM multi-profile support"]
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
    let vb_module = vybex::languages::vb::parse(vb_code).expect("VB parse failed");
    let vb_profile = super::helpers::load_vb_profile();
    let vb_chunks = vybex::compiler::Compiler::with_profile(vb_profile)
        .compile(&vb_module).expect("VB compile failed");
    vm.run(vb_chunks).expect("VB runtime error");

    // Verify VB class is in globals
    assert!(vm.globals.contains_key("greeter"),
        "VB Greeter class should be in globals");

    // Step 2: Compile and run JS that uses the VB class
    let js_code = r#"
        var g = new greeter("Hello");
        console.log(g.greet("World"));
    "#;
    let js_module = vybex::languages::js::parse(js_code).expect("JS parse failed");
    let js_profile = load_js_profile();
    let js_chunks = vybex::compiler::Compiler::with_profile(js_profile)
        .compile(&js_module).expect("JS compile failed");
    vm.run(js_chunks).expect("JS runtime error");

    assert_eq!(output.lock().unwrap().as_slice(), &["Hello World!"]);
}

#[test]
#[ignore = "cross-language requires shared VM multi-profile support"]
fn shared_global_between_languages() {
    let (mut vm, output) = setup_vm();

    // JS sets a global (function — JS stores lowercase)
    let js_code = "function getvalue() { return 42; }";
    let js_module = vybex::languages::js::parse(js_code).expect("parse");
    let js_profile = load_js_profile();
    let js_chunks = vybex::compiler::Compiler::with_profile(js_profile)
        .compile(&js_module).expect("compile");
    vm.run(js_chunks).expect("run");

    // VB calls the JS function (VB lowercases names — matches JS)
    let vb_code = "Console.WriteLine(getvalue())";
    let vb_module = vybex::languages::vb::parse(vb_code).expect("parse");
    let vb_profile = super::helpers::load_vb_profile();
    let vb_chunks = vybex::compiler::Compiler::with_profile(vb_profile)
        .compile(&vb_module).expect("compile");
    vm.run(vb_chunks).expect("run");

    assert_eq!(output.lock().unwrap().as_slice(), &["42"]);
}

#[test]
#[ignore = "cross-language requires shared VM multi-profile support"]
fn js_function_called_from_vb() {
    let (mut vm, output) = setup_vm();

    // JS defines a function
    let js_code = "function double(x) { return x * 2; }";
    let js_module = vybex::languages::js::parse(js_code).expect("parse");
    let js_profile = load_js_profile();
    let js_chunks = vybex::compiler::Compiler::with_profile(js_profile)
        .compile(&js_module).expect("compile");
    vm.run(js_chunks).expect("run");

    // VB calls it
    let vb_code = "Console.WriteLine(double(21))";
    let vb_module = vybex::languages::vb::parse(vb_code).expect("parse");
    let vb_profile = super::helpers::load_vb_profile();
    let vb_chunks = vybex::compiler::Compiler::with_profile(vb_profile)
        .compile(&vb_module).expect("compile");
    vm.run(vb_chunks).expect("run");

    assert_eq!(output.lock().unwrap().as_slice(), &["42"]);
}

/// Minimal JS profile for cross-language tests.
fn load_js_profile() -> vybex::profile::LanguageProfile {
    use vybex::profile::*;
    use std::collections::HashMap;

    LanguageProfile {
        function_return: ReturnStyle::Explicit,
        result_slot_name: "Result".into(),
        self_keyword: "this".into(),
        base_keyword: Some("super".into()),
        constructor_name: "constructor".into(),
        separated_methods: false,
        implicit_self_fields: false,
        explicit_self_param: false,
        enum_as_ordinals: false,
        case_sensitive: true,
        string_indexing: StringIndexing::ZeroBased,
        array_upper_bound_inclusive: false,
        parens_for_index: false,
        entry_point: None,
        hoist_var: true,
        dynamic_add: true,
        commonjs_require: true,
        partial_classes: false,
        byref_boxing: false,
        with_block: false,
        new_with_initializer: false,
        new_from_initializer: false,
        linq_queries: false,
        builtins: {
            let mut b = HashMap::new();
            b.insert("console.log".into(), BuiltinDef { emit: BuiltinEmit::Print, min_args: 0, max_args: 255 });
            b
        },
        intrinsics: HashMap::new(),
        namespaces: NamespaceConfig::default(),
        known_types: HashMap::new(),
        value_methods: HashMap::new(),
        module_aliases: HashMap::new(),
        namespace_constants: HashMap::new(),
        array_methods: HashMap::new(),
    }
}
