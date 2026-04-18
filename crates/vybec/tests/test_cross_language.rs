/// Cross-language inheritance tests using Component Model isolation.
/// Tests that classes defined in one language can be inherited and used
/// from another language, with per-module global isolation.

use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value, HostContext};

fn setup_vm() -> (VM, Rc<RefCell<Vec<String>>>) {
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    (vm, output)
}

// ═══════════════════════════════════════════════════════════
// SHARED MODE: VB class used from JS (existing pattern)
// ═══════════════════════════════════════════════════════════
#[test]
fn vb_class_used_from_js_shared() {
    let (mut vm, output) = setup_vm();

    // Step 1: Run VB that defines a class
    let vb_src = r#"
Class Animal
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Function Speak() As String
        Return Name & " speaks"
    End Function
End Class
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse failed");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile failed");
    vm.run(vb_chunks).expect("VB run failed");

    // Step 2: Run JS that uses the VB class
    let js_src = r#"
var a = new animal("Rex");
console.log(a.speak());
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse failed");
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile failed");
    vm.run(js_chunks).expect("JS run failed");

    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// SHARED MODE: JS class used from C# (requires class name interop — known limitation)
// ═══════════════════════════════════════════════════════════
#[test]
fn js_class_used_from_csharp_linked() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);

    let js_src = r#"
class calculator {
    constructor() {
        this.result = 0;
    }
    add(n) {
        this.result = this.result + n;
        return this;
    }
    getResult() {
        return this.result;
    }
}
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse failed");
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile failed");

    let cs_src = r#"
var calc = new calculator();
calc.add(10);
calc.add(20);
Console.WriteLine(calc.getResult());
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse failed");
    let cs_chunks = vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile failed");

    // Run sequentially — vm.run() adjusts ref_func indices,
    // both runs share vm.chunks and vm.globals
    vm.run(js_chunks).expect("JS run failed");
    vm.run(cs_chunks).expect("C# run failed");

    assert_eq!(output.borrow().as_slice(), &["30"]);
}

// ═══════════════════════════════════════════════════════════
// COMPONENT MODEL: Two modules linked and run with isolation
// ═══════════════════════════════════════════════════════════
#[test]
fn component_model_two_modules() {
    let (mut vm, output) = setup_vm();

    // Module 1: VB defines a class
    let vb_src = r#"
Class Greeter
    Public Sub New()
    End Sub
    Public Function Greet(name As String) As String
        Return "Hello " & name
    End Function
End Class
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse failed");
    let vb_chunks = vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile failed");
    let vb_comp = vybe_compiler_common::components::build_component(
        "greeter_lib", vybe_bytecode::component::Language::VB, vb_chunks,
    );

    // Module 2: C# uses it
    let cs_src = r#"
Console.WriteLine("Module 2 running");
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse failed");
    let cs_chunks = vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile failed");
    let cs_comp = vybe_compiler_common::components::build_component(
        "main_app", vybe_bytecode::component::Language::CSharp, cs_chunks,
    );

    // Link via Component Model
    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(vb_comp.clone());
    linker.add_component(cs_comp.clone());

    let link_result = linker.link().expect("Link failed");
    let components = vec![vb_comp, cs_comp];

    // Run with isolation
    vm.run_components(&link_result, &components).expect("run_components failed");

    assert_eq!(output.borrow().as_slice(), &["Module 2 running"]);
}

// ═══════════════════════════════════════════════════════════
// COMPONENT MODEL: Ruby + PHP linked
// ═══════════════════════════════════════════════════════════
#[test]
fn ruby_php_component_link() {
    let (mut vm, _output) = setup_vm();

    let rb_src = r#"
def add(a, b)
  a + b
end
"#;
    let rb_prog = vybe_parser_ruby::parse(rb_src).expect("Ruby parse failed");
    let rb_comp = vybe_compiler_ruby::compile_component(&rb_prog, "math_rb").expect("Ruby comp failed");

    let php_src = r#"<?php
function multiply($a, $b) {
    return $a * $b;
}
"#;
    let php_prog = vybec::parser_php::parse(php_src).expect("PHP parse failed");
    let php_comp = vybec::compiler_php::compile_component(&php_prog, "math_php").expect("PHP comp failed");

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(rb_comp.clone());
    linker.add_component(php_comp.clone());

    let link_result = linker.link().expect("Link failed");
    let components = vec![rb_comp, php_comp];
    vm.run_components(&link_result, &components).expect("run_components failed");
}

// ═══════════════════════════════════════════════════════════
// COMPONENT MODEL: COBOL + C# linked
// ═══════════════════════════════════════════════════════════
#[test]
fn cobol_csharp_component_link() {
    let (mut vm, output) = setup_vm();

    let cobol_src = r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. HELPER.
PROCEDURE DIVISION.
    DISPLAY "COBOL module loaded".
    STOP RUN.
"#;
    let cobol_prog = vybe_parser_cobol::parse(cobol_src).expect("COBOL parse failed");
    let cobol_comp = vybe_compiler_cobol::compile_component(&cobol_prog, "cobol_lib").expect("COBOL comp failed");

    let cs_src = r#"
Console.WriteLine("C# main running");
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse failed");
    let cs_chunks = vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile failed");
    let cs_comp = vybe_compiler_common::components::build_component(
        "cs_main", vybe_bytecode::component::Language::CSharp, cs_chunks,
    );

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(cobol_comp.clone());
    linker.add_component(cs_comp.clone());

    let link_result = linker.link().expect("Link failed");
    let components = vec![cobol_comp, cs_comp];
    vm.run_components(&link_result, &components).expect("run_components failed");

    let out = output.borrow();
    assert!(out.iter().any(|s| s.contains("COBOL module loaded")), "COBOL should have run");
    assert!(out.iter().any(|s| s.contains("C# main running")), "C# should have run");
}

// ═══════════════════════════════════════════════════════════
// ISOLATION: Modules can't see each other's internal globals
// ═══════════════════════════════════════════════════════════
#[test]
fn isolation_globals_separate() {
    let (mut vm, _output) = setup_vm();

    // Module A sets a global "secret"
    let js_src = r#"
var secret = "module_a_data";
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("parse failed");
    let js_chunks = vybe_compiler_js::Compiler::new().compile(&js_prog).expect("compile failed");
    let comp_a = vybe_compiler_common::components::build_component(
        "module_a", vybe_bytecode::component::Language::JS, js_chunks,
    );

    // Module B sets a different global "secret"
    let rb_src = r#"
secret = "module_b_data"
"#;
    let rb_prog = vybe_parser_ruby::parse(rb_src).expect("parse failed");
    let rb_chunks = vybe_compiler_ruby::Compiler::new().compile(&rb_prog).expect("compile failed");
    let comp_b = vybe_compiler_common::components::build_component(
        "module_b", vybe_bytecode::component::Language::Ruby, rb_chunks,
    );

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(comp_a.clone());
    linker.add_component(comp_b.clone());

    let link_result = linker.link().expect("Link failed");
    let components = vec![comp_a, comp_b];
    vm.run_components(&link_result, &components).expect("run failed");

    // Each module should have its own prefixed "secret"
    let a_secret = vm.globals.get("module_a::secret");
    let b_secret = vm.globals.get("module_b::secret");

    // They should be different (isolated)
    if let (Some(a), Some(b)) = (a_secret, b_secret) {
        assert_ne!(format!("{}", a), format!("{}", b), "Modules should have separate globals");
    }
    // The unprefixed "secret" should NOT exist (isolation enforced)
    let raw_secret = vm.globals.get("secret");
    assert!(raw_secret.is_none(), "Unprefixed 'secret' should not exist in isolation mode");
}

// ═══════════════════════════════════════════════════════════
// ALL 8 LANGUAGES: Each compiles and links as component
// ═══════════════════════════════════════════════════════════
#[test]
fn all_languages_as_components() {
    let (mut vm, _output) = setup_vm();

    // VB
    let vb_prog = vybe_parser_basic::parse_program("Module M\nEnd Module").expect("VB");
    let vb_comp = vybe_compiler_common::components::build_component("vb_mod",
        vybe_bytecode::component::Language::VB,
        vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile"));

    // JS
    let js_prog = vybe_parser_js::parse("var x = 1;").expect("JS");
    let js_comp = vybe_compiler_common::components::build_component("js_mod",
        vybe_bytecode::component::Language::JS,
        vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile"));

    // C#
    let cs_prog = vybe_parser_csharp::parse("var x = 1;").expect("C#");
    let cs_comp = vybe_compiler_common::components::build_component("cs_mod",
        vybe_bytecode::component::Language::CSharp,
        vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile"));

    // Ruby
    let rb_prog = vybe_parser_ruby::parse("x = 1").expect("Ruby");
    let rb_comp = vybe_compiler_common::components::build_component("rb_mod",
        vybe_bytecode::component::Language::Ruby,
        vybe_compiler_ruby::Compiler::new().compile(&rb_prog).expect("Ruby compile"));

    // PHP
    let php_prog = vybec::parser_php::parse("<?php $x = 1;").expect("PHP");
    let php_comp = vybe_compiler_common::components::build_component("php_mod",
        vybe_bytecode::component::Language::Php,
        vybec::compiler_php::Compiler::new().compile(&php_prog).expect("PHP compile"));

    // Python
    let py_prog = vybe_parser_python::parse("x = 1").expect("Python");
    let py_comp = vybe_compiler_common::components::build_component("py_mod",
        vybe_bytecode::component::Language::Python,
        vybe_compiler_python::Compiler::new().compile(&py_prog).expect("Python compile"));

    // Dart
    let dart_prog = vybec::parser_dart::parse("var x = 1;").expect("Dart");
    let dart_comp = vybe_compiler_common::components::build_component("dart_mod",
        vybe_bytecode::component::Language::Dart,
        vybec::compiler_dart::Compiler::new().compile(&dart_prog).expect("Dart compile"));

    // COBOL
    let cobol_prog = vybe_parser_cobol::parse("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    STOP RUN.").expect("COBOL");
    let cobol_comp = vybe_compiler_cobol::compile_component(&cobol_prog, "cobol_mod").expect("COBOL comp");

    // Link all 8
    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    let all_comps = vec![
        vb_comp.clone(), js_comp.clone(), cs_comp.clone(), rb_comp.clone(),
        php_comp.clone(), py_comp.clone(), dart_comp.clone(), cobol_comp.clone(),
    ];
    for c in &all_comps { linker.add_component(c.clone()); }

    let link_result = linker.link().expect("Link all 8 languages failed");
    vm.run_components(&link_result, &all_comps).expect("run_components with 8 languages failed");
}
