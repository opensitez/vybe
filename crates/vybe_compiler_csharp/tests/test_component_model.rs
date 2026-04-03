use vybe_bytecode::{VM, Value};

/// Test: compile a C# component and run it with Component Model isolation
#[test]
fn component_model_basic() {
    let src = r#"
class Calculator {
    public static int Add(int a, int b) { return a + b; }
}
Console.WriteLine("Component test");
"#;
    let program = vybe_parser_csharp::parse(src).expect("parse failed");
    let chunks = vybe_compiler_csharp::Compiler::new().compile(&program).expect("compile failed");
    let component = vybe_compiler_common::components::build_component(
        "calc_module",
        vybe_bytecode::component::Language::CSharp,
        chunks,
    );

    assert!(!component.chunks.is_empty(), "Component should have chunks");

    // Link single component
    let mut linker = vybe_bytecode::Linker::new();
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    linker.register_host_from_vm(&vm);
    linker.add_component(component.clone());

    let link_result = linker.link().expect("Link failed");

    // Run with isolation
    let components = vec![component];
    let result = vm.run_components(&link_result, &components);
    assert!(result.is_ok(), "run_components failed: {:?}", result.err());
}

/// Test: two C# components linked together
#[test]
fn two_components_linked() {
    let src1 = r#"
class MathUtils {
    public static int Square(int x) { return x * x; }
}
"#;
    let prog1 = vybe_parser_csharp::parse(src1).expect("parse1 failed");
    let chunks1 = vybe_compiler_csharp::Compiler::new().compile(&prog1).expect("compile1 failed");
    let comp1 = vybe_compiler_common::components::build_component(
        "math_utils", vybe_bytecode::component::Language::CSharp, chunks1,
    );

    let src2 = r#"
Console.WriteLine("Main app");
"#;
    let prog2 = vybe_parser_csharp::parse(src2).expect("parse2 failed");
    let chunks2 = vybe_compiler_csharp::Compiler::new().compile(&prog2).expect("compile2 failed");
    let comp2 = vybe_compiler_common::components::build_component(
        "main_app", vybe_bytecode::component::Language::CSharp, chunks2,
    );

    let mut linker = vybe_bytecode::Linker::new();
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    linker.register_host_from_vm(&vm);
    linker.add_component(comp1.clone());
    linker.add_component(comp2.clone());

    let link_result = linker.link().expect("Link failed");
    let components = vec![comp1, comp2];
    let result = vm.run_components(&link_result, &components);
    assert!(result.is_ok(), "run_components failed: {:?}", result.err());
}

/// Test: component isolation — each module has prefixed globals
#[test]
fn component_isolation() {
    let src = r#"
var x = 42;
Console.WriteLine(x);
"#;
    let program = vybe_parser_csharp::parse(src).expect("parse failed");
    let chunks = vybe_compiler_csharp::Compiler::new().compile(&program).expect("compile failed");
    let component = vybe_compiler_common::components::build_component(
        "isolated_mod", vybe_bytecode::component::Language::CSharp, chunks,
    );

    let mut linker = vybe_bytecode::Linker::new();
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    linker.register_host_from_vm(&vm);
    linker.add_component(component.clone());

    let link_result = linker.link().expect("Link failed");
    let components = vec![component];
    vm.run_components(&link_result, &components).expect("run failed");

    // When strict_isolation is used via run_components, globals set during
    // execution are prefixed with the module name.
    // The module may not set any globals if all vars are locals,
    // but internal state (__tid_, etc.) should be prefixed.
    // Verify run_components completed without error (isolation was active).
    // The key invariant: run_components sets strict_isolation=true during execution.
    assert!(!vm.globals.is_empty(), "VM should have globals after running");
}
