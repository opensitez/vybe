use std::path::Path;
use vybe_ast::{DirectiveScope, ExprKind, StmtKind, WideIntegerTransport};

fn fixture() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/c/doom_startup/fixed_mul_closed_word.c")
}

#[test]
fn actual_doom_fixed_mul_has_closed_exact_ast_without_module_enablement() {
    let path = fixture();
    let source = std::fs::read_to_string(&path).unwrap();
    let expanded = vybe_language_c::walker::expand_source(&path, &source).unwrap();
    let module = vybe_language_c::parse(&expanded).unwrap();
    assert_eq!(module.directives.wide_integer_transport, None);
    let body = module
        .body
        .iter()
        .find_map(|statement| {
            if let StmtKind::FunctionDecl { name, body, .. } = &statement.kind {
                (name == "FixedMul").then_some(body)
            } else {
                None
            }
        })
        .unwrap();
    assert_eq!(body.len(), 2);
    assert!(
        matches!(&body[0].kind, StmtKind::Directive { set, scope: DirectiveScope::Block }
        if set.wide_integer_transport == Some(WideIntegerTransport::Exact))
    );
    assert!(matches!(&body[1].kind, StmtKind::Return(Some(value))
        if matches!(&value.kind, ExprKind::Cast { type_name, .. } if type_name == "int")));
}

#[test]
fn actual_doom_fixed_mul_is_native_and_validates_without_metadata() {
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&fixture())
        .unwrap()
        .compile()
        .unwrap();
    let function = chunks
        .iter()
        .find(|chunk| chunk.name == "FixedMul")
        .unwrap();
    let dump = vybe_runtime::debug::disassemble(function);
    for opcode in ["i64.mul", "i64.shr_s", "i32.wrap_i64"] {
        assert!(dump.contains(opcode), "missing {opcode}: {dump}");
    }
    assert!(!dump.contains("f64.convert_i64"), "{dump}");
    assert!(
        !function
            .imports
            .iter()
            .any(|import| import.module == "ecma:bigint"),
        "{dump}"
    );
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
