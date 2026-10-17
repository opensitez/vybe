use std::path::Path;
use vybe_ast::ExprKind;
use vybe_runtime::VM;

#[test]
fn postfix_byte_reads_preserve_linear_and_string_backing() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/scoped_byte_pointer.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let mut prepared = bundle.prepared_module().unwrap();
    let mut backing_reads = 0;
    for statement in &mut prepared.body {
        statement.walk_exprs_mut(&mut |value| {
            if matches!(&value.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name)
                    if name == "__c_char_ptr_read"))
            {
                backing_reads += 1;
            }
        });
    }
    assert!(
        backing_reads > 0,
        "mixed byte reads must retain their backing"
    );
    let chunks = bundle.compile().unwrap();
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
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
