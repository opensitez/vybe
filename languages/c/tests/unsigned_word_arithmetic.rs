#[test]
fn unsigned_word_arithmetic_preserves_wrap_and_positive_result() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/unsigned_word_arithmetic.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let mut module = bundle.prepared_module().unwrap();
    let mut word_multiplications = 0;
    for statement in &mut module.body {
        statement.walk_exprs_mut(&mut |value| {
            if matches!(
                &value.kind,
                vybe_ast::ExprKind::Binary {
                    op: vybe_ast::BinOp::Integer(vybe_ast::IntOp::Mul, vybe_ast::BitLane::W32),
                    ..
                }
            ) {
                word_multiplications += 1;
            }
        });
    }
    assert!(word_multiplications > 0, "typed multiply path missing");
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
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
