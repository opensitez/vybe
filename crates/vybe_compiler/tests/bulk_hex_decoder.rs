#[allow(dead_code)]
#[path = "fixtures/bulk_hex_decoder_reference.rs"]
mod reference;
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
fn worker(fast: bool) -> (VM, Value) {
    let mut root = Chunk::new("owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let call = if fast {
        let mut imports = Chunk::new("imports");
        primitives::polyfills::build_runtime_helpers_for_exports(
            &mut imports,
            &["__stdlib_encoding_hex2bin"],
        )
        .chunks
        .remove(0)
    } else {
        let mut c = Chunk::new("hex");
        c.arity = 1;
        c.local_count = 1;
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        reference::emit_hex2bin(std::slice::from_mut(&mut c), 0, 1, 0);
        c.emit_op(Op::RETURN, 0);
        c
    };
    let mut chunks = vec![root, call];
    primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
    primitives::globals::declare_free_globals(&mut chunks);
    primitives::globals::normalize_global_table(&mut chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
        .unwrap();
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(chunks).unwrap();
    (vm, f)
}
#[test]
fn bulk_decoder_preserves_complete_and_invalid_inputs() {
    let mut workers = [worker(false), worker(true)];
    for input in [
        String::new(),
        "00ffAbCd".into(),
        "0".into(),
        "abggcd".into(),
        "ab cd".into(),
        "0é".into(),
        "ab🙂cd".into(),
        "ab\0cd".into(),
        "61".repeat(8192),
        (0..=255).map(|b| format!("{b:02x}")).collect(),
    ] {
        let values = workers.each_mut().map(|(vm, f)| {
            vm.invoke(f, &[Value::String(input.clone().into())])
                .unwrap()
                .to_string()
        });
        assert_eq!(values[0], values[1], "{input:?}");
    }
    for code in 0..=65535 {
        let Some(unit) = char::from_u32(code) else {
            continue;
        };
        for input in [format!("0{unit}"), format!("{unit}0")] {
            let values = workers.each_mut().map(|(vm, f)| {
                vm.invoke(f, &[Value::String(input.clone().into())])
                    .unwrap()
                    .to_string()
            });
            assert_eq!(values[0], values[1], "unit={code:04x}");
        }
    }
}
