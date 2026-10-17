use vybe_runtime::{Chunk, Op, VM, Value};

#[test]
fn writer_keeps_immediate_bytes_while_remapping_locals() {
    let mut chunk = Chunk::new("inhouse-immediate-boundaries");
    chunk.local_count = 1;
    // These floating-point payload bytes resemble local/control opcodes.
    let bits = u64::from_le_bytes([0x20, 0x23, 0x0d, 0x03, 0x0b, 0x24, 0x21, 0x22]);
    chunk.emit_f64_const(f64::from_bits(bits), 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let observer = chunk.add_import("inhouse-boundary-test", "observe");
    chunk.emit_call(observer, 1, 0);
    chunk.emit_op(Op::RETURN, 0);
    let wasm = vybe_platform_wasm::write_wasm(&[chunk]);
    let decoded = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    let mut vm = VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    vybe_platform_ecma::register(&mut vm);
    let observed = std::sync::Arc::new(std::sync::Mutex::new(Vec::new()));
    let captured = observed.clone();
    vm.register_host_fn(
        "inhouse-boundary-test",
        "observe",
        Box::new(move |_, args| {
            let Value::F64(value) = args[0] else {
                panic!("expected original f64, got {:?}", args[0]);
            };
            captured.lock().unwrap().push(value.to_bits());
            args[0].clone()
        }),
    );
    vm.run(decoded).unwrap();
    assert_eq!(&*observed.lock().unwrap(), &[bits]);
}

#[test]
fn declared_reference_call_fixtures_still_roundtrip() {
    let fixtures =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/declared_ref_calls");
    for name in [
        "i32",
        "i64",
        "f32",
        "f64",
        "externref",
        "funcref",
        "mixed",
        "tail_i32",
        "multiple",
    ] {
        let bytes = std::fs::read(fixtures.join(format!("{name}.wasm"))).unwrap();
        let chunks = vybe_platform_wasm::read_wasm(&bytes).unwrap();
        let emitted = vybe_platform_wasm::write_wasm(&chunks);
        vybe_platform_wasm::read_wasm(&emitted).unwrap_or_else(|error| panic!("{name}: {error}"));
    }
}
