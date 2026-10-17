use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
fn worker(decode: bool) -> (VM, Value) {
    let mut root = Chunk::new("uu-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("uu-call");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut chunks = vec![root, c];
    if decode {
        primitives::string_encoding::emit_convert_uudecode(&mut chunks, 1, 1, 0);
    } else {
        primitives::string_encoding::emit_convert_uuencode(&mut chunks, 1, 1, 0);
    }
    chunks[1].emit_op(Op::RETURN, 0);
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
fn uu_helpers_preserve_utf16_units_and_malformed_line_behavior() {
    let mut encoder = worker(false);
    let mut decoder = worker(true);
    // The shared primitive operates on UTF16 units; byte coercion belongs to adapters.
    for (input, encoded) in [
        ("", " \n"),
        ("Cat", "#0V%T\n"),
        ("a", "!80``\n"),
        ("ab", "\"86(`\n"),
        ("é", "!Z0``\n"),
        ("Ω", "!Ċ0``\n"),
    ] {
        let actual = encoder
            .0
            .invoke(&encoder.1, &[Value::String(input.into())])
            .unwrap();
        assert_eq!(actual, Value::String(encoded.into()), "{input:?}");
    }
    for (encoded, expected) in [
        ("#0V%T\n", Value::String("Cat".into())),
        ("!80``\n", Value::String("a".into())),
        ("`", Value::String("".into())),
        ("", Value::Bool(false)),
        ("#0V%", Value::Bool(false)),
    ] {
        assert_eq!(
            decoder
                .0
                .invoke(&decoder.1, &[Value::String(encoded.into())])
                .unwrap(),
            expected,
            "{encoded:?}"
        );
    }
    for w in [&mut encoder, &mut decoder] {
        for input in [
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::I32(7),
        ] {
            assert!(w.0.invoke(&w.1, &[input]).is_err());
        }
    }
}
