use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, Value, VM};

#[test]
fn base64_character_predicate_preserves_both_alphabets_and_signed_boundaries() {
    for url in [false, true] {
        let mut body = Chunk::new("base64-character-predicate");
        body.arity = 1;
        body.param_count = 1;
        body.local_count = 1;
        body.func_sig = Some(("i32".into(), "i32".into()));
        primitives::base64::emit_code_is_allowed(&mut body, 0, url, 0);
        assert_eq!(body.local_count, 1);
        assert!(body.imports.is_empty());
        body.emit_op(Op::RETURN, 0);
        let native = vybe_platform_wasm::write_wasm(&[body.clone()]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&native).unwrap();

        let mut owner = Chunk::new("predicate-owner");
        primitives::functions::emit_ref_func(&mut owner, 1, 0, 0);
        owner.emit_op(Op::RETURN, 0);
        let mut chunks = vec![owner, body];
        primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
        primitives::globals::declare_free_globals(&mut chunks);
        primitives::globals::normalize_global_table(&mut chunks);
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        let function = vm.run(chunks).unwrap();
        for cp in (0..65536).chain([
            i32::MIN, i32::MAX, -1, -32, -65, 65536, 0x10061, 0x7fffff61,
        ]) {
            let expected = matches!(cp, 65..=90 | 97..=122 | 48..=57 | 61)
                || if url { matches!(cp, 45 | 95) } else { matches!(cp, 43 | 47) };
            let actual = vm.invoke(&function, &[Value::I32(cp)]).unwrap();
            assert_eq!(actual.as_i32(), i32::from(expected), "url={url}, code={cp}");
            assert!(vm.last_exception.is_none());
        }
    }
}
