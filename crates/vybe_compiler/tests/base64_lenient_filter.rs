use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, Value, VM};

fn module() -> Vec<Chunk> {
    let mut root = Chunk::new("owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut filter = Chunk::new("filter");
    filter.arity = 1;
    filter.param_count = 1;
    filter.local_count = 1;
    filter.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut chunks = vec![root, filter];
    primitives::base64::emit_filter_lenient(&mut chunks, 1, 0);
    chunks[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
    primitives::globals::declare_free_globals(&mut chunks);
    primitives::globals::normalize_global_table(&mut chunks);
    chunks
}

#[test]
fn base64_filter_retains_only_standard_ascii_and_traps_non_strings() {
    let chunks = module();
    let native = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&native).unwrap();
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let function = vm.run(chunks).unwrap();
    for unit in 0..=65535 {
        if let Some(ch) = char::from_u32(unit) {
            let keep = ch.is_ascii_alphanumeric() || matches!(ch, '+' | '/' | '=');
            for text in [ch.to_string(), format!("A{ch}B")] {
                let expected = if keep { text.clone() } else if text.starts_with('A') && text.len() > 1 { "AB".into() } else { String::new() };
                let result = vm.invoke(&function, &[Value::String(text.into())]).unwrap();
                assert_eq!(result.to_string(), expected, "unit={unit:04x}");
            }
        }
    }
    for input in ["", "A!a-0_9+/=\n", "é🙂東京"] {
        let expected: String = input.chars().filter(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '/' | '=')).collect();
        assert_eq!(vm.invoke(&function, &[Value::String(input.into())]).unwrap().to_string(), expected);
    }
    for value in [Value::Null, Value::Undefined, Value::I32(1), Value::Bool(true)] {
        assert!(vm.invoke(&function, &[value]).is_err());
    }
}
