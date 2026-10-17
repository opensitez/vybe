use std::path::Path;
use vybe_ast::{BitLane, IntOp};
use vybe_runtime::{Chunk, Op, VM};

fn runtime() -> VM {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    vm
}

#[test]
fn stdout_buffer_length_uses_a_static_integer_comparison() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/stdout_buffer_length_helper.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    let chunk = chunks.iter().find(|chunk| chunk.name == "__c_stdout_has_buffered").unwrap();
    let dump = vybe_runtime::debug::disassemble(chunk);
    assert!(dump.contains("i32.gt_u"), "{dump}");
    assert!(!dump.contains("dyngt"), "{dump}");
    assert!(!dump.contains("__vybe_slot_"), "{dump}");
    let mut vm = runtime();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn native_word_cast_does_not_recoerce_the_result() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/native_word_cast.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path).unwrap().compile().unwrap();
    for name in ["cast_or", "cast_and"] {
        let dump = vybe_runtime::debug::disassemble(
            chunks.iter().find(|chunk| chunk.name == name).unwrap(),
        );
        assert!(!dump.contains("f64.div"), "{name}: {dump}");
        assert!(!dump.contains("ecma:number:Number"), "{name}: {dump}");
    }
    let mut vm = runtime();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn wide_c_comparisons_are_static_without_truncating_to_i32() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/static_wide_comparisons.c");
    vybe_language_c::register();
    let chunks = vybe_compiler::projects::load(&path)
        .unwrap()
        .compile()
        .unwrap();
    for (name, opcode) in [
        ("equal_sizes", "i64.eq"),
        ("nonzero_size", "i64.ne"),
        ("last_item", "i64.eq"),
        ("last_read_item", "i64.eq"),
        ("signed_less", "i64.gt_s"),
        ("promoted_less", "i64.gt_s"),
        ("ordered_sizes", "i64.gt_u"),
    ] {
        let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(
            !dump.contains("dyneq"),
            "dynamic equality in {name}: {dump}"
        );
        assert!(dump.contains(opcode), "missing {opcode} in {name}: {dump}");
    }
    let mut vm = runtime();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn shared_comparisons_validate_as_wasm_and_keep_both_lanes() {
    for (lane, left, right, operation, expected) in [
        (BitLane::W32, 4294967295.0, 0.0, IntOp::GtU, 1),
        (BitLane::W32, 4294967295.0, 0.0, IntOp::LtS, 1),
        (BitLane::W32, 4294967295.0, -1.0, IntOp::Eq, 1),
        (BitLane::W32, 1.0, 1.0, IntOp::Ne, 0),
        (BitLane::W64, 1099511627776.0, 1099511627777.0, IntOp::Eq, 0),
        (BitLane::W64, 4294967296.0, 0.0, IntOp::Ne, 1),
        (BitLane::W64, -4294967296.0, 0.0, IntOp::LtS, 1),
        (BitLane::W64, -4294967296.0, 0.0, IntOp::LeS, 1),
        (BitLane::W64, 4294967296.0, 0.0, IntOp::GtS, 1),
        (BitLane::W64, 4294967296.0, 0.0, IntOp::GeS, 1),
        (BitLane::W64, 0.0, 4294967296.0, IntOp::LtU, 1),
        (BitLane::W64, 0.0, 4294967296.0, IntOp::LeU, 1),
        (BitLane::W64, 4294967296.0, 0.0, IntOp::GeU, 1),
        (BitLane::W64, 9223372036854775808.0, 1.0, IntOp::GtU, 1),
        (BitLane::W64, 9223372036854775808.0, 1.0, IntOp::LtS, 1),
        (
            BitLane::W64,
            9223372036854775808.0,
            -9223372036854775808.0,
            IntOp::Eq,
            1,
        ),
    ] {
        let mut chunk = Chunk::new("<script>");
        chunk.emit_f64_const(left, 0);
        chunk.emit_f64_const(right, 0);
        vybe_compiler::primitives::bits::emit_integer_binary(&mut chunk, lane, operation, 0);
        chunk.emit_op(Op::RETURN, 0);
        let chunks = vec![chunk];
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
        let result = runtime().run(chunks).unwrap();
        assert_eq!(result.as_i32(), expected, "{lane:?} {operation:?}");
        let reloaded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
        let mut vm = runtime();
        vm.run(reloaded).unwrap();
        // The decoded startup wrapper intentionally discards _start's result.
        // Invoke the exported source function to check its actual predicate.
        let mut function = vybe_runtime::value::Object::new();
        function.kind = vybe_runtime::value::ObjectKind::Function(vybe_runtime::value::Function {
            name: None,
            arity: 0,
            chunk_index: 1,
            upvalues: Vec::new(),
        });
        let function = vybe_runtime::value::Value::Object(std::sync::Arc::new(
            std::sync::Mutex::new(function),
        ));
        let result = vm.invoke(&function, &[]).unwrap();
        assert_eq!(result.as_i32(), expected, "reloaded {lane:?} {operation:?}");
    }
}
