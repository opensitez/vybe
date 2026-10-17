#[allow(dead_code)]
#[path = "fixtures/linear_byte_compare_reference.rs"]
mod reference;
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
fn worker(fast: bool) -> (VM, Value) {
    let mut root = Chunk::new("owner");
    root.memory_min_pages = vec![1];
    root.memory_is_64 = vec![false];
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut call = Chunk::new("compare");
    call.arity = 3;
    call.param_count = 3;
    call.local_count = 3;
    for i in 0..3 {
        call.emit_op_u16(Op::LOCAL_GET, i, 0);
    }
    let mut chunks = vec![root, call];
    if fast {
        primitives::memory_compare::emit_linear_byte_compare(&mut chunks, 1, 0);
    } else {
        reference::emit_compare(&mut chunks[1], false, false, true, 0);
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
fn linear_simd_preserves_first_difference_bounds_aliases_and_inputs() {
    let mut workers = [worker(false), worker(true)];
    let mut seed = 0x67a918bdu32;
    for case in 0..256 {
        let count = case % 129;
        let mismatch = if count == 0 { 0 } else { case % count };
        let mut bytes = vec![0u8; count];
        for byte in &mut bytes {
            seed ^= seed << 13;
            seed ^= seed >> 17;
            seed ^= seed << 5;
            *byte = seed as u8;
        }
        let outcomes = workers.each_mut().map(|(vm, f)| {
            let before = vm.memory.with_buffer_mut(|m| {
                m.fill(0);
                m[..count].copy_from_slice(&bytes);
                m[16384..16384 + count].copy_from_slice(&bytes);
                if count > 0 && case % 3 != 0 {
                    m[16384 + mismatch] = m[16384 + mismatch].wrapping_add(1);
                }
                m.to_vec()
            });
            let result = vm
                .invoke(
                    f,
                    &[Value::I32(0), Value::I32(16384), Value::I32(count as i32)],
                )
                .unwrap()
                .as_f64();
            vm.memory.with_buffer_mut(|m| assert_eq!(m, &before[..]));
            result
        });
        assert_eq!(outcomes[0], outcomes[1]);
    }
    for (a, b, n) in [
        (65535, 0, 16),
        (65536, 65536, 0),
        (65535, 0, 2),
        (0, 1, 128),
        (0, 0, 128),
    ] {
        let outcomes = workers.each_mut().map(|(vm, f)| {
            vm.memory.with_buffer_mut(|m| {
                m.fill(0x91);
                m[65535] = if n == 2 { 0x91 } else { 0x90 };
            });
            vm.invoke(f, &[Value::I32(a), Value::I32(b), Value::I32(n)])
                .map(|v| v.as_f64())
                .map_err(|e| e.message)
        });
        assert_eq!(outcomes[0], outcomes[1]);
    }
}
#[test]
fn memory64_scalar_fallback_validates_as_standard_wasm() {
    let mut c = Chunk::new("compare64");
    c.arity = 3;
    c.param_count = 3;
    c.local_count = 3;
    c.memory_min_pages = vec![1];
    c.memory_is_64 = vec![true];
    for i in 0..3 {
        c.emit_op_u16(Op::LOCAL_GET, i, 0);
    }
    let mut chunks = vec![c];
    primitives::memory_compare::emit_linear_byte_compare(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
        .unwrap();
}
