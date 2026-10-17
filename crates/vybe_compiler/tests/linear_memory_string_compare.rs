#[allow(dead_code)]
#[path = "fixtures/linear_byte_compare_reference.rs"]
mod reference;
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
fn caller(fast: bool, bounded: bool) -> Chunk {
    let mut c = Chunk::new("compare");
    c.arity = if bounded { 3 } else { 2 };
    c.param_count = c.arity;
    c.local_count = c.arity as u16;
    c.memory_min_pages = vec![1];
    c.memory_is_64 = vec![false];
    for i in 0..c.arity {
        c.emit_op_u16(Op::LOCAL_GET, i as u16, 0);
    }
    if fast {
        primitives::memory_compare::emit_linear_string_compare(
            std::slice::from_mut(&mut c),
            0,
            bounded,
            0,
        );
    } else {
        reference::emit_compare(&mut c, false, true, bounded, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c
}
fn worker(fast: bool, bounded: bool) -> (VM, Value) {
    let mut root = Chunk::new("owner");
    root.memory_min_pages = vec![1];
    root.memory_is_64 = vec![false];
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut chunks = vec![root, caller(fast, bounded)];
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
fn nul_difference_order_bounds_and_page_traps() {
    let mut cases = 0;
    for bounded in [false, true] {
        let mut workers = [worker(false, bounded), worker(true, bounded)];
        for addr in [0usize, 1, 7, 15, 65503, 65520, 65535] {
            for nul in [0usize, 1, 7, 15, 16, 31, 32, 63] {
                for difference in [
                    None,
                    Some(0usize),
                    Some(7),
                    Some(15),
                    Some(16),
                    Some(31),
                    Some(64),
                ] {
                    let results = workers.each_mut().map(|(vm, f)| {
                        vm.memory.with_buffer_mut(|m| {
                            m.fill(0x91);
                            if addr + nul < m.len() {
                                m[addr + nul] = 0;
                            }
                            m[16384 + nul] = 0;
                            if let Some(d) = difference {
                                m[16384 + d] = 0xff;
                            }
                        });
                        let mut args = vec![Value::I32(addr as i32), Value::I32(16384)];
                        if bounded {
                            args.push(Value::I32(48));
                        }
                        let before = vm.memory.with_buffer_mut(|m| m.to_vec());
                        let forward = vm
                            .invoke(f, &args)
                            .map(|v| v.as_f64())
                            .map_err(|e| e.message);
                        args.swap(0, 1);
                        let reverse = vm
                            .invoke(f, &args)
                            .map(|v| v.as_f64())
                            .map_err(|e| e.message);
                        vm.memory.with_buffer_mut(|m| assert_eq!(m, &before[..]));
                        (forward, reverse)
                    });
                    assert_eq!(
                        results[0], results[1],
                        "bounded={bounded} addr={addr} nul={nul} difference={difference:?}"
                    );
                    cases += 2;
                }
            }
        }
        for (addr, count) in [(65536, 0.0), (65535, 0.5), (0, 15.5), (0, 16.5), (0, -1.0)] {
            if !bounded {
                continue;
            }
            let results = workers.each_mut().map(|(vm, f)| {
                vm.memory.with_buffer_mut(|m| {
                    m.fill(0x91);
                    m[65535] = 0;
                });
                vm.invoke(f, &[Value::I32(addr), Value::I32(16384), Value::F64(count)])
                    .map(|v| v.as_f64())
                    .map_err(|e| e.message)
            });
            assert_eq!(results[0], results[1]);
        }
    }
    eprintln!("Matched {cases} NUL/difference/page cases plus fractional bounds");
}
#[test]
fn memory64_scalar_strings_validate_and_preserve_termination() {
    for bounded in [false, true] {
        let mut c = Chunk::new("compare64");
        c.arity = if bounded { 3 } else { 2 };
        c.param_count = c.arity;
        c.local_count = c.arity as u16;
        c.memory_min_pages = vec![1];
        c.memory_is_64 = vec![true];
        for i in 0..c.arity {
            c.emit_op_u16(Op::LOCAL_GET, i as u16, 0);
        }
        primitives::memory_compare::emit_linear_string_compare(
            std::slice::from_mut(&mut c),
            0,
            bounded,
            0,
        );
        c.emit_op(Op::RETURN, 0);
        let mut root = Chunk::new("owner64");
        root.memory_min_pages = vec![1];
        root.memory_is_64 = vec![true];
        primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut chunks = vec![root, c];
        primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
        primitives::globals::declare_free_globals(&mut chunks);
        primitives::globals::normalize_global_table(&mut chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        let f = vm.run(chunks).unwrap();
        for (nul, difference, expected) in [
            (31, 63, 0.0),
            (31, 16, -110.0),
            (0, 63, 0.0),
            (16, 16, -255.0),
        ] {
            vm.memory.with_buffer_mut(|m| {
                m.fill(0x91);
                m[nul] = 0;
                m[16384 + nul] = 0;
                m[16384 + difference] = 0xff;
            });
            let mut args = vec![Value::I32(0), Value::I32(16384)];
            if bounded {
                args.push(Value::I32(64));
            }
            assert_eq!(vm.invoke(&f, &args).unwrap().as_f64(), expected);
        }
    }
}
