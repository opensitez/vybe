#[allow(dead_code)]
#[path = "fixtures/linear_compare_recomputed_addresses.rs"]
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
        reference::emit_linear_string_compare(std::slice::from_mut(&mut c), 0, bounded, 0);
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

#[test]
fn code_size_and_unaffected_paths() {
    for bounded in [false, true] {
        let before = caller(false, bounded);
        let after = caller(true, bounded);
        if let Ok(dir) = std::env::var("VYBE_PRIMITIVE_ADDRESS_EXPORT_DIR") {
            let dir = std::path::Path::new(&dir);
            std::fs::create_dir_all(dir).unwrap();
            for (fast, chunk) in [(false, &before), (true, &after)] {
                std::fs::write(
                    dir.join(format!("native-{bounded}-{fast}.wasm")),
                    vybe_platform_wasm::write_wasm(std::slice::from_ref(chunk)),
                )
                .unwrap();
            }
        }
        assert!(after.code.len() < before.code.len());
        assert!(vybe_platform_wasm::write_wasm(std::slice::from_ref(&after)).len() < vybe_platform_wasm::write_wasm(std::slice::from_ref(&before)).len());
        eprintln!(
            "bounded={bounded} body={}=>{} locals={}=>{} wasm={}=>{}",
            before.code.len(),
            after.code.len(),
            before.local_count,
            after.local_count,
            vybe_platform_wasm::write_wasm(std::slice::from_ref(&before)).len(),
            vybe_platform_wasm::write_wasm(std::slice::from_ref(&after)).len()
        );
    }
    for memory64 in [false, true] {
        let mut before = Chunk::new("unaffected");
        before.memory_is_64 = vec![memory64];
        let mut after = before.clone();
        reference::emit_linear_byte_compare(std::slice::from_mut(&mut before), 0, 0);
        primitives::memory_compare::emit_linear_byte_compare(
            std::slice::from_mut(&mut after),
            0,
            0,
        );
        if memory64 {
            assert_eq!(before.code, after.code);
        } else {
            assert!(after.code.len() < before.code.len());
        }
        assert_eq!(before.local_count, after.local_count);
    }
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for bounded in [false, true] {
        for length in [0usize, 8, 256, 4096] {
            let mut workers = [worker(false, bounded), worker(true, bounded)];
            for (vm, _) in &mut workers {
                vm.memory.with_buffer_mut(|m| {
                    m.fill(0x91);
                    m[length] = 0;
                    m[16384 + length] = 0;
                });
            }
            let mut args = vec![Value::I32(0), Value::I32(16384)];
            if bounded {
                args.push(Value::I32((length + 1) as i32));
            }
            let mut samples = [vec![], vec![]];
            for pair in 0..7 {
                for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let (vm, f) = &mut workers[side];
                    let start = Instant::now();
                    let mut n = 0;
                    while start.elapsed() < Duration::from_millis(100) {
                        assert_eq!(
                            std::hint::black_box(vm.invoke(f, &args).unwrap()).as_f64(),
                            0.0
                        );
                        n += 1;
                    }
                    samples[side].push(start.elapsed().as_secs_f64() / n as f64);
                }
            }
            let mut ratios: Vec<_> = samples[0]
                .iter()
                .zip(&samples[1])
                .map(|(a, b)| a / b)
                .collect();
            ratios.sort_by(f64::total_cmp);
            rows.push(serde_json::json!({"bounded":bounded,"length":length,"seconds_per_call":samples,"median_speedup":ratios[3]}));
        }
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks enabled","pairs":7,"minimum_sample_ms":100,"cases":rows});
    let output =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../ADDRESS_MEASUREMENTS.json");
    std::fs::write(output, serde_json::to_vec_pretty(&result).unwrap()).unwrap();
    println!("{result}");
}

#[test]
fn bounded_tail_counts_preserve_exact_difference() {
    let mut workers=[worker(false,true),worker(true,true)];
    let mut checks=0;
    for nul in [1usize,15,16,17,31,32,37] {
        for difference in [0usize,7,15,16,17,31,32,63] {
            for count in [0usize,1,2,7,8,15,16,17,31,32,33,47,48,49,64,65,79,80] {
                let expected=if difference<=nul && difference<count {if difference==nul {-255.0}else{-110.0}}else{0.0};
                for (vm,f) in &mut workers {
                    vm.memory.with_buffer_mut(|m|{m.fill(0x91);m[nul]=0;m[16384+nul]=0;m[16384+difference]=0xff;});
                    assert_eq!(vm.invoke(f,&[Value::I32(0),Value::I32(16384),Value::I32(count as i32)]).unwrap().as_f64(),expected,"nul={nul} difference={difference} count={count}");
                    checks+=1;
                }
            }
        }
    }
    eprintln!("Checked {checks} explicit bounded-tail values");
}
