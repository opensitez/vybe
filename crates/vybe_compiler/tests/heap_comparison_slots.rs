#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives::{self, heap, ops};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: usize) -> Vec<Chunk> {
    ops::benchmark_with_comparison_slots(fast, || {
        let mut root = Chunk::new("heap-owner");
        primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut c = Chunk::new("entry");
        c.arity = if kind == 6 { 1 } else { 2 };
        c.param_count = c.arity;
        c.local_count = c.arity as u16;
        let receiver = kind == 4 || kind == 5;
        if receiver {
            root.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
            c.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
        }
        for slot in 0..c.arity {
            c.emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
        }
        let keyed = (2..6).contains(&kind);
        if keyed {
            primitives::functions::emit_ref_func(&mut c, 2, 0, 0);
        }
        let mut callback = Chunk::new("key");
        callback.arity = if receiver { 2 } else { 1 };
        callback.param_count = callback.arity;
        callback.local_count = callback.arity as u16;
        callback.emit_op_u16(Op::LOCAL_GET, if receiver { 1 } else { 0 }, 0);
        callback.emit_string_const("key", 0);
        let get = callback.add_import("ecma:object", "get");
        callback.emit_call(get, 2, 0);
        callback.emit_op(Op::RETURN, 0);
        let mut cs = vec![root, c, callback];
        match kind {
            0 | 2 | 4 => heap::emit_nsmallest(&mut cs, 1, if keyed { 3 } else { 2 }, 0),
            1 | 3 | 5 => heap::emit_nlargest(&mut cs, 1, if keyed { 3 } else { 2 }, 0),
            6 => heap::emit_heapify(&mut cs, 1, 1, 0),
            _ => heap::emit_push(&mut cs, 1, 2, 0),
        }
        cs[1].emit_op(Op::RETURN, 0);
        primitives::bundle::finalize_with_runtime_helpers(&mut cs);
        primitives::globals::declare_free_globals(&mut cs);
        primitives::globals::normalize_global_table(&mut cs);
        let aliases: Vec<_> = cs[0]
            .global_inits
            .iter()
            .filter_map(|g| {
                if let vybe_runtime::chunk::ConstExpr::RefFunc(i) = g.init {
                    Some((g.name.clone(), i))
                } else {
                    None
                }
            })
            .collect();
        for (name, i) in aliases {
            let mut f = Chunk::new(format!("binding-factory:{name}"));
            primitives::functions::emit_ref_func(&mut f, i, 0, 0);
            f.emit_op(Op::RETURN, 0);
            cs.push(f);
        }
        cs
    })
}
fn worker(fast: bool, kind: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind)).unwrap();
    (vm, f)
}
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(v),
    ))
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!() };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(v) = &o.kind else {
        panic!()
    };
    v.clone()
}
fn record(key: i32, id: i32) -> Value {
    let mut o = vybe_runtime::value::Object::new();
    o.properties.insert("key".into(), Value::I32(key));
    o.properties.insert("id".into(), Value::I32(id));
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn invoke(w: &mut (VM, Value), a: &[Value]) -> Value {
    w.0.invoke(&w.1, a).unwrap()
}
#[test]
fn results_stability_and_sizes() {
    let mut checks = 0;
    let mut rows = vec![];
    for kind in 0..8 {
        let mut ws = [worker(false, kind), worker(true, kind)];
        if kind < 6 {
            for len in 0..33 {
                for n in [0, 1, 3, 8, 16, 32, 40] {
                    let keyed = kind >= 2;
                    let largest = kind % 2 == 1;
                    let keys: Vec<i32> = (0..len).map(|i| ((i * 17 + len * 3) % 11) - 5).collect();
                    let input: Vec<Value> = keys
                        .iter()
                        .enumerate()
                        .map(|(i, &k)| {
                            if keyed {
                                record(k, i as i32)
                            } else {
                                Value::I32(k)
                            }
                        })
                        .collect();
                    let mut order: Vec<usize> = (0..len as usize).collect();
                    order.sort_by_key(|&i| if largest { -keys[i] } else { keys[i] });
                    order.truncate((n as usize).min(input.len()));
                    let expected: Vec<Value> = order.iter().map(|&i| input[i].clone()).collect();
                    for w in &mut ws {
                        let a = array(input.clone());
                        let result = invoke(w, &[Value::I32(n), a.clone()]);
                        assert_eq!(values(&result), expected, "kind={kind}, len={len}, n={n}");
                        assert_eq!(values(&a), input);
                        checks += 2;
                    }
                }
            }
        } else {
            for len in 0..33 {
                let mut expected: Vec<Value> = (0..len).rev().map(Value::I32).collect();
                if kind == 6 {
                    for w in &mut ws {
                        let a = array(expected.clone());
                        assert_eq!(invoke(w, &[a.clone()]), Value::Null);
                        let mut sorted = expected.clone();
                        sorted.reverse();
                        assert_eq!(values(&a), sorted);
                        checks += 2;
                    }
                }
                if kind == 7 {
                    expected.reverse();
                    for w in &mut ws {
                        let a = array(expected.clone());
                        invoke(w, &[a.clone(), Value::I32(len / 2)]);
                        let mut wanted = expected.clone();
                        wanted.push(Value::I32(len / 2));
                        wanted.sort_by_key(|v| v.as_f64() as i32);
                        assert_eq!(values(&a), wanted);
                        checks += 1;
                    }
                }
            }
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, kind);
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            if std::env::var_os("VYBE_HEAP_SLOTS_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("heap-slots-{kind}-{fast}"));
            }
            sizes.push(serde_json::json!({"bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"locals":cs.iter().map(|c|c.local_count as usize).sum::<usize>(),"native":bytes.len()}));
        }
        assert!(sizes[1]["bodies"].as_u64() < sizes[0]["bodies"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../HEAP_SLOTS_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (kind, len, n) in [
        (0, 256, 8),
        (0, 256, 256),
        (1, 256, 8),
        (2, 256, 8),
        (6, 64, 0),
        (7, 64, 0),
    ] {
        let input: Vec<Value> = (0..len)
            .map(|i| {
                let k = (i * 17) % 101;
                if kind == 2 {
                    record(k, i)
                } else {
                    Value::I32(k)
                }
            })
            .collect();
        let mut ws = [worker(false, kind), worker(true, kind)];
        let call = |w: &mut (VM, Value)| {
            if kind == 6 {
                invoke(w, &[array(input.clone())])
            } else if kind == 7 {
                invoke(
                    w,
                    &[
                        array((0..len).map(Value::I32).collect()),
                        Value::I32(len / 2),
                    ],
                )
            } else {
                invoke(w, &[Value::I32(n), array(input.clone())])
            }
        };
        for w in &mut ws {
            for _ in 0..8 {
                std::hint::black_box(call(w));
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut calls = 0;
                while t.elapsed() < Duration::from_millis(200) || calls < 8 {
                    std::hint::black_box(call(&mut ws[side]));
                    calls += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / calls as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"kind":kind,"len":len,"n":n,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../HEAP_SLOTS_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}

// Return the actual emitted algorithm rather than its public wrapper, so both
// versions can coexist in one VM without mutable public helper aliases selecting
// the first installed body for both sides.
fn direct_program(fast: bool, kind: usize, callback: bool) -> Vec<Chunk> {
    let mut cs = program(fast, kind);
    let index = if callback {
        2
    } else if kind == 7 {
        1
    } else {
        let name = match kind {
            0 => "__stdlib_nsmallest",
            1 => "__stdlib_nlargest",
            2 => "__stdlib_nsmallest_key",
            6 => "__stdlib_heapify",
            _ => unreachable!(),
        };
        cs.iter().position(|c| c.name == name).unwrap()
    };
    assert_eq!(
        Op::decode(cs[0].read_u16(0), cs[0].read_u16(2)),
        Some(Op::REF_FUNC)
    );
    cs[0].code[4..6].copy_from_slice(&u16::try_from(index).unwrap().to_be_bytes());
    cs
}
#[test]
#[ignore]
fn debug_shared_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (kind, len, n, control) in [
        (0, 256, 8, true),
        (0, 256, 8, false),
        (0, 256, 256, false),
        (1, 256, 8, false),
        (2, 256, 8, false),
        (6, 64, 0, false),
        (7, 64, 0, false),
    ] {
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        let old = vm.run(direct_program(false, kind, false)).unwrap();
        let new = if control {
            old.clone()
        } else {
            vm.run(direct_program(true, kind, false)).unwrap()
        };
        let key = if kind == 2 {
            Some(vm.run(direct_program(true, kind, true)).unwrap())
        } else {
            None
        };
        let fs = [old, new];
        let input: Vec<Value> = (0..len)
            .map(|i| {
                let k = (i * 17) % 101;
                if kind == 2 {
                    record(k, i)
                } else {
                    Value::I32(k)
                }
            })
            .collect();
        let call = |vm: &mut VM, f: &Value| {
            let mut args = if kind == 6 {
                vec![array(input.clone())]
            } else if kind == 7 {
                vec![
                    array((0..len).map(Value::I32).collect()),
                    Value::I32(len / 2),
                ]
            } else {
                vec![Value::I32(n), array(input.clone())]
            };
            if let Some(key) = &key {
                args.push(key.clone());
            }
            vm.invoke(f, &args).unwrap()
        };
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(call(&mut vm, f));
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut count = 0;
                while t.elapsed() < Duration::from_millis(500) || count < 8 {
                    std::hint::black_box(call(&mut vm, &fs[side]));
                    count += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / count as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"kind":kind,"len":len,"n":n,"control":control,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"direct_algorithm_entries":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../HEAP_SLOTS_SHARED_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
