#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives::{
    self, collections, csv, exact_integer, heap, instructions::core_wasm, memory_span,
    memory_span_set,
};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: usize, repeats: usize) -> Vec<Chunk> {
    core_wasm::benchmark_with_local_tee(fast, || {
        let mut root = Chunk::new("tee-owner");
        root.memory_min_pages = vec![1];
        root.memory_is_64 = vec![false];
        primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut c = Chunk::new("entry");
        c.arity = match kind {
            1 | 4 => 2,
            3 => 3,
            6 => 0,
            _ => 1,
        };
        c.param_count = c.arity;
        c.local_count = if kind == 0 { 2 } else { c.arity as u16 };
        if kind == 0 {
            for n in 0..repeats {
                c.emit_op_u16(Op::LOCAL_GET, 0, 0);
                core_wasm::local_tee(&mut c, 0, 1);
                if n + 1 < repeats {
                    c.emit_op(Op::DROP, 0);
                }
            }
        } else {
            for slot in 0..c.arity {
                c.emit_op_u16(Op::LOCAL_GET, slot as u16, 0);
            }
        }
        let mut cs = vec![root, c];
        match kind {
            0 => {}
            1 => memory_span::emit_single_byte(&mut cs, 1, false, 0),
            2 => memory_span_set::emit_literal_set(&mut cs, 1, false, b"az", 0),
            3 => csv::emit_parse_document(&mut cs, 1, 0),
            4 => heap::emit_nsmallest(&mut cs, 1, 2, 0),
            5 => exact_integer::emit_normalize_value(&mut cs, 1, true, 0),
            _ => primitives::sorted_collection::emit_sorted_collection_new(&mut cs, 1, 0, false, 0),
        };
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
            f.emit_op_u16(Op::REF_FUNC, i as u16, 0);
            f.emit(0, 0);
            f.emit_op(Op::RETURN, 0);
            cs.push(f);
        }
        cs
    })
}
fn worker(fast: bool, kind: usize, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind, repeats)).unwrap();
    (vm, f)
}
fn object(o: vybe_runtime::value::Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn array(v: Vec<Value>) -> Value {
    object(vybe_runtime::value::Object::new_array(v))
}
fn invoke(w: &mut (VM, Value), args: &[Value]) -> Value {
    w.0.invoke(&w.1, args).unwrap()
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
fn same(a: &Value, b: &Value) {
    match (a, b) {
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        _ => assert_eq!(a, b),
    }
}
fn simple(v: &Value) -> serde_json::Value {
    match v {
        Value::String(s) => serde_json::json!(s.as_ref()),
        Value::I32(n) => serde_json::json!(n),
        Value::Object(_) => serde_json::Value::Array(values(v).iter().map(simple).collect()),
        Value::Bool(b) => serde_json::json!(b),
        Value::Null => serde_json::Value::Null,
        _ => panic!("unexpected scalar"),
    }
}
#[test]
fn live_values_and_real_emitters() {
    let mut checks = 0;
    let mut ws = [worker(false, 0, 1), worker(true, 0, 1)];
    for input in [
        Value::I32(i32::MIN),
        Value::I64(i64::MIN),
        Value::F64(f64::NAN),
        Value::F64(-0.),
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::String("é😀".into()),
        Value::bigint_i64(i64::MAX),
        array(vec![]),
    ] {
        for w in &mut ws {
            same(&invoke(w, &[input.clone()]), &input);
            checks += 1;
        }
    }
    for kind in [1, 2] {
        let mut ws = [worker(false, kind, 1), worker(true, kind, 1)];
        for n in 0..160 {
            let mut data = vec![b'b'; n];
            if n % 3 == 0 && !data.is_empty() {
                data[n / 2] = b'a';
            }
            let expected = data.iter().take_while(|b| !b"az".contains(b)).count();
            data.push(0);
            for w in &mut ws {
                for (i, b) in data.iter().enumerate() {
                    w.0.memory.store_u8(32 + i, *b).unwrap();
                }
                let args = if kind == 1 {
                    vec![Value::I32(32), Value::I32(b'a' as i32)]
                } else {
                    vec![Value::I32(32)]
                };
                assert_eq!(invoke(w, &args).as_f64(), expected as f64);
                checks += 1;
            }
        }
    }
    let mut ws = [worker(false, 3, 1), worker(true, 3, 1)];
    for (s, expected) in [
        ("a,b\nc,d", serde_json::json!([["a", "b"], ["c", "d"]])),
        ("\"a,b\",c", serde_json::json!([["a,b", "c"]])),
        (
            "é漢,😀\nplain,end",
            serde_json::json!([["é漢", "😀"], ["plain", "end"]]),
        ),
    ] {
        for w in &mut ws {
            assert_eq!(
                simple(&invoke(
                    w,
                    &[
                        Value::String(s.into()),
                        Value::String(",".into()),
                        Value::String("\"".into())
                    ]
                )),
                expected
            );
            checks += 1;
        }
    }
    let mut ws = [worker(false, 4, 1), worker(true, 4, 1)];
    for n in [0, 1, 3, 17, 64] {
        let input: Vec<_> = (0..64).rev().map(Value::I32).collect();
        for w in &mut ws {
            assert_eq!(
                values(&invoke(w, &[Value::I32(n), array(input.clone())])),
                (0..n).map(Value::I32).collect::<Vec<_>>()
            );
            checks += 1;
        }
    }
    let mut ws = [worker(false, 5, 1), worker(true, 5, 1)];
    for n in [-17, 0, 17] {
        for w in &mut ws {
            let v = Value::bigint_i64(n);
            assert_eq!(invoke(w, &[v.clone()]), v);
            checks += 1;
        }
    }
    for fast in [false, true] {
        let mut w = worker(fast, 6, 1);
        let result = invoke(&mut w, &[]);
        assert!(values(&result).is_empty());
        let Value::Object(o) = result else {
            unreachable!()
        };
        assert_eq!(
            o.lock()
                .unwrap()
                .properties
                .get(primitives::sorted_collection::SET_COLLECTION_KEY),
            Some(&Value::Bool(true))
        );
        checks += 1;
    }
    let mut rows = vec![];
    for kind in 0..7 {
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, kind, 1);
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
            if std::env::var_os("VYBE_TEE_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("tee-{kind}-{fast}"));
            }
            sizes.push(serde_json::json!({"all_bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"locals":cs.iter().map(|c|c.local_count as usize).sum::<usize>(),"native":bytes.len()}));
        }
        assert!(sizes[1]["all_bodies"].as_u64() < sizes[0]["all_bodies"].as_u64());
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../LOCAL_TEE_GATES.json"),
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
    for (kind, label, args) in [
        (0, "store-load-64", vec![Value::I32(17)]),
        (1, "simd-span-4096", vec![Value::I32(32), Value::I32(97)]),
        (
            3,
            "csv-rows-64",
            vec![
                Value::String("a,b,c,d\n".repeat(64).into()),
                Value::String(",".into()),
                Value::String("\"".into()),
            ],
        ),
        (
            4,
            "heap-smallest-8-of-256",
            vec![
                Value::I32(8),
                array((0..256).rev().map(Value::I32).collect()),
            ],
        ),
    ] {
        let mut ws = [
            worker(false, kind, if kind == 0 { 64 } else { 1 }),
            worker(true, kind, if kind == 0 { 64 } else { 1 }),
        ];
        if kind == 1 {
            for w in &mut ws {
                for i in 0..4096 {
                    w.0.memory.store_u8(32 + i, b'b').unwrap();
                }
                w.0.memory.store_u8(32 + 4096, 0).unwrap();
            }
        }
        for w in &mut ws {
            for _ in 0..8 {
                std::hint::black_box(invoke(w, &args));
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(200) || n < 8 {
                    std::hint::black_box(invoke(&mut ws[side], &args));
                    n += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / n as f64);
            }
        }
        let mut r: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        r.sort_by(f64::total_cmp);
        rows.push(
            serde_json::json!({"case":label,"median_speedup":r[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../LOCAL_TEE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
