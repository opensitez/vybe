use vybe_compiler::primitives::{self, channels, ops};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/number_lhs_compare_before.rs"]
mod before;
#[path = "fixtures/channel_compare_before.rs"]
mod channel_before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: usize, repeats: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("compare-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = if kind == 0 {
        let mut c = Chunk::new("compare");
        c.arity = 2;
        c.param_count = 2;
        c.local_count = 2;
        for n in 0..repeats {
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
            c.emit_op_u16(Op::LOCAL_GET, 1, 0);
            if fast {
                ops::emit_dyn_lt_number_lhs_into(&mut root, &mut c, 0);
            } else {
                before::emit_lt(&mut root, &mut c, 0);
            }
            if n + 1 < repeats {
                c.emit_op(Op::DROP, 0);
            }
        }
        c.emit_op(Op::RETURN, 0);
        c
    } else {
        match (kind, fast) {
            (1, true) => channels::build_chan_ready_send(&mut root),
            (1, false) => channel_before::build_chan_ready_send(&mut root),
            (2, true) => channels::build_chan_try_send(&mut root),
            (2, false) => channel_before::build_chan_try_send(&mut root),
            (3, true) => channels::build_chan_try_recv(&mut root),
            (3, false) => channel_before::build_chan_try_recv(&mut root),
            (4, true) => channels::build_chan_try_peek(&mut root),
            (4, false) => channel_before::build_chan_try_peek(&mut root),
            (5, true) => channels::build_chan_send(&mut root),
            (5, false) => channel_before::build_chan_send(&mut root),
            (_, true) => channels::build_chan_recv(&mut root),
            (_, false) => channel_before::build_chan_recv(&mut root),
        }
    };
    c.name = "compare".into();
    let mut cs = vec![root, c];
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
}
fn worker(fast: bool, kind: usize, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind, repeats)).unwrap();
    vm.memory.grow(1);
    (vm, f)
}
fn object(o: vybe_runtime::value::Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn array(v: Vec<Value>) -> Value {
    object(vybe_runtime::value::Object::new_array(v))
}
fn channel(cap: i32, closed: bool, values: Vec<Value>) -> Value {
    let mut o = vybe_runtime::value::Object::new();
    for (k, v) in [
        ("capacity", Value::I32(cap)),
        ("closed", Value::Bool(closed)),
        ("queue", array(values)),
        ("__futex", Value::I32(0)),
        ("__zero", Value::I32(0)),
    ] {
        o.properties.insert(k.into(), v);
    }
    object(o)
}
#[test]
fn values_sizes_channels() {
    let mut checks = 0;
    for a in [
        Value::I32(-17),
        Value::I32(0),
        Value::I32(17),
        Value::F64(-0.),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
    ] {
        for b in [
            Value::I32(-1),
            Value::I32(0),
            Value::I32(19),
            Value::F64(0.5),
            Value::F64(f64::NAN),
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::Bool(true),
            Value::String("0".into()),
            Value::bigint_i64(3),
            object(vybe_runtime::value::Object::new()),
        ] {
            let mut results = vec![];
            for fast in [false, true] {
                let (mut vm, f) = worker(fast, 0, 1);
                results.push(vm.invoke(&f, &[a.clone(), b.clone()]));
            }
            assert_eq!(results[0].is_ok(), results[1].is_ok());
            if let (Ok(a), Ok(b)) = (&results[0], &results[1]) {
                assert_eq!(a, b);
            }
            checks += 1;
        }
    }
    for kind in 1..5 {
        for fast in [false, true] {
            for (cap, closed, count, waiters) in [
                (2, false, 0, 0),
                (2, false, 1, 0),
                (2, false, 2, 0),
                (2, true, 1, 0),
                (0, false, 0, 0),
                (0, false, 0, 1),
            ] {
                let (mut vm, f) = worker(fast, kind, 1);
                vm.memory.store_i32(0, count).unwrap();
                vm.memory.store_i32(4, waiters).unwrap();
                let ch = channel(
                    cap,
                    closed,
                    (0..count).map(|n| Value::I32(10 + n)).collect(),
                );
                let args = if kind == 2 {
                    vec![ch, Value::Bool(false)]
                } else {
                    vec![ch]
                };
                let result = vm.invoke(&f, &args).unwrap();
                if kind == 1 {
                    assert_eq!(
                        result,
                        Value::Bool(!closed && (count < cap || (cap == 0 && waiters > 0)))
                    );
                } else if kind == 2 {
                    assert_eq!(
                        result,
                        Value::Bool(!closed && if cap == 0 { waiters > 0 } else { count < cap })
                    );
                } else {
                    let Value::Object(o) = result else {
                        panic!("pair")
                    };
                    let o = o.lock().unwrap();
                    let vybe_runtime::value::ObjectKind::Array(v) = &o.kind else {
                        panic!("pair")
                    };
                    assert_eq!(v[1], Value::Bool(count > 0));
                    if count > 0 {
                        assert_eq!(v[0], Value::I32(10));
                    }
                }
                checks += 1;
            }
        }
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
            let validation =
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&bytes)
                    .map(|_| "valid".to_owned())
                    .unwrap_or_else(|e| e.message().to_owned());
            assert_eq!(validation, "valid");
            if std::env::var_os("VYBE_CHANNEL_COMPARE_EXPORT").is_some() {
                let label = if kind == 0 {
                    format!("channel-compare-{fast}")
                } else {
                    format!("channel-{kind}-{fast}")
                };
                native_export::export_chunks(&cs, &label);
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native_bytes":bytes.len(),"validation":validation}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert_eq!(sizes[0]["validation"], sizes[1]["validation"]);
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../CHANNEL_COMPARE_GATES.json"),
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
    for (kind, label) in [
        (0, "number-lt-64"),
        (1, "ready-send-empty"),
        (1, "ready-send-full"),
    ] {
        let mut ws = [worker(false, kind, 64), worker(true, kind, 64)];
        let args = if kind == 0 {
            vec![Value::I32(17), Value::I32(19)]
        } else {
            let full = label.ends_with("full");
            for w in &mut ws {
                w.0.memory.store_i32(0, if full { 2 } else { 0 }).unwrap();
            }
            vec![channel(
                2,
                false,
                if full {
                    vec![Value::I32(10), Value::I32(11)]
                } else {
                    vec![]
                },
            )]
        };
        for w in &mut ws {
            for _ in 0..8 {
                std::hint::black_box(w.0.invoke(&w.1, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(200) || n < 8 {
                    let function = ws[side].1.clone();
                    std::hint::black_box(ws[side].0.invoke(&function, &args).unwrap());
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
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../CHANNEL_COMPARE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
