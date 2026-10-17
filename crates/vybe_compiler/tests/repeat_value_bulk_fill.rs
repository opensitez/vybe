use vybe_compiler::primitives::{self, collections};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/repeat_value_push_loop.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool) -> Vec<Chunk> {
    let mut owner = Chunk::new("repeat-owner");
    primitives::functions::emit_ref_func(&mut owner, 1, 0, 0);
    owner.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("repeat");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    for i in 0..2 {
        c.emit_op_u16(Op::LOCAL_GET, i, 0);
    }
    let mut cs = vec![owner, c];
    if fast {
        collections::emit_repeat_value(&mut cs, 1, 0);
    } else {
        before::emit_repeat_value(&mut cs, 1, 0);
    }
    cs[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn worker(fast: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast)).unwrap();
    (vm, f)
}
fn items(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
fn repair_baseline_carrier(cs: &mut [Chunk]) {
    let c = &mut cs[1];
    let mut at = 0;
    while at < c.code.len() {
        let op = Op::from_code_at(&c.code, at);
        if op == Op::ARRAY_NEW_FIXED {
            assert_eq!(&c.code[at + 4..at + 8], &[0, 0, 0, 0]);
            assert_eq!(Op::from_code_at(&c.code, at + 8), Op::LOCAL_SET);
            let slot = u16::from_be_bytes([c.code[at + 12], c.code[at + 13]]);
            c.local_names.push(vybe_runtime::chunk::LocalName::new(
                slot,
                format!("<native-growable-array:{at}>"),
                vybe_runtime::chunk::LocalOrigin::Compiler,
            ));
            return;
        }
        at += 4 + op.operand_format().size_in(&c.code, at + 4);
    }
    panic!("allocation");
}
#[test]
fn bulk_repeat_preserves_counts_values_identity_and_freshness() {
    let object = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let values = vec![
        Value::I32(17),
        Value::F64(-0.0),
        Value::Bool(false),
        Value::Null,
        Value::Undefined,
        Value::String("é😀".into()),
        Value::bigint_i64(9),
        object.clone(),
    ];
    let mut checks = 0;
    for value in values {
        let mut workers = [worker(false), worker(true)];
        for count in [-8i32, 0, 1, 2, 17, 255, 4096] {
            let out = workers
                .each_mut()
                .map(|(vm, f)| vm.invoke(f, &[value.clone(), Value::I32(count)]).unwrap());
            let a = items(&out[0]);
            let b = items(&out[1]);
            assert_eq!(a, b);
            assert_eq!(b.len(), count.max(0) as usize);
            for item in &b {
                match (&value, item) {
                    (Value::F64(v), Value::F64(w)) => assert_eq!(v.to_bits(), w.to_bits()),
                    (Value::Object(a), Value::Object(b)) => assert!(std::sync::Arc::ptr_eq(a, b)),
                    _ => assert_eq!(item, &value),
                }
            }
            let (vm, f) = &mut workers[1];
            let other = vm.invoke(f, &[value.clone(), Value::I32(count)]).unwrap();
            let (Value::Object(a), Value::Object(b)) = (&out[1], other) else {
                panic!("array")
            };
            assert!(!std::sync::Arc::ptr_eq(a, &b));
            checks += 1;
        }
    }
    for count in [
        Value::F64(-0.0),
        Value::F64(3.75),
        Value::F64(-3.75),
        Value::F64(f64::NAN),
        Value::Bool(true),
        Value::Null,
        Value::Undefined,
        Value::String("3".into()),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
        Value::bigint_i64(3),
    ] {
        let mut out = vec![];
        for fast in [false, true] {
            let (mut vm, f) = worker(fast);
            out.push(
                vm.invoke(&f, &[Value::I32(7), count.clone()])
                    .map(|v| items(&v))
                    .map_err(|e| e.message),
            );
        }
        assert_eq!(out[0], out[1], "count={count:?}");
        checks += 1;
    }
    eprintln!("Verified {checks} repeat count/value/identity cases");
}
#[test]
fn size_and_native_exports() {
    let mut old = program(false);
    let unrepaired = old.clone();
    let unrepaired_native = vybe_platform_wasm::write_wasm_with_options(&unrepaired,vybe_platform_wasm::WasmWriteOptions{include_name_section:false,include_vybe_metadata:false}).len();
    repair_baseline_carrier(&mut old);
    let new = program(true);
    let programs = [old, new];
    let mut sizes = vec![];
    for (side, cs) in programs.iter().enumerate() {
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            cs,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: false,
                include_vybe_metadata: false,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        if std::env::var_os("VYBE_PRIMITIVE_REPEAT_EXPORT").is_some() {
            native_export::export_chunks(cs, &format!("repeat-{side}"));
        }
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"native":bytes.len(),"locals":cs[1].local_count}));
    }
    if std::env::var_os("VYBE_PRIMITIVE_REPEAT_EXPORT").is_some() {
        native_export::export_chunks(&unrepaired, "repeat-unrepaired");
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../REPEAT_GATES.json"),
            serde_json::to_vec_pretty(&sizes).unwrap(),
        )
        .unwrap();
    }
    assert!(programs[1][1].code.len() < programs[0][1].code.len());
    assert!(programs[1][1].local_count < programs[0][1].local_count);
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    eprintln!("{sizes:?}; unrepaired native bytes={unrepaired_native}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for count in [0i32, 1, 8, 4096] {
        let mut workers = [worker(false), worker(true)];
        let args = [Value::I32(17), Value::I32(count)];
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f) = &mut workers[side];
                let start = Instant::now();
                let mut n = 0;
                while start.elapsed() < Duration::from_millis(100) || n < 8 {
                    std::hint::black_box(vm.invoke(f, &args).unwrap());
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
        rows.push(serde_json::json!({"count":count,"median_speedup":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../REPEAT_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
