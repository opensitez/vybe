use vybe_compiler::primitives::{self, ops};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/js_compare_redundant_extract.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, op: usize, known: u8, repetitions: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("js-compare-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("js-compare");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    for repeat in 0..repetitions {
        if known & 1 != 0 {
            c.emit_f64_const(3.0, 0);
        } else {
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        if known & 2 != 0 {
            c.emit_f64_const(7.0, 0);
        } else {
            c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        }
        let f = if fast {
            [
                ops::emit_js_lt,
                ops::emit_js_gt,
                ops::emit_js_le,
                ops::emit_js_ge,
            ][op]
        } else {
            [
                before::emit_js_lt,
                before::emit_js_gt,
                before::emit_js_le,
                before::emit_js_ge,
            ][op]
        };
        f(&mut c, 0, known & 1 != 0, known & 2 != 0);
        if repeat + 1 != repetitions {
            c.emit_op(Op::DROP, 0);
        }
    }
    let mut cs = vec![root, c];
    cs[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    let bindings: Vec<_> = cs[0]
        .global_inits
        .iter()
        .filter_map(|init| {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = init.init {
                Some((init.name.clone(), index))
            } else {
                None
            }
        })
        .collect();
    for (name, index) in bindings {
        let mut f = Chunk::new(format!("binding-factory:{name}"));
        f.emit_op_u16(Op::REF_FUNC, index as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
}
fn worker(fast: bool, op: usize, known: u8, repetitions: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, op, known, repetitions)).unwrap();
    (vm, f)
}
#[test]
fn comparisons_preserve_numeric_and_mixed_operands() {
    let values = vec![
        Value::I32(-17),
        Value::I32(0),
        Value::I32(17),
        Value::F64(-0.0),
        Value::F64(0.25),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::F64(f64::NEG_INFINITY),
        Value::F32(1.25),
        Value::I64(9007199254740993),
        Value::String("".into()),
        Value::String("3".into()),
        Value::String("x".into()),
        Value::String("é😀".into()),
        Value::Bool(true),
        Value::Null,
        Value::Undefined,
        Value::bigint_i64(2),
        Value::bigint_i64(7),
    ];
    let mut checks = 0;
    for op in 0..4 {
        for known in 0..4 {
            let mut workers = [worker(false, op, known, 1), worker(true, op, known, 1)];
            for a in &values {
                for b in &values {
                    let out = workers.each_mut().map(|(vm, f)| {
                        vm.invoke(f, &[a.clone(), b.clone()]).map_err(|e| e.message)
                    });
                    assert_eq!(out[0], out[1], "{op} {known} {a:?} {b:?}");
                    let left = if known & 1 != 0 {
                        Some(3.0)
                    } else {
                        match a {
                            Value::I32(n) => Some(*n as f64),
                            Value::F64(n) => Some(*n),
                            _ => None,
                        }
                    };
                    let right = if known & 2 != 0 {
                        Some(7.0)
                    } else {
                        match b {
                            Value::I32(n) => Some(*n as f64),
                            Value::F64(n) => Some(*n),
                            _ => None,
                        }
                    };
                    if let (Some(a), Some(b)) = (left, right) {
                        let expected = [a < b, a > b, a <= b, a >= b][op];
                        assert_eq!(out[1], Ok(Value::I32(i32::from(expected))));
                    }
                    checks += 1;
                }
            }
        }
    }
    eprintln!("Verified {checks} JS relational operator/constant-proof/type cases");
}
#[test]
fn size_and_native_exports() {
    let mut rows = vec![];
    for op in 0..4 {
        for known in 0..4 {
            let mut sizes = vec![];
            for fast in [false, true] {
                let cs = program(fast, op, known, 1);
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
                if std::env::var_os("VYBE_PRIMITIVE_JS_COMPARE_EXPORT").is_some() {
                    native_export::export_chunks(
                        &cs,
                        &format!("js-compare-{op}-{known}-{}", u8::from(fast)),
                    );
                }
                sizes.push(serde_json::json!({"body":cs[1].code.len(),"native":bytes.len(),"locals":cs[1].local_count}));
            }
            if known == 3 {
                assert_eq!(sizes[0], sizes[1]);
            } else {
                assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
            }
            rows.push(serde_json::json!({"operator":op,"known":known,"sizes":sizes}));
        }
    }
    if std::env::var_os("VYBE_PRIMITIVE_JS_COMPARE_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../JS_COMPARE_GATES.json"),
            serde_json::to_vec_pretty(&rows).unwrap(),
        )
        .unwrap();
    }
    eprintln!("{rows:?}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (label, known, a, b) in [
        ("mixed-numbers", 0, Value::I32(13), Value::F64(17.0)),
        ("floats", 0, Value::F64(13.0), Value::F64(17.0)),
        ("integers", 0, Value::I32(13), Value::I32(17)),
        ("left-constant", 1, Value::I32(13), Value::F64(17.0)),
        ("right-constant", 2, Value::I32(13), Value::F64(17.0)),
        ("constant-control", 3, Value::I32(13), Value::F64(17.0)),
        (
            "strings-control",
            0,
            Value::String("abc".into()),
            Value::String("def".into()),
        ),
        (
            "mixed-control",
            0,
            Value::I32(13),
            Value::String("17".into()),
        ),
    ] {
        let mut workers = [worker(false, 0, known, 64), worker(true, 0, known, 64)];
        let args = [a, b];
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
        rows.push(
            serde_json::json!({"case":label,"median_speedup":ratios[3],"seconds_per_call":samples}),
        );
    }
    let result = serde_json::json!({"profile":"opt0 assertions and overflow checks","pairs":7,"minimum_sample_ms":100,"minimum_calls_per_sample":8,"comparisons_per_call":64,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../JS_COMPARE_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
