use vybe_compiler::primitives::{self, collections};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/scan_scalar_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, kind: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("scan-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("scan");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut cs = vec![root, c];
    match (kind, fast) {
        (0, true) => before::emit_byte_length(&mut cs, 1, 0),
        (0, false) => before::emit_byte_length(&mut cs, 1, 0),
        (_, true) => collections::emit_is_sorted(&mut cs, 1, 0),
        (_, false) => before::emit_is_sorted(&mut cs, 1, 0),
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
        f.emit_op_u16(Op::REF_FUNC, i as u16, 0);
        f.emit(0, 0);
        f.emit_op(Op::RETURN, 0);
        cs.push(f);
    }
    cs
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
fn run(w: &mut (VM, Value), v: &Value) -> Value {
    w.0.invoke(&w.1, &[v.clone()]).unwrap()
}
#[test]
fn values_sizes_native() {
    let mut rows = vec![];
    let mut checks = 0;
    for kind in 0..2 {
        let mut ws = [worker(false, kind), worker(true, kind)];
        if kind == 0 {
            let mut cases = vec!["".to_owned(), "abc\0\n".into(), "é漢😀".into()];
            for cp in [
                0x7f, 0x80, 0x7ff, 0x800, 0xd7ff, 0xe000, 0xffff, 0x10000, 0x10ffff,
            ] {
                cases.push(char::from_u32(cp).unwrap().to_string());
            }
            let mut seed = 17u32;
            for n in 0..192 {
                let mut s = String::new();
                for _ in 0..(n % 65) {
                    seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
                    if let Some(c) = char::from_u32(seed % 0x110000) {
                        s.push(c);
                    }
                }
                cases.push(s);
            }
            for s in cases {
                let input = Value::String(s.clone().into());
                for w in &mut ws {
                    assert_eq!(run(w, &input).as_i32() as usize, s.len());
                    checks += 1;
                }
            }
        } else {
            let mut seed = 23u32;
            for n in 0..192 {
                let mut vals: Vec<i32> = (0..n % 65)
                    .map(|_| {
                        seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
                        (seed % 100) as i32 - 50
                    })
                    .collect();
                if n % 2 == 0 {
                    vals.sort();
                }
                let expected = vals.windows(2).all(|v| v[0] <= v[1]);
                let input = array(vals.into_iter().map(Value::I32).collect());
                for w in &mut ws {
                    assert_eq!(run(w, &input), Value::Bool(expected));
                    checks += 1;
                }
            }
            for vals in [
                vec![],
                vec!["a"],
                vec!["a", "b", "z"],
                vec!["z", "a"],
                vec!["same", "same"],
            ] {
                let expected = vals.windows(2).all(|v| v[0] <= v[1]);
                let input = array(vals.into_iter().map(|s| Value::String(s.into())).collect());
                for w in &mut ws {
                    assert_eq!(run(w, &input), Value::Bool(expected));
                    checks += 1;
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
            if std::env::var_os("VYBE_SCAN_EXPORT").is_some() {
                native_export::export_chunks(&cs, &format!("scan-{kind}-{fast}"));
            }
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        }
        if kind == 0 {
            // Both smaller guest-loop candidates increased native size. The
            // frozen pre-stringref loop serves as this historical control.
            // Current production is checked by byte_length_bulk.
            assert_eq!(program(false, kind)[1].code, program(true, kind)[1].code);
            assert_eq!(sizes[0], sizes[1]);
        } else {
            assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        }
        rows.push(serde_json::json!({"kind":kind,"sizes":sizes}));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SCAN_STACK_GATES.json"),
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
    for (label, kind, input) in [
        (
            "bytes-ascii-4096",
            0,
            Value::String("a".repeat(4096).into()),
        ),
        (
            "bytes-mixed-512",
            0,
            Value::String("aé漢😀".repeat(128).into()),
        ),
        ("bytes-short", 0, Value::String("aé😀".into())),
        ("sorted-512", 1, array((0..512).map(Value::I32).collect())),
        (
            "sorted-first-failure",
            1,
            array((0..512).rev().map(Value::I32).collect()),
        ),
        ("sorted-empty", 1, array(vec![])),
    ] {
        let mut ws = [worker(false, kind), worker(true, kind)];
        for w in &mut ws {
            for _ in 0..8 {
                std::hint::black_box(run(w, &input));
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(200) || n < 8 {
                    std::hint::black_box(run(&mut ws[side], &input));
                    n += 1;
                }
                samples[side].push(t.elapsed().as_secs_f64() / n as f64);
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
    let result = serde_json::json!({"pairs":7,"minimum_sample_ms":200,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../SCAN_STACK_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
