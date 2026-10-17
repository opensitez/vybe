use vybe_compiler::primitives::{self, collections, strings};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/inverted_format_before.rs"]
mod format_before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
#[path = "fixtures/inverted_reverse_before.rs"]
mod reverse_before;
fn op(i: usize) -> Op {
    [
        Op::I32_EQ,
        Op::I32_NE,
        Op::I32_LT_S,
        Op::I32_LE_S,
        Op::I32_GT_S,
        Op::I32_GE_S,
        Op::I32_LT_U,
        Op::I32_LE_U,
        Op::I32_GT_U,
        Op::I32_GE_U,
    ][i]
}
fn inverse(i: usize) -> Op {
    [
        Op::I32_NE,
        Op::I32_EQ,
        Op::I32_GE_S,
        Op::I32_GT_S,
        Op::I32_LE_S,
        Op::I32_LT_S,
        Op::I32_GE_U,
        Op::I32_GT_U,
        Op::I32_LE_U,
        Op::I32_LT_U,
    ][i]
}
fn expected(i: usize, a: i32, b: i32) -> bool {
    match i {
        0 => a != b,
        1 => a == b,
        2 => a >= b,
        3 => a > b,
        4 => a <= b,
        5 => a < b,
        6 => (a as u32) >= (b as u32),
        7 => (a as u32) > (b as u32),
        8 => (a as u32) <= (b as u32),
        _ => (a as u32) < (b as u32),
    }
}
fn program(fast: bool, kind: usize, operator: usize, repeats: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("inversion-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let c = if kind == 1 {
        if fast {
            strings::build_format_map(&mut root)
        } else {
            format_before::build_format_map(&mut root)
        }
    } else {
        let mut c = Chunk::new("entry");
        c.arity = if kind == 0 { 2 } else { 1 };
        c.param_count = c.arity;
        c.local_count = c.arity as u16;
        if kind == 0 {
            for n in 0..repeats {
                c.emit_op_u16(Op::LOCAL_GET, 0, 0);
                c.emit_op_u16(Op::LOCAL_GET, 1, 0);
                c.emit_op(
                    if fast {
                        inverse(operator)
                    } else {
                        op(operator)
                    },
                    0,
                );
                if !fast {
                    c.emit_op(Op::I32_EQZ, 0);
                }
                if n + 1 < repeats {
                    c.emit_op(Op::DROP, 0);
                }
            }
            c.emit_op(Op::RETURN, 0);
        }
        c
    };
    let mut cs = vec![root, c];
    if kind == 2 {
        cs[1].emit_op_u16(Op::LOCAL_GET, 0, 0);
        collections::emit_reversed(&mut cs, 1, 0);
        cs[1].emit_op(Op::RETURN, 0);
    }
    primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    if kind == 2 && !fast {
        let target = cs
            .iter()
            .position(|c| c.name.starts_with("__stdlib_reverse_copy"))
            .unwrap();
        let mut body = primitives::functions::create_function_chunk(&cs[target].name, 1);
        body.local_count = 1;
        let mut old = vec![body];
        reverse_before::emit_reverse_copy_body(&mut old, 0, 0);
        old[0].emit_op(Op::RETURN, 0);
        cs[target] = old.remove(0);
        primitives::bundle::finalize_with_runtime_helpers(&mut cs);
    }
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
    cs[1].name = "entry".into();
    cs
}
fn worker(fast: bool, kind: usize, op: usize, repeats: usize) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, kind, op, repeats)).unwrap();
    (vm, f)
}
fn object(o: vybe_runtime::value::Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn array(v: Vec<Value>) -> Value {
    object(vybe_runtime::value::Object::new_array(v))
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
fn invoke(w: &mut (VM, Value), args: &[Value]) -> Value {
    w.0.invoke(&w.1, args).unwrap()
}
#[test]
fn oracle_and_actual_builders() {
    let mut checks = 0;
    let inputs = vec![
        Value::I32(i32::MIN),
        Value::I32(i32::MIN + 1),
        Value::I32(-17),
        Value::I32(-1),
        Value::I32(0),
        Value::I32(1),
        Value::I32(17),
        Value::I32(i32::MAX - 1),
        Value::I32(i32::MAX),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::F64(-0.5),
        Value::Bool(false),
        Value::Null,
        Value::String("3".into()),
    ];
    let mut rows = vec![];
    for operator in 0..10 {
        let mut ws = [worker(false, 0, operator, 1), worker(true, 0, operator, 1)];
        for a in &inputs {
            for b in &inputs {
                let expected = Value::I32(i32::from(expected(operator, a.as_i32(), b.as_i32())));
                for w in &mut ws {
                    assert_eq!(invoke(w, &[a.clone(), b.clone()]), expected);
                    checks += 1;
                }
            }
        }
        for fast in [false, true] {
            let cs = program(fast, 0, operator, 1);
            validate_export(&cs, &format!("inversion-{operator}-{fast}"));
        }
        rows.push(size_pair(0, operator));
    }
    let mut mapping = vybe_runtime::value::Object::new();
    mapping
        .properties
        .insert("x".into(), Value::String("VALUE".into()));
    let mapping = object(mapping);
    let mut ws = [worker(false, 1, 0, 1), worker(true, 1, 0, 1)];
    for s in [
        "",
        "literal",
        "{x}",
        "a{x}b",
        "{unknown}",
        "{unterminated",
        "é😀{x}",
    ] {
        let args = [Value::String(s.into()), mapping.clone()];
        let a = invoke(&mut ws[0], &args);
        let b = invoke(&mut ws[1], &args);
        assert_eq!(a, b);
        checks += 1;
    }
    let mut ws = [worker(false, 2, 0, 1), worker(true, 2, 0, 1)];
    for n in 0..96 {
        let v: Vec<_> = (0..n).map(Value::I32).collect();
        let input = array(v.clone());
        for w in &mut ws {
            let mut expected = v.clone();
            expected.reverse();
            assert_eq!(values(&invoke(w, &[input.clone()])), expected);
            checks += 1;
        }
    }
    for kind in [1, 2] {
        for fast in [false, true] {
            let cs = program(fast, kind, 0, 1);
            validate_export(&cs, &format!("inversion-integration-{kind}-{fast}"));
        }
        rows.push(size_pair(kind, 0));
    }
    let result = serde_json::json!({"checks":checks,"rows":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../INVERTED_I32_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
fn validate_export(cs: &[Chunk], label: &str) {
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
    if std::env::var_os("VYBE_INVERSION_EXPORT").is_some() {
        native_export::export_chunks(cs, label);
    }
}
fn size_pair(kind: usize, operator: usize) -> serde_json::Value {
    let sizes:Vec<_>=[false,true].into_iter().map(|fast|{let cs=program(fast,kind,operator,1);let bytes=vybe_platform_wasm::write_wasm_with_options(&cs,vybe_platform_wasm::WasmWriteOptions{include_name_section:false,include_vybe_metadata:false});serde_json::json!({"all_bodies":cs.iter().map(|c|c.code.len()).sum::<usize>(),"locals":cs.iter().map(|c|c.local_count as usize).sum::<usize>(),"native":bytes.len()})}).collect();
    assert!(sizes[1]["all_bodies"].as_u64() < sizes[0]["all_bodies"].as_u64());
    serde_json::json!({"kind":kind,"operator":operator,"sizes":sizes})
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (kind, label, args) in [
        (0, "integer-ge-64", vec![Value::I32(17), Value::I32(19)]),
        (
            1,
            "format-literal-256",
            vec![
                Value::String("abcd".repeat(64).into()),
                object(vybe_runtime::value::Object::new()),
            ],
        ),
        (
            2,
            "reverse-4096",
            vec![array((0..4096).map(Value::I32).collect())],
        ),
    ] {
        let mut ws = [
            worker(false, kind, 2, if kind == 0 { 64 } else { 1 }),
            worker(true, kind, 2, if kind == 0 { 64 } else { 1 }),
        ];
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
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../INVERTED_I32_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
