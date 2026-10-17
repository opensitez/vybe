use vybe_compiler::primitives::{bundle, globals, platforms, slices};
use vybe_runtime::{
    Chunk, Op, VM, Value,
    value::{Object, ObjectKind},
};
#[path = "fixtures/slice_positive_before.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn output(n: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("SLICE_POSITIVE_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn program(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("slice-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 4;
    c.param_count = 4;
    c.local_count = 4;
    for i in 0..4 {
        c.emit_op_u16(Op::LOCAL_GET, i, 0);
    }
    let mut cs = vec![root, c];
    if fast {
        slices::emit_splice_assign(&mut cs, 1, 0);
    } else {
        before::emit_splice_assign(&mut cs, 1, 0);
    }
    cs[1].emit_op_u16(Op::LOCAL_GET, 0, 0);
    cs[1].emit_op(Op::RETURN, 0);
    cs[0].imports = cs[1].imports.clone();
    bundle::finalize_with_runtime_helpers(&mut cs);
    globals::declare_free_globals(&mut cs);
    globals::normalize_global_table(&mut cs);
    cs
}
fn native(cs: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        cs,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: false,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&b)
        .unwrap();
    b
}
// Diagnostic factories initialize the original imported helper bindings using
// actual Wasm REF_FUNC closures. Original guest functions remain unchanged.
fn export_native(chunks: &[Chunk], side: &str) {
    let mut cs = chunks.to_vec();
    let bindings: Vec<_> = chunks
        .iter()
        .flat_map(|c| &c.global_inits)
        .filter_map(|init| {
            if let vybe_runtime::chunk::ConstExpr::RefFunc(index) = init.init {
                Some((init.name.clone(), index))
            } else {
                None
            }
        })
        .collect();
    for (name, index) in bindings {
        let mut c = Chunk::new(format!("binding-factory:{name}"));
        c.emit_op_u16(Op::REF_FUNC, index as u16, 0);
        c.emit(0, 0);
        c.emit_op(Op::RETURN, 0);
        cs.push(c);
    }
    native_export::export_chunks(&cs, side);
}
fn worker(cs: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    platforms::register_platforms_all(&mut v);
    let f = v.run(cs).unwrap();
    (v, f)
}
fn arr(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(v)))
}
fn values(x: &Value) -> Vec<Value> {
    let Value::Object(o) = x else { panic!("array") };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
#[test]
fn dynamic_step_surfaces_keep_the_exact_original_emission() {
    for raises in [false, true] {
        for kind in 0..3 {
            let mut a = vec![Chunk::new("original")];
            let mut b = a.clone();
            match kind {
                0 => {
                    before::emit_stepped(&mut a, 0, 17, before::Options::new(raises));
                    slices::emit_stepped(&mut b, 0, 17, slices::Options::new(raises));
                }
                1 => {
                    before::emit_strided_assign(&mut a, 0, 17, before::Options::new(raises));
                    slices::emit_strided_assign(&mut b, 0, 17, slices::Options::new(raises));
                }
                _ => {
                    before::emit_strided_del(&mut a, 0, 17, before::Options::new(raises));
                    slices::emit_strided_del(&mut b, 0, 17, slices::Options::new(raises));
                }
            }
            assert_eq!(a.len(), b.len());
            for (a, b) in a.iter().zip(&b) {
                assert_eq!(a.code, b.code);
                assert_eq!(a.imports, b.imports);
                assert_eq!(a.local_count, b.local_count);
            }
        }
    }
}
#[test]
fn original_bounds_mutation_aliases_and_errors() {
    let old = program(false);
    let new = program(true);
    let bytes = [native(&old), native(&new)];
    assert!(new[1].code.len() < old[1].code.len());
    export_native(&old, "slice-positive-before");
    export_native(&new, "slice-positive-after");
    let sizes = serde_json::json!({"old_body":old[1].code.len(),"new_body":new[1].code.len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count,"old_native":bytes[0].len(),"new_native":bytes[1].len(),"old_chunks":old.len(),"new_chunks":new.len()});
    let (va, fa) = worker(old);
    let (vb, fb) = worker(new);
    let mut vs = [va, vb];
    let mut fs = [fa, fb];
    let bounds = [
        Value::Null,
        Value::I32(-100),
        Value::I32(-1),
        Value::I32(0),
        Value::I32(1),
        Value::I32(3),
        Value::I32(100),
        Value::I32(i32::MIN),
        Value::I32(i32::MAX),
        Value::F64(1.5),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::F64(f64::NEG_INFINITY),
        Value::String("2".into()),
        Value::String("bad".into()),
        Value::Bool(true),
        Value::Undefined,
    ];
    let mut checks = 0;
    let mut rows = vec![];
    for len in [0, 1, 5, 12] {
        for lo in &bounds {
            for hi in &bounds {
                for alias in [false, true] {
                    let mut outcomes = vec![];
                    for side in 0..2 {
                        let input = arr((0..len).map(Value::I32).collect());
                        let replacement = if alias {
                            input.clone()
                        } else {
                            arr(vec![Value::I32(90), Value::I32(91)])
                        };
                        let r = vs[side].invoke(
                            &fs[side],
                            &[input.clone(), lo.clone(), hi.clone(), replacement],
                        );
                        let status = r.map(|_| ()).map_err(|e| {
                            let trap = e.is_trap();
                            (e.message, trap)
                        });
                        let failed = status.is_err();
                        outcomes.push((status, values(&input)));
                        // An unhandled invoke error leaves the VM frame active.
                        // Start a clean worker so subsequent cases are independent.
                        if failed {
                            (vs[side], fs[side]) = worker(program(side == 1));
                        }
                        checks += 1;
                    }
                    assert_eq!(
                        outcomes[0], outcomes[1],
                        "{len} {lo:?} {hi:?} alias={alias}"
                    );
                    if !alias
                        && matches!(lo, Value::Null | Value::I32(_))
                        && matches!(hi, Value::Null | Value::I32(_))
                    {
                        rows.push(serde_json::json!({"length":len,"lo":if matches!(lo,Value::Null){serde_json::Value::Null}else{serde_json::json!(lo.as_i32())},"hi":if matches!(hi,Value::Null){serde_json::Value::Null}else{serde_json::json!(hi.as_i32())},"error":outcomes[1].0.as_ref().err().map(|e|e.0.clone()),"result":outcomes[1].1.iter().map(Value::as_i32).collect::<Vec<_>>()}));
                    }
                }
            }
        }
    }
    output("GATES", &serde_json::json!({"checks":checks,"sizes":sizes}));
    output("JS_CORPUS", &serde_json::json!({"queries":rows}));
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, len, lo, hi, control) in [
        ("defaults", 16, Value::Null, Value::Null, false),
        ("positive", 16, Value::I32(2), Value::I32(6), false),
        ("negative", 16, Value::I32(-6), Value::I32(-2), false),
        ("clamped", 16, Value::I32(-100), Value::I32(100), false),
        ("reversed", 16, Value::I32(10), Value::I32(2), false),
        ("empty", 0, Value::Null, Value::Null, false),
        ("control", 16, Value::I32(2), Value::I32(6), true),
    ] {
        let (va, fa) = worker(program(false));
        let (vb, fb) = worker(program(!control));
        let mut vs = [va, vb];
        let fs = [fa, fb];
        let template: Vec<_> = (0..len).map(Value::I32).collect();
        let inputs = [arr(template.clone()), arr(template.clone())];
        let replacement = arr(vec![Value::I32(90), Value::I32(91)]);
        let args = [
            vec![
                inputs[0].clone(),
                lo.clone(),
                hi.clone(),
                replacement.clone(),
            ],
            vec![inputs[1].clone(), lo, hi, replacement],
        ];
        let mut wall = [vec![], vec![]];
        let mut cpu = [vec![], vec![]];
        let mut calls = vec![];
        for pair in 0..7 {
            let mut secs = [0., 0.];
            let mut cpus = [0., 0.];
            let mut counts = [0usize, 0];
            let mut turn = pair;
            while secs[0] < 0.1 || secs[1] < 0.1 || counts.iter().any(|n| *n < 8) {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    if secs[side] >= 0.1 && counts[side] >= 8 {
                        continue;
                    }
                    let Value::Object(o) = &inputs[side] else {
                        unreachable!()
                    };
                    o.lock().unwrap().kind = ObjectKind::Array(template.clone());
                    let c = thread_cpu_seconds();
                    let t = Instant::now();
                    std::hint::black_box(vs[side].invoke(&fs[side], &args[side]).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                    cpus[side] += thread_cpu_seconds() - c;
                    counts[side] += 1;
                }
                turn += 1;
            }
            for side in 0..2 {
                wall[side].push(secs[side] / counts[side] as f64);
                cpu[side].push(cpus[side] / counts[side] as f64);
            }
            calls.push(counts);
        }
        let mut ratios: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
        ratios.sort_by(f64::total_cmp);
        let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
        cr.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"length":len,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"separate_vms":true,"cases":rows}),
    );
}
#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> f64 {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut t = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut t) }, 0);
    t.sec as f64 + t.nsec as f64 * 1e-9
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> f64 {
    0.0
}
