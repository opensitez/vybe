use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{self, collections};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/index_func_walk_after_match.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(fast: bool, callback: bool) -> Vec<Chunk> {
    let mut owner = Chunk::new("index-owner");
    primitives::functions::emit_ref_func(&mut owner, 1, 0, 0);
    owner.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("index");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    for slot in 0..2 {
        c.emit_op_u16(Op::LOCAL_GET, slot, 0);
    }
    let mut chunks = vec![owner, c];
    if fast {
        collections::emit_index_func(&mut chunks, 1, 0);
    } else {
        before::emit_index_func(&mut chunks, 1, 0);
    }
    chunks[1].emit_op(Op::RETURN, 0);
    if callback {
        let mut p = Chunk::new("predicate");
        p.arity = 1;
        p.param_count = 1;
        p.local_count = 1;
        p.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let i = p.add_import("test:index", "predicate");
        p.emit_call(i, 1, 0);
        p.emit_op(Op::RETURN, 0);
        chunks.push(p);
        let mut factory = Chunk::new("predicate-owner");
        primitives::functions::emit_ref_func(&mut factory, 2, 0, 0);
        factory.emit_op(Op::RETURN, 0);
        chunks.push(factory);
    }
    primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
    primitives::globals::declare_free_globals(&mut chunks);
    primitives::globals::normalize_global_table(&mut chunks);
    chunks
}
fn worker(
    fast: bool,
    predicate: Box<dyn Fn(&mut vybe_runtime::HostContext, &[Value]) -> Value + Send + Sync>,
) -> (VM, Value, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm.register_host_fn("test:index", "predicate", predicate);
    let f = vm.run(program(fast, true)).unwrap();
    let mut object = Object::new();
    object.kind = ObjectKind::Function(Function {
        name: Some("predicate".into()),
        arity: 1,
        chunk_index: 2,
        upvalues: vec![],
    });
    (vm, f, Value::Object(vybe_runtime::heap::alloc(object)))
}
fn array(n: usize) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(
        (0..n).map(|i| Value::I32(i as i32)).collect(),
    )))
}
#[test]
fn early_exit_preserves_values_callback_order_and_errors() {
    let mut checks = 0;
    for length in [0usize, 1, 2, 8, 128] {
        for hit in [None, Some(0usize), Some(length / 2), length.checked_sub(1)] {
            for truth in [
                Value::Bool(true),
                Value::I32(1),
                Value::String("yes".into()),
                Value::Object(vybe_runtime::heap::alloc(Object::new())),
            ] {
                let mut outcomes = vec![];
                for fast in [false, true] {
                    let trace = Arc::new(Mutex::new(vec![]));
                    let log = trace.clone();
                    let answer = truth.clone();
                    let (mut vm, f, p) = worker(
                        fast,
                        Box::new(move |_, a| {
                            let value = a[0].as_f64() as usize;
                            log.lock().unwrap().push(value);
                            if Some(value) == hit {
                                answer.clone()
                            } else {
                                Value::Bool(false)
                            }
                        }),
                    );
                    let result = vm
                        .invoke(&f, &[array(length), p])
                        .map(|v| v.as_f64())
                        .map_err(|e| e.message);
                    outcomes.push((result, trace.lock().unwrap().clone()));
                }
                assert_eq!(outcomes[0], outcomes[1]);
                let valid = hit.filter(|i| *i < length);
                assert_eq!(outcomes[1].0, Ok(valid.map(|i| i as f64).unwrap_or(-1.0)));
                assert_eq!(outcomes[1].1.len(), valid.map(|i| i + 1).unwrap_or(length));
                checks += 1;
            }
        }
    }
    for throw_at in [0usize, 3, 7] {
        let mut outcomes = vec![];
        for fast in [false, true] {
            let trace = Arc::new(Mutex::new(vec![]));
            let log = trace.clone();
            let (mut vm, f, p) = worker(
                fast,
                Box::new(move |ctx, a| {
                    let value = a[0].as_f64() as usize;
                    log.lock().unwrap().push(value);
                    if value == throw_at {
                        ctx.throw_value(Value::String("predicate-error".into()));
                    }
                    Value::Bool(false)
                }),
            );
            outcomes.push((
                vm.invoke(&f, &[array(8), p])
                    .map(|v| v.as_f64())
                    .map_err(|e| e.message),
                trace.lock().unwrap().clone(),
            ));
        }
        assert_eq!(outcomes[0], outcomes[1]);
        assert!(outcomes[1].0.is_err());
        checks += 1;
    }
    eprintln!("Verified {checks} callback/result/error scenarios");
}
#[test]
fn code_shrinks_and_native_modules_validate() {
    let programs = [program(false, true), program(true, true)];
    let mut sizes = vec![];
    for (side, chunks) in programs.iter().enumerate() {
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: false,
                include_vybe_metadata: false,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        if std::env::var_os("VYBE_PRIMITIVE_INDEX_EXPORT").is_some() {
            native_export::export_chunks(chunks, &format!("index-early-{side}"));
        }
        sizes.push(serde_json::json!({"internal":chunks[1].code.len(),"locals":chunks[1].local_count,"native":bytes.len()}));
    }
    assert!(programs[1][1].code.len() < programs[0][1].code.len());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    eprintln!("index sizes={sizes:?}");
    if std::env::var_os("VYBE_PRIMITIVE_INDEX_EXPORT").is_some() {
        std::fs::write(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../INDEX_GATES.json"),
            serde_json::to_vec_pretty(&sizes).unwrap(),
        )
        .unwrap();
    }
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for hit in [Some(0usize), Some(2048), Some(4095), None] {
        let mut workers = [false, true].map(|fast| {
            worker(
                fast,
                Box::new(move |_, a| Value::Bool(Some(a[0].as_f64() as usize) == hit)),
            )
        });
        let input = array(4096);
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, f, p) = &mut workers[side];
                let args = [input.clone(), p.clone()];
                let start = Instant::now();
                let mut n = 0;
                while start.elapsed() < Duration::from_millis(100) {
                    assert_eq!(
                        std::hint::black_box(vm.invoke(f, &args).unwrap()).as_f64(),
                        hit.map(|i| i as f64).unwrap_or(-1.0)
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
        rows.push(serde_json::json!({"length":4096,"hit":hit,"seconds_per_call":samples,"median_speedup":ratios[3]}));
    }
    let result = serde_json::json!({"profile":"opt0, assertions and overflow checks enabled","pairs":7,"minimum_sample_ms":100,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../INDEX_MEASUREMENTS.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
