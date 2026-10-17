use std::sync::{
    atomic::{AtomicUsize, Ordering},
    Arc,
};
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program(direct: bool, edges: usize, repeats: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("call-cache-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut helper = Chunk::new("zero");
    helper.func_sig = Some((String::new(), "externref".into()));
    helper.local_count = 0;
    helper.emit_i32_const(0, 0);
    helper.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    let zero = c.add_import("test:call-cache", "zero");
    assert_eq!(zero, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_if(0);
    for _ in 0..edges {
        if direct {
            c.emit_direct_defined_call(2, &helper, false, 0).unwrap();
        } else {
            c.emit_call(zero, 0, 0);
        }
        c.emit_op(Op::DROP, 0);
    }
    c.emit_end(0);
    let test = c.add_import("wasm:js-number", "test");
    for _ in 0..repeats {
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        // Keep this fixture on actual CALL dispatch, not local-get fusion.
        c.emit_op(Op::NOP, 0);
        c.emit_call(test, 1, 0);
        c.emit_op(Op::DROP, 0);
    }
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::RETURN, 0);
    let mut cs = vec![root, c, helper];
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn setup() -> (VM, Arc<AtomicUsize>) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let count = Arc::new(AtomicUsize::new(0));
    let capture = count.clone();
    vm.register_host_fn(
        "test:call-cache",
        "zero",
        Box::new(move |_, _| {
            capture.fetch_add(1, Ordering::Relaxed);
            Value::I32(0)
        }),
    );
    (vm, count)
}
#[test]
fn imports_direct_and_prepared() {
    let mut checks = 0;
    for prepared in [false, true] {
        for direct in [false, true] {
            let (mut vm, count) = setup();
            let cs = program(direct, 5, 8);
            assert_eq!(cs[1].code, program(!direct, 5, 8)[1].code);
            let f = if prepared {
                let p = vybe_runtime::link_preparation::PreparedLinkedCode::prepare(
                    cs,
                    8 * 1024 * 1024,
                )
                .ok()
                .unwrap();
                vm.run_linked_prepared(p, vec![]).unwrap()
            } else {
                vm.run(cs).unwrap()
            };
            assert_eq!(
                vm.invoke(&f, &[Value::I32(17), Value::I32(0)]).unwrap(),
                Value::I32(17)
            );
            assert_eq!(count.load(Ordering::Relaxed), 0);
            assert_eq!(
                vm.invoke(&f, &[Value::I32(23), Value::I32(1)]).unwrap(),
                Value::I32(23)
            );
            assert_eq!(count.load(Ordering::Relaxed), if direct { 0 } else { 5 });
            checks += 4;
        }
    }
    let (mut vm, count) = setup();
    let f = vm.run(program(false, 5, 8)).unwrap();
    vm.set_reload_hook(Box::new(|_| Ok(program(true, 5, 8))));
    vm.debug_reload().unwrap();
    assert_eq!(
        vm.invoke(&f, &[Value::I32(31), Value::I32(1)]).unwrap(),
        Value::I32(31)
    );
    assert_eq!(count.load(Ordering::Relaxed), 0);
    vm.set_reload_hook(Box::new(|_| Ok(program(false, 5, 8))));
    vm.debug_reload().unwrap();
    assert_eq!(
        vm.invoke(&f, &[Value::I32(37), Value::I32(1)]).unwrap(),
        Value::I32(37)
    );
    assert_eq!(count.load(Ordering::Relaxed), 5);
    checks += 4;
    for direct in [false, true] {
        let cs = program(direct, 5, 8);
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
        if std::env::var_os("VYBE_CALL_CACHE_EXPORT").is_some() {
            native_export::export_chunks(&cs, &format!("call-cache-{direct}"));
        }
    }
    let result = serde_json::json!({"checks":checks});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../CALL_OPERAND_GATES.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_penalty_pairs() {
    use std::time::{Duration, Instant};
    let mut rows = vec![];
    for (edges, control) in [(5, true), (1, false), (5, false), (64, false)] {
        let (mut vm, _) = setup();
        let a = program(false, edges, 64);
        let b = program(!control, edges, 64);
        assert_eq!(a[1].code, b[1].code);
        let old = vm.run(a).unwrap();
        let new = if control {
            old.clone()
        } else {
            vm.run(b).unwrap()
        };
        let fs = [old, new];
        let args = [Value::I32(17), Value::I32(0)];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(vm.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        for pair in 0..7 {
            for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                let t = Instant::now();
                let mut n = 0;
                while t.elapsed() < Duration::from_millis(500) || n < 8 {
                    std::hint::black_box(vm.invoke(&fs[side], &args).unwrap());
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
        rows.push(serde_json::json!({"edges":edges,"control":control,"map_over_no_map_throughput_ratio":ratios[3],"seconds_per_call":samples}));
    }
    let result = serde_json::json!({"same_vm":true,"identical_guest_code":true,"pairs":7,"minimum_sample_ms":500,"minimum_calls":8,"warmups":8,"cases":rows});
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../CALL_OPERAND_PENALTY.json"),
        serde_json::to_vec_pretty(&result).unwrap(),
    )
    .unwrap();
    eprintln!("{result}");
}

#[test]
fn reset_wide_tail_and_rejection() {
    let (mut vm, count) = setup();
    let base = vm.run(program(false, 5, 8)).unwrap();
    let snapshot = vm.snapshot();
    let appended = vm.run(program(true, 5, 8)).unwrap();
    assert_eq!(
        vm.invoke(&appended, &[Value::I32(1), Value::I32(1)])
            .unwrap(),
        Value::I32(1)
    );
    assert_eq!(count.load(Ordering::Relaxed), 0);
    vm.reset_to(&snapshot);
    let reloaded = vm.run(program(false, 5, 8)).unwrap();
    assert_eq!(
        vm.invoke(&reloaded, &[Value::I32(2), Value::I32(1)])
            .unwrap(),
        Value::I32(2)
    );
    assert_eq!(count.load(Ordering::Relaxed), 5);
    assert_eq!(
        vm.invoke(&base, &[Value::I32(3), Value::I32(0)]).unwrap(),
        Value::I32(3)
    );
    let (mut vm, count) = setup();
    let mut cs = program(true, 2, 1);
    let helper = cs.pop().unwrap();
    while cs.len() <= u16::MAX as usize {
        let mut p = Chunk::new("padding");
        p.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        p.emit_op(Op::RETURN, 0);
        cs.push(p);
    }
    let target = cs.len() as u32;
    cs.push(helper);
    for e in cs[1].direct_defined_calls.values_mut() {
        e.target = target;
    }
    let f = vm.run(cs).unwrap();
    assert_eq!(
        vm.invoke(&f, &[Value::I32(7), Value::I32(1)]).unwrap(),
        Value::I32(7)
    );
    assert_eq!(count.load(Ordering::Relaxed), 0);
    for prepared in [false, true] {
        for direct in [false, true] {
            let mut root = Chunk::new("tail-owner");
            primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
            root.emit_op(Op::RETURN, 0);
            let mut helper = Chunk::new("tail-target");
            helper.arity = 1;
            helper.param_count = 1;
            helper.local_count = 1;
            helper.func_sig = Some(("externref".into(), "externref".into()));
            helper.emit_op_u16(Op::LOCAL_GET, 0, 0);
            helper.emit_op(Op::RETURN, 0);
            let mut c = Chunk::new("tail");
            c.arity = 1;
            c.param_count = 1;
            c.local_count = 1;
            if !direct {
                primitives::functions::emit_ref_func(&mut c, 2, 0, 0);
            }
            c.emit_op_u16(Op::LOCAL_GET, 0, 0);
            if direct {
                c.emit_direct_defined_call(2, &helper, true, 0).unwrap();
            } else {
                c.emit_op(Op::RETURN_CALL, 0);
                c.emit(1, 0);
                c.emit(1, 0);
            }
            let cs = vec![root, c, helper];
            let (mut vm, _) = setup();
            let f = if prepared {
                let p = vybe_runtime::PreparedLinkedCode::prepare(cs, 8 * 1024 * 1024)
                    .ok()
                    .unwrap();
                vm.run_linked_prepared(p, vec![]).unwrap()
            } else {
                vm.run(cs).unwrap()
            };
            assert_eq!(vm.invoke(&f, &[Value::I32(41)]).unwrap(), Value::I32(41));
        }
    }
    for mode in 0..3 {
        let mut cs = program(true, 1, 1);
        let at = *cs[1].direct_defined_calls.keys().next().unwrap();
        match mode {
            0 => cs[1].direct_defined_calls.get_mut(&at).unwrap().target = u32::MAX,
            1 => cs[1].direct_defined_calls.get_mut(&at).unwrap().results = "i32".into(),
            _ => cs[1].code[at + 6] = 1,
        };
        let (mut vm, count) = setup();
        assert!(vm.run(cs).is_err());
        assert_eq!(count.load(Ordering::Relaxed), 0);
    }
}
