use vybe_compiler::primitives::{bundle, globals, platforms, sprintf};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/sprintf_optimization_arguments.rs"]
mod baseline;
fn program(old: bool, count: u8, repeats: usize) -> Vec<Chunk> {
    program_with_identity(old, count, repeats, false)
}
fn program_with_identity(old: bool, count: u8, repeats: usize, identity: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<sprintf-arguments>")];
    for repeat in 0..repeats {
        globals::emit_read(&mut chunks[0], "fmt", 0);
        for i in 0..count {
            globals::emit_read(&mut chunks[0], &format!("arg{i}"), 0);
        }
        if old {
            baseline::emit_sprintf(&mut chunks, 0, count + 1, 0);
        } else {
            sprintf::emit_sprintf(&mut chunks, 0, count + 1, 0);
        }
        if repeat + 1 < repeats {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    // Replace before linking: formatter string imports must not remain in
    // the normalized module namespace after its body is removed.
    if identity {
        // Test the call ABI directly, independently of formatter coercions.
        let helper = chunks
            .iter_mut()
            .find(|c| c.name == "__fmt_sprintf")
            .unwrap();
        *helper = Chunk::new("__fmt_sprintf");
        helper.arity = 2;
        helper.local_count = 2;
        helper.emit_op_u16(Op::LOCAL_GET, 1, 0);
        helper.emit_op(Op::RETURN, 0);
    }
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn execute(chunks: &[Chunk], fmt: &str, args: &[Value]) -> Value {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("fmt", Value::String(fmt.into()));
    for (i, value) in args.iter().enumerate() {
        vm.set_global(&format!("arg{i}"), value.clone());
    }
    vm.run(chunks.to_vec()).unwrap()
}
#[test]
fn arguments_preserve_order_and_values() {
    let cases = [
        ("literal", vec![]),
        ("%s", vec![Value::String("é😀".into())]),
        (
            "%s|%d|%s|%d",
            vec![
                Value::String("first".into()),
                Value::I32(12),
                Value::String("last".into()),
                Value::F64(-3.0),
            ],
        ),
        (
            "%s:%s:%s",
            vec![Value::Null, Value::Bool(true), Value::I32(7)],
        ),
    ];
    for (fmt, args) in cases {
        let old = program(true, args.len() as u8, 1);
        let new = program(false, args.len() as u8, 1);
        assert_eq!(
            execute(&new, fmt, &args),
            execute(&old, fmt, &args),
            "{fmt}"
        );
        for chunks in [&old, &new] {
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(chunks))
                .unwrap();
        }
    }
}
#[test]
fn repeated_argument_packing_uses_constant_scratch_space() {
    let old = program(true, 32, 20);
    let new = program(false, 32, 20);
    assert!(new[0].local_count < old[0].local_count);
    assert!(new[0].code.len() < old[0].code.len());
    println!(
        "sprintf 20 calls/32 args locals {} -> {}, bytecode {} -> {}",
        old[0].local_count,
        new[0].local_count,
        old[0].code.len(),
        new[0].code.len()
    );
    for chunks in [&old, &new] {
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(chunks))
            .unwrap();
    }
}

#[test]
fn packing_retains_array_argument_identity() {
    let object = Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(vec![Value::I32(5)]),
    ));
    let args = vec![Value::I32(1), object.clone(), Value::String("end".into())];
    for old in [true, false] {
        let chunks = program_with_identity(old, 3, 1, true);
        let result = execute(&chunks, "%s%s%s", &args);
        let Value::Object(array) = result else {
            panic!("expected packed arguments")
        };
        let array = array.lock().unwrap();
        let vybe_runtime::value::ObjectKind::Array(values) = &array.kind else {
            panic!("expected array")
        };
        assert_eq!(values, &args);
        let (Value::Object(actual), Value::Object(original)) = (&values[1], &object) else {
            unreachable!()
        };
        assert!(std::sync::Arc::ptr_eq(actual, original));
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
    }
}

#[test]
#[ignore = "interleaved opt-level-0 packing benchmark"]
fn benchmark_variadic_packing() {
    use std::time::Instant;
    for count in [4, 32, 128] {
        let old = program(true, count, 128);
        let new = program(false, count, 128);
        let mut times = [Vec::new(), Vec::new()];
        for sample in 0..7 {
            for which in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut vm = VM::new();
                platforms::register_platforms_all(&mut vm);
                vm.set_global("fmt", Value::String("literal".into()));
                for i in 0..count {
                    vm.set_global(&format!("arg{i}"), Value::I32(i as i32));
                }
                let chunks = if which == 0 { old.clone() } else { new.clone() };
                let start = Instant::now();
                let result = vm.run(chunks).unwrap();
                times[which].push(start.elapsed().as_micros());
                assert_eq!(result, Value::String("literal".into()));
            }
        }
        for time in &mut times {
            time.sort_unstable();
        }
        println!(
            "sprintf 128 calls/{count} args median {} -> {} us, {:.2}x",
            times[0][3],
            times[1][3],
            times[0][3] as f64 / times[1][3] as f64
        );
    }
}
