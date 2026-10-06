// Separate writable-public-helper regression; active diagnostic remains frozen.
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, functions, globals, platforms};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/text_optimization_str_replace_typed_frozen.rs"]
mod frozen;
fn program(selected: bool) -> Vec<Chunk> {
    let mut chunks = vec![
        Chunk::new("<replace-alias-owner>"),
        Chunk::new("replace_alias_callee"),
    ];
    functions::emit_ref_func(&mut chunks[0], 1, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    chunks[1].arity = 3;
    chunks[1].local_count = 3;
    for slot in 0..3 {
        chunks[1].emit_op_u16(Op::LOCAL_GET, slot, 0);
    }
    if selected {
        static REGISTER: std::sync::Once = std::sync::Once::new();
        REGISTER.call_once(vybe_language_php::register);
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let decl = &profile.builtins["str_replace"];
        let vybe_compiler::profile::BuiltinEmit::Common(key) = &decl.emit else {
            panic!("installed PHP profile must use its registered common emitter");
        };
        assert_eq!(key, "php.str_replace");
        assert!(vybe_compiler::primitives::dispatch::emit_common(
            key,
            &mut chunks,
            1,
            3,
            0
        ));
    } else {
        frozen::emit_str_replace(&mut chunks, 1, 3, 0);
    }
    chunks[1].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    chunks
}
fn bytecode_helper(vm: &mut VM) -> Value {
    let mut root = Chunk::new("<replace-writable-helper-owner>");
    functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut body = Chunk::new("replace-writable-dyneq");
    body.arity = 2;
    body.local_count = 2;
    body.emit_op_u16(Op::LOCAL_GET, 0, 0);
    body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let import = body.add_import("test:replace", "dyneq");
    body.emit_call(import, 2, 0);
    body.emit_op(Op::RETURN, 0);
    root.imports = body.imports.clone();
    let host = vm
        .resolve_host_function_index("test:replace", "dyneq")
        .unwrap();
    vm.run_linked(
        vec![root, body],
        vec![vybe_runtime::vm::ImportTarget::Host(host)],
    )
    .unwrap()
}
#[test]
fn rebound_public_dyneq_keeps_callback_branch_and_checked_error_contract() {
    let programs = [program(false), program(true)];
    for result in [Value::I32(0), Value::I32(1), Value::F64(0.5)] {
        let mut observed = vec![];
        for chunks in &programs {
            let mut vm = VM::new();
            platforms::register_platforms_all(&mut vm);
            let callee = vm.run(chunks.clone()).unwrap();
            let trace = Arc::new(Mutex::new(Vec::<Vec<Value>>::new()));
            let capture = trace.clone();
            let returned = result.clone();
            vm.register_host_fn(
                "test:replace",
                "dyneq",
                Box::new(move |_, args| {
                    capture.lock().unwrap().push(args.to_vec());
                    returned.clone()
                }),
            );
            // CALL_REF uses the genuine bytecode arity-two function ABI;
            // its linked CALL then delivers both explicit operands to the host.
            let helper = bytecode_helper(&mut vm);
            vm.set_global("__vybe_dyneq", helper);
            let output = vm
                .invoke(
                    &callee,
                    &[
                        Value::String("a".into()),
                        Value::String("X".into()),
                        Value::String("banana".into()),
                    ],
                )
                .map_err(|mut e| {
                    for frame in &mut e.call_stack {
                        frame.offset = 0;
                    }
                    e.to_string()
                });
            let calls = trace.lock().unwrap().clone();
            assert_eq!(
                calls,
                vec![vec![Value::I32(1), Value::F64(0.0)]],
                "public helper still called exactly once with original producer values"
            );
            match result {
                Value::I32(0) => assert_eq!(output, Ok(Value::String("bXnXnX".into()))),
                Value::I32(1) => assert_eq!(output, Ok(Value::String("banana".into()))),
                _ => assert!(
                    output.is_err(),
                    "checked toI32 must reject fractional rebound-helper output"
                ),
            }
            observed.push((output, calls));
        }
        assert_eq!(observed[0], observed[1]);
    }
}
