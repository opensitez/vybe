//! Execute the common collection emitters independently of frontend lowering.
use vybe_compiler::primitives::{functions, globals, loops};
use vybe_runtime::{Chunk, Op, Value, VM, value::{Object, ObjectKind}};

fn check(kind: &str, items: &[i32], expected: &[i32], expected_visits: i32, grow: bool) {
    let mut entry = Chunk::new("callback_loop");
    entry.local_count = 6;
    functions::emit_ref_func(&mut entry, 1, 0, 1);
    entry.emit_op_u16(Op::LOCAL_SET, 0, 1);
    globals::emit_read(&mut entry, "items", 1);
    entry.emit_op_u16(Op::LOCAL_SET, 1, 1);
    let mut chunks = vec![entry];
    match kind {
        "map" => loops::emit_map(&mut chunks, 0, 0, 1, 2, 3, 1),
        "filter" => loops::emit_filter(&mut chunks, 0,
            vybe_runtime::chunk::ReceiverAbi::Ambient, 0, 1, 2, 3, 4, 1),
        "foreach" => loops::emit_foreach(&mut chunks, 0, 0, 1, 3, 1),
        "any" | "every" => loops::emit_any_every(&mut chunks, 0, 0, 1, 3, 2, kind == "any", 1),
        _ => unreachable!(),
    }
    chunks[0].emit_op(Op::RETURN, 1);
    let mut callback = Chunk::new("callback");
    callback.arity = if matches!(kind, "map" | "foreach") { 2 } else { 1 };
    callback.param_count = callback.arity;
    callback.local_count = callback.arity as u16;
    if grow {
        globals::emit_read(&mut callback, "visits", 1);
        callback.emit_op(Op::I32_EQZ, 1);
        callback.emit_if(1);
        globals::emit_read(&mut callback, "items", 1);
        callback.emit_i32_const(2, 1);
        let push = callback.add_import("ecma:array", "push");
        callback.emit_call(push, 2, 1);
        callback.emit_op(Op::DROP, 1);
        callback.emit_end(1);
    }
    globals::emit_read(&mut callback, "visits", 1);
    callback.emit_i32_const(1, 1);
    callback.emit_op(Op::I32_ADD, 1);
    globals::emit_write(&mut callback, "visits", 1);
    callback.emit_op_u16(Op::LOCAL_GET, 0, 1);
    callback.emit_i32_const(if kind == "map" { 2 } else { 1 }, 1);
    callback.emit_op(if kind == "map" { Op::I32_MUL } else { Op::I32_AND }, 1);
    callback.emit_op(Op::RETURN, 1);
    chunks.push(callback);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);

    let wasm = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    vybe_platform_wasm::validate_binary(&wasm).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.set_global_owned("items", Value::Object(vybe_runtime::heap::alloc(
        Object::new_array(items.iter().copied().map(Value::I32).collect()),
    )));
    vm.set_global_owned("visits", Value::I32(0));
    let result = vm.run(chunks).unwrap();
    assert_eq!(vm.global("visits").unwrap().as_i32(), expected_visits, "{kind}");
    match kind {
        "map" | "filter" => {
            let Value::Object(result) = result else { panic!("expected array") };
            let result = result.lock().unwrap();
            let ObjectKind::Array(values) = &result.kind else { panic!("expected array") };
            assert_eq!(values.iter().map(Value::as_i32).collect::<Vec<_>>(), expected);
        }
        "any" | "every" => assert_eq!(result.as_i32(), expected[0], "{kind}"),
        "foreach" => assert!(matches!(result, Value::Null)),
        _ => unreachable!(),
    }
}

#[test]
fn callback_loops_preserve_results_and_short_circuiting() {
    check("map", &[1,2,3], &[2,4,6], 3, false);
    check("filter", &[1,2,3], &[1,3], 3, false);
    check("foreach", &[1,2,3], &[], 3, false);
    check("any", &[1,2,3], &[1], 1, false);
    check("every", &[1,2,3], &[0], 2, false);
    check("any", &[], &[0], 0, false);
    check("every", &[], &[1], 0, false);
}

#[test]
fn callback_loop_retains_existing_live_length_behavior() {
    check("foreach", &[1], &[], 2, true);
}
