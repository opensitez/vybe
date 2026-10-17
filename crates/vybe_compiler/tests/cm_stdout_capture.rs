#[path = "fixtures/cm_stdout_capture.rs"]
mod capture;

use std::sync::{Arc, Mutex};
use vybe_runtime::{Chunk, Op, VM, Value};

#[test]
fn php_bundle_output_reaches_sink_first_capture() {
    use vybe_compiler::bundle::{Bundle, EntryPoint, SourceFile};
    vybe_language_php::register();
    let bundle = Bundle {
        name: "cm_stdout_capture".into(),
        language: vybe_compiler::languages::find_by_name("php").unwrap(),
        sources: vec![SourceFile {
            path: std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("tests/fixtures/cm_stdout_capture.php"),
            code: include_str!("fixtures/cm_stdout_capture.php").into(),
        }],
        wasm_files: vec![],
        entry_point: EntryPoint::Auto,
    };
    let compiled = bundle.compile_full().unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let output = Arc::new(Mutex::new(Vec::new()));
    capture::register(&mut vm, Arc::clone(&output));
    vm.run(compiled.chunks).unwrap();
    assert_eq!(output.lock().unwrap().concat(), "left\u{df}right");
}

#[test]
fn sink_first_capture_waits_for_eof_and_preserves_split_utf8() {
    let output = Arc::new(Mutex::new(Vec::new()));
    let mut vm = VM::new();
    capture::register(&mut vm, Arc::clone(&output));
    let mut start = Chunk::new("capture-open-stream");
    start.memory_min_pages.push(1);
    let new = start.add_import("canon", "stream.new");
    let sink = start.add_import("wasi:cli/stdout", "write-via-stream");
    let packed = start.i64_scratch();
    start.emit_call(new, 0, 0);
    start.emit_op_u16(Op::LOCAL_SET, packed, 0);
    start.emit_op_u16(Op::LOCAL_GET, packed, 0);
    start.emit_op(Op::I32_WRAP_I64, 0);
    start.emit_call(sink, 1, 0);
    start.emit_op(Op::DROP, 0);
    start.emit_op_u16(Op::LOCAL_GET, packed, 0);
    start.emit_i64_const(32, 0);
    start.emit_op(Op::I64_SHR_U, 0);
    start.emit_op(Op::I32_WRAP_I64, 0);
    start.emit_op(Op::RETURN, 0);
    let writer = vm.run(vec![start]).unwrap().as_i32();
    assert!(output.lock().unwrap().is_empty());
    assert_eq!(
        vm.event_loop
            .borrow()
            .future_states
            .values()
            .next()
            .unwrap()
            .phase,
        vybe_runtime::event_loop::FuturePhase::Pending
    );
    vm.memory.write_array(8, &[0xc3, 0x9f]).unwrap();
    let mut finish = Chunk::new("capture-split-codepoint");
    let write = finish.add_import("canon", "stream.write");
    let close = finish.add_import("canon", "stream.drop-writable");
    for ptr in [8, 9] {
        finish.emit_i32_const(writer, 0);
        finish.emit_i32_const(ptr, 0);
        finish.emit_i32_const(1, 0);
        finish.emit_call(write, 3, 0);
        finish.emit_op(Op::DROP, 0);
    }
    finish.emit_i32_const(writer, 0);
    finish.emit_call(close, 1, 0);
    finish.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    finish.emit_op(Op::RETURN, 0);
    vm.run(vec![finish]).unwrap();
    assert_eq!(*output.lock().unwrap(), vec!["\u{df}"]);
    let events = vm.event_loop.borrow();
    let completion = events.future_states.values().next().unwrap();
    assert_eq!(
        completion.phase,
        vybe_runtime::event_loop::FuturePhase::Resolved
    );
    assert!(matches!(completion.value, Some(Value::Null)));
}

#[test]
fn invalid_capture_stream_preserves_canonical_error_arm() {
    let mut vm = VM::new();
    capture::register(&mut vm, Arc::new(Mutex::new(Vec::new())));
    let mut chunk = Chunk::new("capture-invalid-stream");
    let sink = chunk.add_import("wasi:cli/stdout", "write-via-stream");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_call(sink, 1, 0);
    chunk.emit_op(Op::RETURN, 0);
    assert!(matches!(vm.run(vec![chunk]).unwrap(), Value::I32(_)));
    let events = vm.event_loop.borrow();
    let completion = events.future_states.values().next().unwrap();
    let (case, payload) =
        vybe_runtime::canon_value::result_parts_public(completion.value.as_ref().unwrap());
    assert_eq!(case, 1);
    assert!(matches!(payload, Value::String(ref code) if code.as_ref() == "io"));
}
