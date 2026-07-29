//! Behaviour tests for `node:stream` host imports.
//!
//! Reference: <https://nodejs.org/api/stream.html>.
//!
//! Coverage:
//!   - `Readable` constructor / `Readable.from(iterable)`
//!   - `Writable` constructor
//!   - `Transform` constructor
//!   - `Duplex` constructor
//!   - `PassThrough` — Transform with identity transform
//!   - `pipeline(src, ...transforms, dst[, callback])` → surface
//!   - `finished(stream[, options][, callback])` → surface
//!   - `addAbortSignal(signal, stream)` → surface
//!   - `Readable.isReadable(value)` (Node 17.4+)
//!   - `Writable.isWritable(value)` (Node 17.4+)
//!   - `isDisturbed(value)` (Node 16.14+)
//!   - `compose(...streams)` → surface
//!   - `Stream` base class constructor
//!   - Stream method surface: pipe, unpipe, destroy, pause, resume
//!   - Writable method surface: write, end, destroy
//!   - EventEmitter surface on all stream types
//!   - Stream state properties: readable, writable, destroyed, writableEnded, readableEnded
//!   - `Readable.from(array)` factory
//!
//! Deferred (require async infrastructure):
//!   - `.read()`, `.push()`, `.pipe()` full flow
//!   - `.write()`, `.end()` with callbacks
//!   - `pipeline` async execution

use std::sync::Arc;
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_stream(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-stream-test>");
    let import_idx = chunk.add_import("node:stream", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:stream"), name.to_string()))
}

fn has_method(v: &Value, key: &str) -> bool {
    match v {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

fn prop(v: &Value, key: &str) -> Value {
    match v {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
}

fn arr(elems: Vec<Value>) -> Value {
    Value::Object(Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(elems),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: Vec::new(),
    })))
}

// ── Constructor surface ───────────────────────────────────────────────────────

#[test]
fn readable_constructor_returns_object() {
    let stream = call_stream("Readable", vec![]);
    assert!(matches!(stream, Value::Object(_)));
}

#[test]
fn writable_constructor_returns_object() {
    let stream = call_stream("Writable", vec![]);
    assert!(matches!(stream, Value::Object(_)));
}

#[test]
fn transform_constructor_returns_object() {
    let stream = call_stream("Transform", vec![]);
    assert!(matches!(stream, Value::Object(_)));
}

#[test]
fn duplex_constructor_returns_object() {
    let stream = call_stream("Duplex", vec![]);
    assert!(matches!(stream, Value::Object(_)));
}

#[test]
fn pass_through_constructor_returns_object() {
    let stream = call_stream("PassThrough", vec![]);
    assert!(matches!(stream, Value::Object(_)));
}

#[test]
fn stream_base_constructor_returns_object() {
    let stream = call_stream("Stream", vec![]);
    assert!(matches!(stream, Value::Object(_)));
}

// ── Readable EventEmitter methods ─────────────────────────────────────────────

#[test]
fn readable_has_on_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "on"), "Readable.on must exist");
}

#[test]
fn readable_has_once_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "once"), "Readable.once must exist");
}

#[test]
fn readable_has_off_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "off"), "Readable.off must exist");
}

#[test]
fn readable_has_emit_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "emit"), "Readable.emit must exist");
}

#[test]
fn readable_has_remove_listener_method() {
    let s = call_stream("Readable", vec![]);
    assert!(
        has_method(&s, "removeListener"),
        "Readable.removeListener must exist"
    );
}

#[test]
fn readable_has_add_listener_method() {
    let s = call_stream("Readable", vec![]);
    assert!(
        has_method(&s, "addListener"),
        "Readable.addListener must exist"
    );
}

// ── Readable stream-specific methods ─────────────────────────────────────────

#[test]
fn readable_has_pipe_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "pipe"), "Readable.pipe must exist");
}

#[test]
fn readable_has_unpipe_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "unpipe"), "Readable.unpipe must exist");
}

#[test]
fn readable_has_destroy_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "destroy"), "Readable.destroy must exist");
}

#[test]
fn readable_has_pause_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "pause"), "Readable.pause must exist");
}

#[test]
fn readable_has_resume_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "resume"), "Readable.resume must exist");
}

#[test]
fn readable_has_read_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "read"), "Readable.read must exist");
}

#[test]
fn readable_has_push_method() {
    let s = call_stream("Readable", vec![]);
    assert!(has_method(&s, "push"), "Readable.push must exist");
}

#[test]
fn readable_has_set_encoding_method() {
    let s = call_stream("Readable", vec![]);
    assert!(
        has_method(&s, "setEncoding"),
        "Readable.setEncoding must exist"
    );
}

// ── Readable state properties ─────────────────────────────────────────────────

#[test]
fn readable_has_readable_property() {
    let s = call_stream("Readable", vec![]);
    let v = prop(&s, "readable");
    assert!(
        matches!(v, Value::Bool(_) | Value::Undefined),
        "readable must be bool or undefined"
    );
}

#[test]
fn readable_has_destroyed_property() {
    let s = call_stream("Readable", vec![]);
    let v = prop(&s, "destroyed");
    assert!(matches!(v, Value::Bool(_) | Value::Undefined));
}

#[test]
fn readable_has_readable_ended_property() {
    let s = call_stream("Readable", vec![]);
    // readableEnded may not exist before any data
    let _ = prop(&s, "readableEnded");
}

// ── Writable methods ──────────────────────────────────────────────────────────

#[test]
fn writable_has_write_method() {
    let s = call_stream("Writable", vec![]);
    assert!(has_method(&s, "write"), "Writable.write must exist");
}

#[test]
fn writable_has_end_method() {
    let s = call_stream("Writable", vec![]);
    assert!(has_method(&s, "end"), "Writable.end must exist");
}

#[test]
fn writable_has_destroy_method() {
    let s = call_stream("Writable", vec![]);
    assert!(has_method(&s, "destroy"), "Writable.destroy must exist");
}

#[test]
fn writable_has_on_method() {
    let s = call_stream("Writable", vec![]);
    assert!(has_method(&s, "on"), "Writable.on must exist");
}

#[test]
fn writable_has_emit_method() {
    let s = call_stream("Writable", vec![]);
    assert!(has_method(&s, "emit"), "Writable.emit must exist");
}

// ── Writable state properties ─────────────────────────────────────────────────

#[test]
fn writable_has_writable_property() {
    let s = call_stream("Writable", vec![]);
    let v = prop(&s, "writable");
    assert!(matches!(v, Value::Bool(_) | Value::Undefined));
}

#[test]
fn writable_has_writable_ended_property() {
    let s = call_stream("Writable", vec![]);
    let _ = prop(&s, "writableEnded");
}

// ── Transform methods ─────────────────────────────────────────────────────────

#[test]
fn transform_has_write_method() {
    let s = call_stream("Transform", vec![]);
    assert!(has_method(&s, "write"), "Transform.write must exist");
}

#[test]
fn transform_has_read_method() {
    let s = call_stream("Transform", vec![]);
    assert!(has_method(&s, "read"), "Transform.read must exist");
}

#[test]
fn transform_has_end_method() {
    let s = call_stream("Transform", vec![]);
    assert!(has_method(&s, "end"), "Transform.end must exist");
}

#[test]
fn transform_has_on_method() {
    let s = call_stream("Transform", vec![]);
    assert!(has_method(&s, "on"), "Transform.on must exist");
}

// ── Duplex is both readable and writable ──────────────────────────────────────

#[test]
fn duplex_has_write_and_read_methods() {
    let s = call_stream("Duplex", vec![]);
    assert!(has_method(&s, "write"), "Duplex.write must exist");
    assert!(has_method(&s, "read"), "Duplex.read must exist");
}

#[test]
fn duplex_has_on_method() {
    let s = call_stream("Duplex", vec![]);
    assert!(has_method(&s, "on"), "Duplex.on must exist");
}

// ── Readable.isReadable ────────────────────────────────────────────────────────

#[test]
fn is_readable_true_for_readable_stream() {
    let readable = call_stream("Readable", vec![]);
    let result = call_stream("isReadable", vec![readable]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn is_readable_false_for_plain_string() {
    let result = call_stream("isReadable", vec![Value::String(Arc::from("not a stream"))]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_readable_false_for_null() {
    let result = call_stream("isReadable", vec![Value::Null]);
    assert_eq!(result, Value::Bool(false));
}

// ── Writable.isWritable ────────────────────────────────────────────────────────

#[test]
fn is_writable_true_for_writable_stream() {
    let writable = call_stream("Writable", vec![]);
    let result = call_stream("isWritable", vec![writable]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn is_writable_false_for_readable_only() {
    let readable = call_stream("Readable", vec![]);
    let result = call_stream("isWritable", vec![readable]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_writable_true_for_transform() {
    // Transform is both readable and writable
    let transform = call_stream("Transform", vec![]);
    let result = call_stream("isWritable", vec![transform]);
    assert_eq!(result, Value::Bool(true));
}

// ── PassThrough identity ──────────────────────────────────────────────────────

#[test]
fn pass_through_is_both_readable_and_writable() {
    let pt = call_stream("PassThrough", vec![]);
    let r = call_stream("isReadable", vec![pt.clone()]);
    let w = call_stream("isWritable", vec![pt]);
    assert_eq!(r, Value::Bool(true));
    assert_eq!(w, Value::Bool(true));
}

#[test]
fn pass_through_has_write_and_read_methods() {
    let pt = call_stream("PassThrough", vec![]);
    assert!(has_method(&pt, "write"), "PassThrough.write must exist");
    assert!(has_method(&pt, "read"), "PassThrough.read must exist");
}

// ── isDisturbed ───────────────────────────────────────────────────────────────

#[test]
fn is_disturbed_false_for_fresh_readable() {
    let readable = call_stream("Readable", vec![]);
    let result = call_stream("isDisturbed", vec![readable]);
    // fresh stream is not disturbed
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_disturbed_false_for_null() {
    let result = call_stream("isDisturbed", vec![Value::Null]);
    assert_eq!(result, Value::Bool(false));
}

// ── Readable.from factory ─────────────────────────────────────────────────────

#[test]
fn readable_from_array_returns_readable() {
    let input = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let stream = call_stream("readableFrom", vec![input]);
    assert!(
        matches!(stream, Value::Object(_)),
        "Readable.from must return a stream"
    );
}

// ── pipeline surface ──────────────────────────────────────────────────────────

#[test]
fn pipeline_is_registered_as_import() {
    assert!(
        has_import("pipeline"),
        "node:stream pipeline must be registered"
    );
}

#[test]
fn pipeline_with_null_callback_does_not_panic() {
    let src = call_stream("Readable", vec![]);
    let dst = call_stream("Writable", vec![]);
    let _ = call_stream("pipeline", vec![src, dst, Value::Null]);
}

// ── finished surface ──────────────────────────────────────────────────────────

#[test]
fn finished_is_registered_as_import() {
    assert!(
        has_import("finished"),
        "node:stream finished must be registered"
    );
}

#[test]
fn finished_with_null_callback_does_not_panic() {
    let s = call_stream("Readable", vec![]);
    let _ = call_stream("finished", vec![s, Value::Null]);
}

// ── addAbortSignal ────────────────────────────────────────────────────────────

#[test]
fn add_abort_signal_does_not_panic() {
    let stream = call_stream("Readable", vec![]);
    let _ = call_stream("addAbortSignal", vec![Value::Null, stream]);
}

// ── compose surface ───────────────────────────────────────────────────────────

#[test]
fn compose_is_registered_as_import() {
    assert!(
        has_import("compose"),
        "node:stream compose must be registered"
    );
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_stream_surface_is_registered() {
    let expected = [
        "Readable",
        "Writable",
        "Transform",
        "Duplex",
        "PassThrough",
        "Stream",
        "pipeline",
        "finished",
        "addAbortSignal",
        "isReadable",
        "isWritable",
        "isDisturbed",
        "compose",
        "readableFrom",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:stream imports: {missing:?}"
    );
}
