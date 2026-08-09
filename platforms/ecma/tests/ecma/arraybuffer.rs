//! Behaviour tests for `ecma:arraybuffer` and `ecma:sharedarraybuffer`
//! host imports.
//!
//! Reference: ECMA-262 §25.1 ArrayBuffer, §25.2 SharedArrayBuffer.
//!
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.globals.insert(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke(ns: &str, name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-arraybuffer-test>");
    let import_idx = chunk.add_import(ns, name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn ab(name: &str, args: Vec<Value>) -> Value {
    invoke("ecma:arraybuffer", name, args)
}
fn sab(name: &str, args: Vec<Value>) -> Value {
    invoke("ecma:sharedarraybuffer", name, args)
}

fn plain_obj() -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new())))
}

// ── ArrayBuffer construction ──────────────────────────────────────────────────

#[test]
fn new_with_length_sets_byte_length() {
    let buf = ab("newWithLength", vec![Value::I32(16)]);
    assert_eq!(ab("byteLength", vec![buf]), Value::I32(16));
}

#[test]
fn new_with_zero_length_has_byte_length_zero() {
    let buf = ab("newWithLength", vec![Value::I32(0)]);
    assert_eq!(ab("byteLength", vec![buf]), Value::I32(0));
}

// ── ArrayBuffer.isView — distinguishes views from buffers ────────────────────

#[test]
fn is_view_false_for_raw_arraybuffer() {
    // The buffer itself is not a view; only TypedArrays and DataView are views.
    let buf = ab("newWithLength", vec![Value::I32(8)]);
    assert_eq!(ab("isView", vec![buf]), Value::Bool(false));
}

#[test]
fn is_view_false_for_plain_object() {
    assert_eq!(ab("isView", vec![plain_obj()]), Value::Bool(false));
}

#[test]
fn is_view_false_for_null() {
    assert_eq!(ab("isView", vec![Value::Null]), Value::Bool(false));
}

// ── slice — creates an independent COPY ──────────────────────────────────────

#[test]
fn slice_returns_new_buffer_with_different_pointer() {
    let buf = ab("newWithLength", vec![Value::I32(8)]);
    let sliced = ab("slice", vec![buf.clone(), Value::I32(0), Value::I32(4)]);
    let a_ptr = match &buf {
        Value::Object(o) => Arc::as_ptr(o) as usize,
        _ => 0,
    };
    let b_ptr = match &sliced {
        Value::Object(o) => Arc::as_ptr(o) as usize,
        _ => 1,
    };
    assert_ne!(
        a_ptr, b_ptr,
        "slice must return an independent buffer, not the same Arc"
    );
}

#[test]
fn slice_byte_length_reflects_the_requested_range() {
    let buf = ab("newWithLength", vec![Value::I32(16)]);
    let sliced = ab("slice", vec![buf, Value::I32(4), Value::I32(12)]);
    assert_eq!(ab("byteLength", vec![sliced]), Value::I32(8));
}

#[test]
fn slice_negative_begin_counts_from_end() {
    // slice(-4) on a 16-byte buffer → last 4 bytes → byteLength 4.
    let buf = ab("newWithLength", vec![Value::I32(16)]);
    let sliced = ab("slice", vec![buf, Value::I32(-4)]);
    assert_eq!(ab("byteLength", vec![sliced]), Value::I32(4));
}

#[test]
fn slice_with_end_past_buffer_clamps_to_byte_length() {
    let buf = ab("newWithLength", vec![Value::I32(8)]);
    let sliced = ab("slice", vec![buf, Value::I32(0), Value::I32(999)]);
    assert_eq!(ab("byteLength", vec![sliced]), Value::I32(8));
}

// ── Resizable ArrayBuffer (stage 3 / V8 / Safari) ────────────────────────────

#[test]
fn new_resizable_sets_resizable_flag_to_true() {
    let buf = ab("newResizable", vec![Value::I32(8), Value::I32(64)]);
    assert_eq!(ab("resizable", vec![buf]), Value::Bool(true));
}

#[test]
fn new_resizable_stores_max_byte_length() {
    let buf = ab("newResizable", vec![Value::I32(8), Value::I32(64)]);
    assert_eq!(ab("maxByteLength", vec![buf]), Value::I32(64));
}

#[test]
fn fixed_buffer_is_not_resizable() {
    let buf = ab("newWithLength", vec![Value::I32(8)]);
    assert_eq!(ab("resizable", vec![buf]), Value::Bool(false));
}

#[test]
fn resize_changes_byte_length_within_max() {
    let buf = ab("newResizable", vec![Value::I32(8), Value::I32(64)]);
    ab("resize", vec![buf.clone(), Value::I32(32)]);
    assert_eq!(ab("byteLength", vec![buf]), Value::I32(32));
}

// ── transfer — detaches the original, returns new buffer ─────────────────────

#[test]
fn transfer_returns_new_buffer_with_same_byte_length() {
    let buf = ab("newWithLength", vec![Value::I32(16)]);
    let transferred = ab("transfer", vec![buf]);
    assert_eq!(ab("byteLength", vec![transferred]), Value::I32(16));
}

#[test]
fn transfer_detaches_original_buffer() {
    // After transfer, the original becomes detached (byteLength → 0 or
    // a property `detached` → true, depending on the host).
    let buf = ab("newWithLength", vec![Value::I32(16)]);
    let _transferred = ab("transfer", vec![buf.clone()]);
    let byte_len = ab("byteLength", vec![buf.clone()]);
    let detached = ab("detached", vec![buf]);
    // Either convention is acceptable.
    assert!(
        byte_len == Value::I32(0) || detached == Value::Bool(true),
        "original buffer must be detached after transfer"
    );
}

// ── SharedArrayBuffer — no detach, can grow ───────────────────────────────────

#[test]
fn sharedarraybuffer_new_with_length_sets_byte_length() {
    let buf = sab("newWithLength", vec![Value::I32(32)]);
    assert_eq!(sab("byteLength", vec![buf]), Value::I32(32));
}

#[test]
fn sharedarraybuffer_is_not_growable_by_default() {
    let buf = sab("newWithLength", vec![Value::I32(16)]);
    assert_eq!(sab("growable", vec![buf]), Value::Bool(false));
}

#[test]
fn sharedarraybuffer_new_growable_sets_growable_true_and_max_byte_length() {
    let buf = sab("newGrowable", vec![Value::I32(8), Value::I32(128)]);
    assert_eq!(sab("growable", vec![buf.clone()]), Value::Bool(true));
    assert_eq!(sab("maxByteLength", vec![buf]), Value::I32(128));
}

#[test]
fn sharedarraybuffer_grow_increases_byte_length() {
    let buf = sab("newGrowable", vec![Value::I32(8), Value::I32(128)]);
    sab("grow", vec![buf.clone(), Value::I32(64)]);
    assert_eq!(sab("byteLength", vec![buf]), Value::I32(64));
}

// ── ArrayBuffer.prototype.transferToFixedLength (ES2024 §25.1.5.6) ───────────

#[test]
fn transfer_to_fixed_length_returns_non_resizable_buffer() {
    // ECMA-262 ES2024: transferToFixedLength creates a fixed-length copy and detaches original.
    let buf = ab("newResizable", vec![Value::I32(8), Value::I32(64)]);
    let fixed = ab("transferToFixedLength", vec![buf.clone()]);
    // The new buffer must be a non-resizable ArrayBuffer.
    assert!(matches!(fixed, Value::Object(_)));
    assert_eq!(ab("resizable", vec![fixed.clone()]), Value::Bool(false));
    assert_eq!(ab("byteLength", vec![fixed]), Value::I32(8));
}

#[test]
fn transfer_to_fixed_length_detaches_the_original() {
    // After transferToFixedLength, the original buffer must be detached.
    let buf = ab("newWithLength", vec![Value::I32(16)]);
    let _fixed = ab("transferToFixedLength", vec![buf.clone()]);
    // Detached buffer has byteLength 0 or detached flag.
    let after = ab("byteLength", vec![buf.clone()]);
    let detached = ab("detached", vec![buf]);
    assert!(
        matches!(after, Value::I32(0)) || matches!(detached, Value::Bool(true)),
        "original must be detached"
    );
}
