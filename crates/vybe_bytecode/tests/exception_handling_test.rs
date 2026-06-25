//! Tests for WASM exception handling: THROW (0x08), THROW_REF (0x0A), TRY_TABLE (0x1F).
//!
//! TRY_TABLE binary layout:
//!   [op:2][handler_count:1][ (tag:1)(offset_hi:1)(offset_lo:1) × N ]
//!   ip after operands = start of body.
//!   catch_ip = ip + offset  (big-endian u16 offset from body start).

use std::sync::Arc;
use vybe_bytecode::value::Value;
use vybe_bytecode::wasm;
use vybe_bytecode::{Chunk, Op, VM};

fn write_leb_u32(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn push_section(out: &mut Vec<u8>, id: u8, payload: &[u8]) {
    out.push(id);
    write_leb_u32(out, payload.len() as u32);
    out.extend_from_slice(payload);
}

fn standard_eh_module(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut out, 3, &[0x01, 0x00]);

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(body_ops);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    run_locals(0, emit)
}

fn run_locals(local_count: u16, emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = local_count;
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    VM::new().run(vec![chunk]).expect("run failed")
}

fn run_err(emit: impl FnOnce(&mut Chunk)) -> String {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    VM::new().run(vec![chunk]).unwrap_err().to_string()
}

#[test]
fn standard_rethrow_must_not_decode_as_noop() {
    let bytes = standard_eh_module(&[
        0x09, 0x00, // rethrow 0
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("rethrow should decode");
    assert!(chunks[1].code.windows(4).any(|w| w == Op::RETHROW.encode()));
}

#[test]
fn standard_delegate_must_not_decode_as_noop() {
    let bytes = standard_eh_module(&[
        0x18, 0x00, // delegate 0
    ]);

    let chunks = wasm::read_wasm(&bytes).expect("delegate should decode");
    assert!(chunks[1].code.windows(4).any(|w| w == Op::DELEGATE.encode()));
}

/// Emit TRY_TABLE with one catch-all handler. Returns the offset_pos to patch.
fn emit_try_table_catch_all_start(c: &mut Chunk) -> usize {
    c.emit_op(Op::TRY_TABLE, 0);
    c.emit(1, 0); // handler_count = 1
    c.emit(0, 0); // tag = 0 (catch-all)
    let offset_pos = c.current_offset();
    c.emit(0, 0); // offset hi placeholder
    c.emit(0, 0); // offset lo placeholder
    offset_pos
}

fn patch_try_table(c: &mut Chunk, offset_pos: usize) {
    let body_bytes = c.current_offset() - (offset_pos + 2);
    c.code[offset_pos] = (body_bytes >> 8) as u8;
    c.code[offset_pos + 1] = (body_bytes & 0xFF) as u8;
}

fn emit_rethrow(c: &mut Chunk, depth: u32) {
    c.emit_op(Op::RETHROW, 0);
    c.emit_leb_u32(depth, 0);
}

fn emit_delegate(c: &mut Chunk, depth: u32) {
    c.emit_op(Op::DELEGATE, 0);
    c.emit_leb_u32(depth, 0);
}

// ── THROW — uncaught ──────────────────────────────────────────────────────

#[test]
fn throw_uncaught_propagates_as_error() {
    let e = run_err(|c| {
        let k = c.add_constant(Value::String(Arc::from("boom")));
        c.emit_op_u16(Op::CONST, k, 0);
        c.emit_op(Op::THROW, 0);
    });
    assert!(e.contains("boom"));
}

// ── THROW_REF — uncaught ──────────────────────────────────────────────────

#[test]
fn throw_ref_uncaught_propagates() {
    let e = run_err(|c| {
        let k = c.add_constant(Value::String(Arc::from("ref-throw")));
        c.emit_op_u16(Op::CONST, k, 0);
        c.emit_op(Op::THROW_REF, 0);
    });
    assert!(e.contains("ref-throw"));
}

// ── TRY_TABLE + THROW — catch-all ────────────────────────────────────────

#[test]
fn try_table_catch_all_intercepts_throw() {
    let r = run(|c| {
        let err = c.add_constant(Value::String(Arc::from("oops")));
        let ok = c.add_constant(Value::I32(99));

        let off = emit_try_table_catch_all_start(c);
        c.emit_op_u16(Op::CONST, err, 0);
        c.emit_op(Op::THROW, 0);
        patch_try_table(c, off);

        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, ok, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn rethrow_in_inner_handler_is_caught_by_outer_handler() {
    let r = run(|c| {
        let err = c.add_constant(Value::String(Arc::from("nested")));
        let ok = c.add_constant(Value::I32(77));

        let outer = emit_try_table_catch_all_start(c);
        let inner = emit_try_table_catch_all_start(c);
        c.emit_op_u16(Op::CONST, err, 0);
        c.emit_op(Op::THROW, 0);
        patch_try_table(c, inner);
        emit_rethrow(c, 0);
        patch_try_table(c, outer);
        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, ok, 0);
    });
    assert_eq!(r.as_i32(), 77);
}

#[test]
fn delegate_in_inner_handler_is_caught_by_outer_handler() {
    let r = run(|c| {
        let err = c.add_constant(Value::String(Arc::from("delegated")));
        let ok = c.add_constant(Value::I32(91));

        let outer = emit_try_table_catch_all_start(c);
        let inner = emit_try_table_catch_all_start(c);
        c.emit_op_u16(Op::CONST, err, 0);
        c.emit_op(Op::THROW, 0);
        patch_try_table(c, inner);
        emit_delegate(c, 0);
        patch_try_table(c, outer);
        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, ok, 0);
    });
    assert_eq!(r.as_i32(), 91);
}

#[test]
fn delegate_depth_skips_enclosing_handler() {
    let r = run(|c| {
        let err = c.add_constant(Value::String(Arc::from("skip-one")));
        let outer = c.add_constant(Value::I32(111));
        let middle = c.add_constant(Value::I32(222));

        let o = emit_try_table_catch_all_start(c);
        let m = emit_try_table_catch_all_start(c);
        let i = emit_try_table_catch_all_start(c);
        c.emit_op_u16(Op::CONST, err, 0);
        c.emit_op(Op::THROW, 0);
        patch_try_table(c, i);
        emit_delegate(c, 1);
        patch_try_table(c, m);

        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, middle, 0);
        patch_try_table(c, o);

        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, outer, 0);
    });
    assert_eq!(r.as_i32(), 111);
}

#[test]
fn try_table_thrown_value_available_in_handler() {
    let r = run(|c| {
        let thrown = c.add_constant(Value::I32(42));

        let off = emit_try_table_catch_all_start(c);
        c.emit_op_u16(Op::CONST, thrown, 0);
        c.emit_op(Op::THROW, 0);
        patch_try_table(c, off);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn try_table_no_throw_falls_through() {
    let r = run(|c| {
        let ok = c.add_constant(Value::I32(7));

        let off = emit_try_table_catch_all_start(c);
        c.emit_op_u16(Op::CONST, ok, 0);
        // run() adds RETURN — no throw, handler never runs
        // Patch offset to point past the body (even though handler won't run)
        patch_try_table(c, off);
    });
    assert_eq!(r.as_i32(), 7);
}

// ── THROW_REF behaves like THROW for handler lookup ───────────────────────

#[test]
fn throw_ref_caught_by_try_table() {
    let r = run(|c| {
        let err = c.add_constant(Value::I32(55));
        let ok = c.add_constant(Value::I32(55));

        let off = emit_try_table_catch_all_start(c);
        c.emit_op_u16(Op::CONST, err, 0);
        c.emit_op(Op::THROW_REF, 0);
        patch_try_table(c, off);

        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, ok, 0);
    });
    assert_eq!(r.as_i32(), 55);
}

// ═══════════════════════════════════════════════════════════════════════════
// Typed exception handling (non-zero tag)
// Tag N = typed catch: matches if thrown value contains the tag name.
// String exceptions match when the string starts with or contains the tag.
// ═══════════════════════════════════════════════════════════════════════════

/// Emit TRY_TABLE with one typed handler. Returns offset_pos to patch.
fn emit_try_table_typed_start(c: &mut Chunk, tag_byte: u8) -> usize {
    c.emit_op(Op::TRY_TABLE, 0);
    c.emit(1, 0); // handler_count
    c.emit(tag_byte, 0); // tag
    let offset_pos = c.current_offset();
    c.emit(0, 0); // offset hi placeholder
    c.emit(0, 0); // offset lo placeholder
    offset_pos
}

fn emit_try_table_typed_then_catch_all_start(c: &mut Chunk, tag_byte: u8) -> (usize, usize) {
    c.emit_op(Op::TRY_TABLE, 0);
    c.emit(2, 0); // handler_count = 2
    c.emit(tag_byte, 0); // typed handler tag
    let typed_off_pos = c.current_offset();
    c.emit(0, 0); // typed handler offset hi placeholder
    c.emit(0, 0); // typed handler offset lo placeholder
    c.emit(0, 0); // catch-all tag
    let catchall_off_pos = c.current_offset();
    c.emit(0, 0); // catch-all offset hi placeholder
    c.emit(0, 0); // catch-all offset lo placeholder
    (typed_off_pos, catchall_off_pos)
}

fn patch_try_table_two(c: &mut Chunk, typed_off_pos: usize, catchall_off_pos: usize, catch_all_start: usize) {
    let typed_bytes = catch_all_start - (typed_off_pos + 2);
    c.code[typed_off_pos] = (typed_bytes >> 8) as u8;
    c.code[typed_off_pos + 1] = (typed_bytes & 0xFF) as u8;
    let catchall_bytes = catch_all_start - (catchall_off_pos + 2);
    c.code[catchall_off_pos] = (catchall_bytes >> 8) as u8;
    c.code[catchall_off_pos + 1] = (catchall_bytes & 0xFF) as u8;
}

#[test]
fn typed_catch_matches_when_exception_type_matches() {
    // Register tag "ValueError" = tag index 1 (0 is catch-all sentinel).
    // Throw a string starting with "ValueError:" — typed handler should catch it.
    let r = run(|c| {
        let tag_idx = c.add_exception_tag("ValueError");
        let err_str = c.add_constant(Value::String(Arc::from("ValueError: bad")));
        let caught = c.add_constant(Value::I32(1));
        let missed = c.add_constant(Value::I32(0));

        let off = emit_try_table_typed_start(c, tag_idx);
        c.emit_op_u16(Op::CONST, err_str, 0);
        c.emit_op(Op::THROW, 0);
        patch_try_table(c, off);

        // handler: drop thrown value, push 1
        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, caught, 0);
        let _ = missed;
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn typed_catch_does_not_match_different_exception_type() {
    // Tag "TypeError" should NOT catch a "ValueError" exception.
    // The exception falls through to become an uncaught error.
    let mut chunk = Chunk::new("<script>");
    let _type_tag = chunk.add_exception_tag("TypeError"); // tag 1

    let err_str = chunk.add_constant(Value::String(Arc::from("ValueError: wrong")));
    let fallback = chunk.add_constant(Value::I32(0));

    let off = emit_try_table_typed_start(&mut chunk, 1);
    chunk.emit_op_u16(Op::CONST, err_str, 0);
    chunk.emit_op(Op::THROW, 0);
    patch_try_table(&mut chunk, off);
    // handler (never reached for wrong type)
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::CONST, fallback, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(
        err.contains("ValueError"),
        "uncaught error should contain the thrown value"
    );
}

#[test]
fn typed_catch_with_object_exception_type() {
    // Throw an object with __exception_type = "RangeError".
    // A handler tagged "RangeError" should catch it.
    let r = run_locals(1, |c| {
        let tag_idx = c.add_exception_tag("RangeError");
        let type_key = c.add_constant(Value::String(Arc::from("__exception_type")));
        let type_val = c.add_constant(Value::String(Arc::from("RangeError")));
        let caught = c.add_constant(Value::I32(42));

        // Build an object with __exception_type = "RangeError"
        let off = emit_try_table_typed_start(c, tag_idx);
        c.emit_op_u16(Op::STRUCT_NEW, 0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::CONST, type_val, 0);
        c.emit_op_u16(Op::STRUCT_SET, type_key, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op(Op::THROW, 0);
        patch_try_table(c, off);

        // handler: drop, push 42
        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, caught, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn typed_catch_falls_through_to_later_catch_all() {
    let r = run(|c| {
        let tag_idx = c.add_exception_tag("TypeError");
        let err_str = c.add_constant(Value::String(Arc::from("ValueError: wrong")));
        let caught = c.add_constant(Value::I32(77));

        let (typed_off, catchall_off) = emit_try_table_typed_then_catch_all_start(c, tag_idx);
        c.emit_op_u16(Op::CONST, err_str, 0);
        c.emit_op(Op::THROW, 0);
        let catch_start = c.current_offset();
        patch_try_table_two(c, typed_off, catchall_off, catch_start);

        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, caught, 0);
    });
    assert_eq!(r.as_i32(), 77);
}

#[test]
fn typed_catch_precedes_later_catch_all_when_it_matches() {
    let r = run(|c| {
        let tag_idx = c.add_exception_tag("TypeError");
        let err_str = c.add_constant(Value::String(Arc::from("TypeError: bad")));
        let caught = c.add_constant(Value::I32(88));

        let (typed_off, catchall_off) = emit_try_table_typed_then_catch_all_start(c, tag_idx);
        c.emit_op_u16(Op::CONST, err_str, 0);
        c.emit_op(Op::THROW, 0);
        let catch_start = c.current_offset();
        patch_try_table_two(c, typed_off, catchall_off, catch_start);

        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, caught, 0);
    });
    assert_eq!(r.as_i32(), 88);
}
