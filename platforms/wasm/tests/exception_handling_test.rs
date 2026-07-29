//! Tests for WASM exception handling: THROW (0x08), THROW_REF (0x0A), TRY_TABLE (0x1F).
//!
//! Semantics follow the exception-handling proposal (final/exnref phase):
//! tags are entities, `throw <tagidx>` packages the payload with its tag,
//! catch clauses match by TAG IDENTITY only. Full semantic coverage lives in
//! `vybe_runtime/tests/wasm_eh_compliance_test.rs`; this file covers the
//! platform layer: binary decode of the EH ops plus the internal encoding
//! round-trip.
//!
//! Internal fixed-width TRY_TABLE layout:
//!   [op:2][clause_count:1][ (kind:1)(tag:2)(offset:2) × N ]
//!   kind: 0=catch 1=catch_ref 2=catch_all 3=catch_all_ref
//!   catch_ip = ip-after-this-clause + offset (big-endian u16).

use std::sync::Arc;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};
use vybe_platform_wasm as wasm;

const KIND_CATCH: u8 = 0;
const KIND_CATCH_ALL: u8 = 2;

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

// Legacy (pre-3.0) exception handling is not supported: it must be rejected
// with a clear error — never silently mis-decoded into a broken chunk.

#[test]
fn legacy_rethrow_rejected_with_clear_error() {
    let bytes = standard_eh_module(&[
        0x09, 0x00, // rethrow 0 (legacy)
    ]);
    let err = wasm::read_wasm(&bytes).expect_err("legacy rethrow must be rejected");
    assert!(err.contains("legacy exception-handling"), "got: {err}");
}

#[test]
fn legacy_delegate_rejected_with_clear_error() {
    let bytes = standard_eh_module(&[
        0x18, 0x00, // delegate 0 (legacy)
    ]);
    let err = wasm::read_wasm(&bytes).expect_err("legacy delegate must be rejected");
    assert!(err.contains("legacy exception-handling"), "got: {err}");
}

#[test]
fn legacy_try_catch_block_rejected_with_clear_error() {
    // try (0x06) void … catch (0x07) tag0 … end — the legacy block form.
    let bytes = standard_eh_module(&[
        0x06, 0x40, // try (void blocktype)
        0x07, 0x00, // catch tag0
        0x0B, // end
    ]);
    let err = wasm::read_wasm(&bytes).expect_err("legacy try/catch must be rejected");
    assert!(err.contains("legacy exception-handling"), "got: {err}");
}

/// The language-exception tag every emitter throw uses (imported by name so
/// all chunks resolve to the same entity — the single-tag toolchain design).
fn lang_tag(c: &mut Chunk) -> u16 {
    c.import_exception_tag("vybe:exception", 1)
}

/// Spec `throw <tagidx>` with the payload already on the stack.
fn emit_throw(c: &mut Chunk, tag: u16) {
    c.emit_op(Op::THROW, 0);
    c.emit((tag >> 8) as u8, 0);
    c.emit((tag & 0xff) as u8, 0);
}

/// Emit TRY_TABLE with one clause. Returns the offset_pos to patch.
fn emit_try_table_start(c: &mut Chunk, kind: u8, tag: u16) -> usize {
    c.emit_op(Op::TRY_TABLE, 0);
    c.emit(1, 0); // clause_count = 1
    c.emit(kind, 0);
    c.emit((tag >> 8) as u8, 0);
    c.emit((tag & 0xff) as u8, 0);
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

// ── Import + RUN foreign modules (round-trip through read_wasm) ────────────
// These build real `.wasm` binaries with a tag section and structured EH, then
// import them via `read_wasm` and EXECUTE the decoded chunk — proving the
// reader lowers foreign EH to something the VM actually runs correctly.

/// Build a module: type0 = `[i32]->[]` (the tag's type), type1 = `[]->[i32]`
/// (the function), one exception tag, and one function whose body is `body_ops`
/// (a trailing `end` is appended). Returns the raw `.wasm` bytes.
fn eh_import_module(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    // Type section: type0 [i32]->[], type1 []->[i32].
    let mut types = vec![0x02];
    types.extend_from_slice(&[0x60, 0x01, 0x7F, 0x00]); // [i32] -> []
    types.extend_from_slice(&[0x60, 0x00, 0x01, 0x7F]); // [] -> [i32]
    push_section(&mut out, 1, &types);
    // Function section: func0 : type1.
    push_section(&mut out, 3, &[0x01, 0x01]);
    // Tag section (id 13): one tag, attribute 0x00, type 0.
    push_section(&mut out, 13, &[0x01, 0x00, 0x00]);

    // Code section.
    let mut body = Vec::new();
    body.push(0x00); // 0 local groups
    body.extend_from_slice(body_ops);
    body.push(0x0B); // function end
    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);
    out
}

/// Run the first non-script decoded chunk as a standalone program.
fn import_and_run(wasm_bytes: &[u8]) -> Result<Value, String> {
    let mut chunks = wasm::read_wasm(wasm_bytes)?;
    // chunk[0] is the synthetic script; chunk[1] is the imported function.
    let func = chunks.remove(1);
    VM::new().run(vec![func]).map_err(|e| e.to_string())
}

#[test]
fn import_new_form_try_table_catches_and_returns_payload() {
    // (func (result i32)
    //   (block $h (result i32)
    //     (try_table (catch $e $h)   ;; catch → label 1 ($h)
    //       i32.const 99
    //       throw $e)
    //     i32.const 0))              ;; normal path (unreached): $h = value
    let body = [
        0x02, 0x7F, // block $h (result i32)
        0x1F, 0x40, 0x01, 0x00, 0x00, 0x01, // try_table void, catch tag0 label1
        0x41, 0xE3, 0x00, // i32.const 99
        0x08, 0x00, // throw tag0
        0x0B, // end try_table
        0x41, 0x00, // i32.const 0
        0x0B, // end block $h
    ];
    let wasm = eh_import_module(&body);
    let result = import_and_run(&wasm).expect("caught new-form try_table must run");
    assert_eq!(
        format!("{result}"),
        "99",
        "catch must deliver the payload 99"
    );
}

#[test]
fn import_new_form_uncaught_throw_propagates() {
    // A throw with a try_table whose clause does NOT match... simplest: throw
    // with no enclosing try_table at all → escapes as a runtime error.
    let body = [
        0x41, 0xE3, 0x00, // i32.const 99
        0x08, 0x00, // throw tag0 — uncaught
        0x41, 0x00, // i32.const 0 (unreached; keeps the func result-typed)
    ];
    let wasm = eh_import_module(&body);
    let err = import_and_run(&wasm).expect_err("uncaught throw must surface as an error");
    assert!(!err.is_empty());
}

// ── THROW — uncaught ──────────────────────────────────────────────────────

#[test]
fn throw_uncaught_propagates_as_error() {
    let e = run_err(|c| {
        let tag = lang_tag(c);
        let k = c.add_constant(Value::String(Arc::from("boom")));
        c.emit_op_u16(Op::CONST, k, 0);
        emit_throw(c, tag);
    });
    assert!(e.contains("boom"));
}

// ── THROW_REF — spec: operand must be an exnref ──────────────────────────

#[test]
fn throw_ref_of_non_exnref_traps() {
    let e = run_err(|c| {
        let k = c.add_constant(Value::String(Arc::from("ref-throw")));
        c.emit_op_u16(Op::CONST, k, 0);
        c.emit_op(Op::THROW_REF, 0);
    });
    assert!(
        e.contains("exnref"),
        "throw_ref must reject a non-exnref operand, got: {e}"
    );
}

// ── TRY_TABLE + THROW — the language tag ─────────────────────────────────

#[test]
fn catch_lang_tag_intercepts_throw() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let err = c.add_constant(Value::String(Arc::from("oops")));
        let ok = c.add_constant(Value::I32(99));

        let off = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_op_u16(Op::CONST, err, 0);
        emit_throw(c, tag);
        patch_try_table(c, off);

        c.emit_op(Op::DROP, 0); // drop delivered payload
        c.emit_op_u16(Op::CONST, ok, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn rethrow_in_inner_handler_is_caught_by_outer_handler() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let err = c.add_constant(Value::String(Arc::from("nested")));
        let ok = c.add_constant(Value::I32(77));

        let outer = emit_try_table_start(c, KIND_CATCH, tag);
        let inner = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_op_u16(Op::CONST, err, 0);
        emit_throw(c, tag);
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
        let tag = lang_tag(c);
        let err = c.add_constant(Value::String(Arc::from("delegated")));
        let ok = c.add_constant(Value::I32(91));

        let outer = emit_try_table_start(c, KIND_CATCH, tag);
        let inner = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_op_u16(Op::CONST, err, 0);
        emit_throw(c, tag);
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
        let tag = lang_tag(c);
        let err = c.add_constant(Value::String(Arc::from("skip-one")));
        let outer = c.add_constant(Value::I32(111));
        let middle = c.add_constant(Value::I32(222));

        let o = emit_try_table_start(c, KIND_CATCH, tag);
        let m = emit_try_table_start(c, KIND_CATCH, tag);
        let i = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_op_u16(Op::CONST, err, 0);
        emit_throw(c, tag);
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
fn try_table_thrown_payload_available_in_catch_handler() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let thrown = c.add_constant(Value::I32(42));

        let off = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_op_u16(Op::CONST, thrown, 0);
        emit_throw(c, tag);
        patch_try_table(c, off);
        // handler: the tag's payload (the thrown value) is on the stack
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn try_table_no_throw_falls_through() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let ok = c.add_constant(Value::I32(7));

        let off = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_op_u16(Op::CONST, ok, 0);
        // run() adds RETURN — no throw, handler never runs
        patch_try_table(c, off);
    });
    assert_eq!(r.as_i32(), 7);
}

// ── Typed catch — TAG IDENTITY (spec), never payload inspection ──────────

#[test]
fn catch_matches_by_tag_identity() {
    let r = run(|c| {
        let t = c.declare_exception_tag("ValueError", 1);
        let err = c.add_constant(Value::String(Arc::from("ValueError: bad")));
        let caught = c.add_constant(Value::I32(1));

        let off = emit_try_table_start(c, KIND_CATCH, t);
        c.emit_op_u16(Op::CONST, err, 0);
        emit_throw(c, t);
        patch_try_table(c, off);

        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, caught, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn catch_for_different_tag_does_not_match() {
    // A clause for tag A must not catch a throw of tag B — even when the
    // payload STRING contains A's debug name (identity, not inspection).
    let mut chunk = Chunk::new("<script>");
    let t_a = chunk.declare_exception_tag("TypeError", 1);
    let t_b = chunk.declare_exception_tag("ValueError", 1);

    let err_str = chunk.add_constant(Value::String(Arc::from("TypeError: baited")));
    let fallback = chunk.add_constant(Value::I32(0));

    let off = emit_try_table_start(&mut chunk, KIND_CATCH, t_a);
    chunk.emit_op_u16(Op::CONST, err_str, 0);
    emit_throw(&mut chunk, t_b);
    patch_try_table(&mut chunk, off);
    // handler (never reached — thrown tag differs)
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::CONST, fallback, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(
        err.contains("ValueError"),
        "the tag-B throw must escape uncaught, got: {err}"
    );
}

#[test]
fn nonmatching_typed_clause_falls_through_to_enclosing_catch_all() {
    let r = run(|c| {
        let t_a = c.declare_exception_tag("TypeError", 1);
        let t_b = c.declare_exception_tag("ValueError", 1);
        let err = c.add_constant(Value::String(Arc::from("wrong-tag")));
        let caught = c.add_constant(Value::I32(77));
        let sentinel = c.add_constant(Value::I32(77));

        c.emit_op_u16(Op::CONST, sentinel, 0); // catch_all pushes nothing
        let outer = emit_try_table_start(c, KIND_CATCH_ALL, 0);
        let inner = emit_try_table_start(c, KIND_CATCH, t_a);
        c.emit_op_u16(Op::CONST, err, 0);
        emit_throw(c, t_b);
        patch_try_table(c, inner);
        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::CONST, caught, 0);
        c.emit_op(Op::RETURN, 0);
        patch_try_table(c, outer);
        // catch_all handler: no payload pushed — sentinel is TOS
    });
    assert_eq!(r.as_i32(), 77);
}
