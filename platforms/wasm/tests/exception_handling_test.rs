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
use vybe_platform_wasm as wasm;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

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

/// Marks an open try region. Opaque: a test should never spell the clause
/// encoding out, which is exactly how the `labelidx`-as-byte-offset defect
/// stayed invisible to this file.
#[derive(Clone, Copy)]
struct TryTok;

/// Open a spec try region: the HANDLER BLOCK, then a one-clause `try_table`
/// naming it as `labelidx 0`.
///
/// Routed through `Chunk::emit_try_table_clauses` — the single source of truth
/// — rather than re-emitting the bytes here.
fn emit_try_table_start(c: &mut Chunk, kind: u8, tag: u16) -> TryTok {
    // The handler block's result arity is what the clause delivers: one payload
    // value for `catch`, nothing for `catch_all` (spec: no values pushed).
    let arity = if kind == KIND_CATCH { 1 } else { 0 };
    c.emit_block_typed(0, arity);
    c.emit_try_table_clauses(0, 0, &[(kind, tag, 0)], 0);
    TryTok
}

/// Begin this region's handler. Nothing is patched — closing the handler block
/// is what places the handler, because its `end` IS the clause's branch target.
fn patch_try_table(c: &mut Chunk, _tok: TryTok) {
    c.emit_op(Op::END, 0); // close the try_table body
    c.emit_op(Op::END, 0); // close the handler block
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
        c.emit_string_const("boom", 0);
        emit_throw(c, tag);
    });
    assert!(e.contains("boom"));
}

// ── THROW_REF — spec: operand must be an exnref ──────────────────────────

#[test]
fn throw_ref_of_non_exnref_traps() {
    let e = run_err(|c| {
        c.emit_string_const("ref-throw", 0);
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
        let off = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_string_const("oops", 0);
        emit_throw(c, tag);
        patch_try_table(c, off);

        c.emit_op(Op::DROP, 0); // drop delivered payload
        c.emit_i32_const(99, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn rethrow_in_inner_handler_is_caught_by_outer_handler() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let outer = emit_try_table_start(c, KIND_CATCH, tag);
        let inner = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_string_const("nested", 0);
        emit_throw(c, tag);
        patch_try_table(c, inner);
        emit_rethrow(c, 0);
        patch_try_table(c, outer);
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(77, 0);
    });
    assert_eq!(r.as_i32(), 77);
}

#[test]
fn delegate_in_inner_handler_is_caught_by_outer_handler() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let outer = emit_try_table_start(c, KIND_CATCH, tag);
        let inner = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_string_const("delegated", 0);
        emit_throw(c, tag);
        patch_try_table(c, inner);
        emit_delegate(c, 0);
        patch_try_table(c, outer);
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(91, 0);
    });
    assert_eq!(r.as_i32(), 91);
}

#[test]
fn delegate_depth_skips_enclosing_handler() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let o = emit_try_table_start(c, KIND_CATCH, tag);
        let m = emit_try_table_start(c, KIND_CATCH, tag);
        let i = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_string_const("skip-one", 0);
        emit_throw(c, tag);
        patch_try_table(c, i);
        emit_delegate(c, 1);
        patch_try_table(c, m);

        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(222, 0);
        patch_try_table(c, o);

        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(111, 0);
    });
    assert_eq!(r.as_i32(), 111);
}

#[test]
fn try_table_thrown_payload_available_in_catch_handler() {
    let r = run(|c| {
        let tag = lang_tag(c);
        let off = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_i32_const(42, 0);
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
        let off = emit_try_table_start(c, KIND_CATCH, tag);
        c.emit_i32_const(7, 0);
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
        let off = emit_try_table_start(c, KIND_CATCH, t);
        c.emit_string_const("ValueError: bad", 0);
        emit_throw(c, t);
        patch_try_table(c, off);

        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(1, 0);
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

    let off = emit_try_table_start(&mut chunk, KIND_CATCH, t_a);
    chunk.emit_string_const("TypeError: baited", 0);
    emit_throw(&mut chunk, t_b);
    patch_try_table(&mut chunk, off);
    // handler (never reached — thrown tag differs)
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(0, 0);
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
        c.emit_i32_const(77, 0); // catch_all pushes nothing
        let outer = emit_try_table_start(c, KIND_CATCH_ALL, 0);
        let inner = emit_try_table_start(c, KIND_CATCH, t_a);
        c.emit_string_const("wrong-tag", 0);
        emit_throw(c, t_b);
        patch_try_table(c, inner);
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(77, 0);
        c.emit_op(Op::RETURN, 0);
        patch_try_table(c, outer);
        // catch_all handler: no payload pushed — sentinel is TOS
    });
    assert_eq!(r.as_i32(), 77);
}


// ── WRITER: the STANDARD sections, not the `vybe` custom one ──────────────
//
// Everything above exercises the READER. Nothing exercised the writer's spec
// output, and it could not: `write_wasm` embeds the original bytecode in a
// `vybe` custom section and `read_wasm` returns that verbatim when it is
// present, so a write→read round-trip hands back the bytes it started with.
// The standard sections were never decoded by anything.
//
// Under that blind spot the writer was wrong in four ways:
//
//   * every `catch`/`catch_ref` clause was written with tagidx 0, and so was
//     every `throw` — a module with several distinct tags serialized into one
//     where they were all the same tag, which is the entire matching rule of
//     `try_table`;
//   * `throw_ref` was written as opcode 0x08 (`throw`) plus a tagidx: a
//     DIFFERENT instruction, which ignores the exnref operand and raises a
//     fresh exception under the shared tag;
//   * a `try_table` blocktype was written as void-or-externref, so one that
//     takes operands or yields several values changed type in transit.
//
// These decode the standard sections back and assert on the RESULT. They stop
// short of running it: the emitted module declares the runtime's host imports,
// which a bare `VM` has no bindings for.

/// Drop the `vybe` custom section so `read_wasm` must decode the real thing.
fn strip_vybe_section(bytes: &[u8]) -> Vec<u8> {
    let mut out = bytes[..8].to_vec();
    let mut i = 8;
    while i < bytes.len() {
        let start = i;
        let id = bytes[i];
        i += 1;
        let mut size = 0usize;
        let mut shift = 0;
        loop {
            let b = bytes[i];
            i += 1;
            size |= ((b & 0x7f) as usize) << shift;
            if b & 0x80 == 0 {
                break;
            }
            shift += 7;
        }
        let body = &bytes[i..i + size];
        i += size;
        let is_vybe = id == 0 && {
            let mut j = 0usize;
            let mut nlen = 0usize;
            let mut sh = 0;
            loop {
                let b = body[j];
                j += 1;
                nlen |= ((b & 0x7f) as usize) << sh;
                if b & 0x80 == 0 {
                    break;
                }
                sh += 7;
            }
            body.get(j..j + nlen) == Some(b"vybe")
        };
        if !is_vybe {
            out.extend_from_slice(&bytes[start..i]);
        }
    }
    out
}

/// Write `chunk` out and decode its STANDARD sections back.
fn writer_roundtrip(chunk: Chunk) -> Vec<Chunk> {
    let bytes = strip_vybe_section(&wasm::write_wasm(&[chunk]));
    wasm::read_wasm(&bytes).expect("standard-section decode failed")
}

/// The chunk the round-trip produced for our function — the last one, after
/// whatever preamble chunks the module carries.
fn user_chunk(chunks: &[Chunk]) -> &Chunk {
    chunks.last().expect("at least one chunk")
}

/// Walk a decoded chunk and collect `(op, first u16 immediate)` for the EH ops:
/// every `throw`'s tag, and every `catch`/`catch_ref` clause's tag.
fn eh_tag_refs(chunk: &Chunk) -> (Vec<u16>, Vec<u16>, usize) {
    let mut throws = Vec::new();
    let mut clauses = Vec::new();
    let mut throw_refs = 0usize;
    let code = &chunk.code;
    let mut ip = 0usize;
    while ip + 3 < code.len() {
        let g = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
        let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
        let Some(op) = Op::decode(g, sub) else {
            ip += 4;
            continue;
        };
        if op == Op::THROW {
            throws.push(((code[ip + 4] as u16) << 8) | code[ip + 5] as u16);
        } else if op == Op::THROW_REF {
            throw_refs += 1;
        } else if op == Op::TRY_TABLE {
            let n = ((code[ip + 6] as usize) << 8) | code[ip + 7] as usize;
            for k in 0..n {
                let base = ip + 8 + k * 5;
                let kind = code[base];
                if kind == 0x00 || kind == 0x01 {
                    clauses.push(((code[base + 1] as u16) << 8) | code[base + 2] as u16);
                }
            }
        }
        ip += wasm::writer::code::opcode_size(op, code, ip);
    }
    (throws, clauses, throw_refs)
}

/// `block join (result …) { block h (result …) { try_table … end unreachable }
/// … }` — the spec-valid shape the compiler emits, so the reader's stack-shape
/// validator accepts it. `handler_arity` is what the clause delivers.
fn build_try_chunk(
    tag_for_clause: u16,
    tag_for_throw: u16,
    join_arity: u8,
    handler_arity: u8,
    payload: &[i32],
    declare: impl FnOnce(&mut Chunk) -> (u16, u16),
) -> Chunk {
    let mut c = Chunk::new("<script>");
    let (clause_tag, throw_tag) = declare(&mut c);
    let clause_tag = if tag_for_clause == u16::MAX {
        clause_tag
    } else {
        tag_for_clause
    };
    let throw_tag = if tag_for_throw == u16::MAX {
        throw_tag
    } else {
        tag_for_throw
    };
    c.emit_block_typed(0, join_arity); // join
    c.emit_block_typed(0, handler_arity); // handler target
    c.emit_try_table_clauses(0, 0, &[(KIND_CATCH, clause_tag, 0)], 0);
    for v in payload {
        c.emit_i32_const(*v, 0);
    }
    c.emit_op(Op::THROW, 0);
    c.emit((throw_tag >> 8) as u8, 0);
    c.emit((throw_tag & 0xff) as u8, 0);
    c.emit_op(Op::END, 0); // close try_table
    c.emit_op(Op::UNREACHABLE, 0); // the body always throws
    c.emit_op(Op::END, 0); // close handler block
    c
}

#[test]
fn writer_keeps_distinct_tags_distinct() {
    // A clause on TagA and a throw of TagB must reference DIFFERENT tags after
    // the round-trip. Collapsing both onto the shared exception tag made the
    // clause match a throw it must never catch.
    let chunk = build_try_chunk(u16::MAX, u16::MAX, 1, 1, &[99], |c| {
        let a = c.declare_exception_tag("TagA", 1);
        let b = c.declare_exception_tag("TagB", 1);
        (a, b)
    });
    let mut chunk = chunk;
    chunk.emit_op(Op::END, 0);
    chunk.emit_op(Op::RETURN, 0);
    let decoded = writer_roundtrip(chunk);
    let (throws, clauses, _) = eh_tag_refs(user_chunk(&decoded));
    assert_eq!(throws.len(), 1, "expected one throw");
    assert_eq!(clauses.len(), 1, "expected one typed clause");
    assert_ne!(
        throws[0], clauses[0],
        "TagA and TagB survived as the SAME tag — the clause would catch a \
         throw it must not match"
    );
}

#[test]
fn writer_keeps_one_tag_one_tag() {
    // The other half: a clause and a throw naming the SAME tag must still
    // agree afterwards. Guards "make them distinct" against over-correcting
    // into "nothing matches".
    let chunk = build_try_chunk(u16::MAX, u16::MAX, 1, 1, &[42], |c| {
        let a = c.declare_exception_tag("TagA", 1);
        let _b = c.declare_exception_tag("TagB", 1);
        (a, a)
    });
    let mut chunk = chunk;
    chunk.emit_op(Op::END, 0);
    chunk.emit_op(Op::RETURN, 0);
    let decoded = writer_roundtrip(chunk);
    let (throws, clauses, _) = eh_tag_refs(user_chunk(&decoded));
    assert_eq!(
        throws[0], clauses[0],
        "one tag came back as two — the clause would no longer match its throw"
    );
}

#[test]
fn writer_preserves_a_two_ary_tag() {
    // A 2-ary tag needs a functype of its own in the tag section; borrowing
    // the one-param exception type cannot express it, and the handler block
    // that receives both values needs a real typeidx blocktype.
    let chunk = build_try_chunk(u16::MAX, u16::MAX, 1, 2, &[3, 4], |c| {
        let t = c.declare_exception_tag("Pair", 2);
        (t, t)
    });
    let mut chunk = chunk;
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::END, 0);
    chunk.emit_op(Op::RETURN, 0);
    let decoded = writer_roundtrip(chunk);
    let c = user_chunk(&decoded);
    let (throws, clauses, _) = eh_tag_refs(c);
    assert_eq!(throws[0], clauses[0]);
    assert_eq!(
        c.tags[throws[0] as usize].arity, 2,
        "the tag's payload arity did not survive"
    );
}

#[test]
fn writer_emits_throw_ref_not_throw() {
    // `throw_ref` is opcode 0x0A and takes NO tag immediate. Written as 0x08
    // it decodes back as `throw`, which never inspects the exnref operand.
    let mut c = Chunk::new("<script>");
    c.emit_i32_const(1, 0);
    c.emit_op(Op::THROW_REF, 0);
    c.emit_op(Op::RETURN, 0);
    let decoded = writer_roundtrip(c);
    let (throws, _, throw_refs) = eh_tag_refs(user_chunk(&decoded));
    assert_eq!(throw_refs, 1, "throw_ref did not survive as throw_ref");
    assert!(
        throws.is_empty(),
        "throw_ref came back as a plain `throw` — a different instruction"
    );
}
