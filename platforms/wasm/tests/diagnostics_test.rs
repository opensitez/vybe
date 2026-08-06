//! Diagnostic tools for the anyref/ABI migration.
//!
//! Two facilities:
//!   1. **Type recorder** — per-slot `Value`-variant histograms.
//!      Measures how much of the Phase-3 typed-locals refactor can
//!      actually pay off.
//!   2. **WAT emitter** — textual WebAssembly dump of our chunks.
//!      Useful for visually inspecting what we produce without
//!      running through an external disassembler.

use vybe_runtime::value::ValueTag;
use vybe_runtime::{Chunk, Op, VM};
use vybe_platform_wasm::disassembler::{write_wat, write_wat_chunk};

// ── Type recorder ──────────────────────────────────────────────────

#[test]
fn type_recorder_is_off_by_default() {
    let vm = VM::new();
    assert!(vm.type_recorder.is_none());
}

#[test]
fn type_recorder_captures_monomorphic_slot() {
    // A slot that only ever holds an I32 should be recorded as
    // monomorphic(I32). This is the base-case signal for Phase 3:
    // if a slot is mono-I32, it can become a native i32 local.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_i32_const(42, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_types(true);
    vm.run(vec![chunk]).unwrap();
    let rec = vm.take_type_record().unwrap();

    let chunk_obs = &rec.slots()[0];
    assert_eq!(chunk_obs[0].monomorphic_tag(), Some(ValueTag::I32));
    assert!(
        chunk_obs[0].counts[ValueTag::I32.as_usize()] >= 2,
        "slot should see both LOCAL_SET and LOCAL_GET traffic"
    );
}

#[test]
fn type_recorder_flags_polymorphic_slot() {
    // A slot that holds I32 then String is polymorphic(2). Phase 3
    // must keep polymorphic slots on the anyref path — the recorder's
    // role is to flag them so we don't pick wrong.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_i32_const(1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_string_const("hi", 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_types(true);
    vm.run(vec![chunk]).unwrap();
    let rec = vm.take_type_record().unwrap();

    let slot = &rec.slots()[0][0];
    assert_eq!(slot.distinct_variants(), 2);
    assert!(slot.monomorphic_tag().is_none());
    assert!(slot.counts[ValueTag::I32.as_usize()] > 0);
    assert!(slot.counts[ValueTag::String.as_usize()] > 0);
}

#[test]
fn type_recorder_summary_reports_monomorphic_percentage() {
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;
    // slot 0: monomorphic I32
    chunk.emit_i32_const(1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    // slot 1: polymorphic I32 then String
    chunk.emit_i32_const(1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_string_const("x", 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_types(true);
    vm.run(vec![chunk]).unwrap();
    let rec = vm.take_type_record().unwrap();
    let s = rec.summary();
    assert_eq!(s.observed_slots, 2);
    assert_eq!(s.monomorphic, 1);
    assert_eq!(s.polymorphic, 1);
    assert!((s.mono_percent() - 50.0).abs() < 0.1);
}

#[test]
fn type_recorder_off_means_zero_overhead() {
    // With recording disabled, the VM should keep the `type_recorder`
    // field at `None` — no allocation path taken on LOCAL_SET.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_i32_const(42, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    // deliberately DO NOT call record_types(true)
    vm.run(vec![chunk]).unwrap();
    assert!(vm.type_recorder.is_none(), "recorder should stay disabled");
}

// ── WAT emitter ────────────────────────────────────────────────────

#[test]
fn wat_emits_module_scaffold() {
    let mut chunk = Chunk::new("main");
    chunk.arity = 1;
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wat = write_wat(&vec![chunk]);
    assert!(wat.starts_with("(module\n"));
    assert!(wat.contains("(func "));
    assert!(wat.contains("(param $p0 externref)"));
    assert!(wat.contains("(result externref)"));
    assert!(wat.contains("local.get"));
    assert!(wat.trim_end().ends_with(')'));
}

#[test]
fn wat_renders_const_comment_for_known_values() {
    // Spec consts carry their literal inline (`i32.const 42`), so the WAT
    // is readable with no constant-table cross-referencing. (The legacy
    // pool-indexed CONST needed a `;; 42i32` comment for this; retired.)
    let mut chunk = Chunk::new("main");
    chunk.arity = 0;
    chunk.emit_i32_const(42, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wat = write_wat_chunk(&chunk);
    assert!(
        wat.contains("i32.const 42") || wat.contains("42"),
        "expected the literal value inline in the WAT, got:\n{wat}"
    );
}

#[test]
fn wat_indents_block_bodies() {
    // Nested `block ... end` should show indentation so control flow
    // is visually discernible at a glance.
    let mut chunk = Chunk::new("main");
    chunk.arity = 0;
    let _bp = chunk.emit_block(0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::END, 0);
    chunk.patch_block(_bp);
    chunk.emit_op(Op::RETURN, 0);

    let wat = write_wat_chunk(&chunk);
    // A `null` inside the block should be more indented than the
    // surrounding `block` / `end` — verify by looking at leading
    // whitespace on the line containing `ref.null` (opcode name).
    let lines: Vec<&str> = wat.lines().collect();
    let block_indent = lines
        .iter()
        .find(|l| l.trim_start().starts_with("block"))
        .map(|l| l.len() - l.trim_start().len())
        .unwrap_or(0);
    let null_indent = lines
        .iter()
        .find(|l| l.trim_start().starts_with("ref.null"))
        .map(|l| l.len() - l.trim_start().len())
        .unwrap_or(0);
    assert!(
        null_indent > block_indent,
        "inner `ref.null` should be indented past outer `block` — got {null_indent} vs {block_indent}\n{wat}"
    );
}
