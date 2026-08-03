//! Pre-migration performance baseline harness for the dynamic-runtime
//! refactor. Captures current ns/op numbers for operations that will
//! change representation during Phase D migrations.
//!
//! Purpose: establish a reference point for `Phase F.6` to verify that
//! Vybe-VM-path performance after the migration stays within 2× of
//! these numbers. Without a committed baseline we have no way to
//! notice a 5× regression until users report it.
//!
//! These tests are marked `#[ignore]` by default so they don't run in
//! the normal test suite — invoke with `cargo test -p vybe_runtime
//! --test perf_baseline -- --ignored --nocapture` to capture
//! measurements.
//!
//! See `dynamicruntime_support.md` Phase B0.3.

use std::time::Instant;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

const ITERATIONS: usize = 100_000;

/// Measure the wall-clock time to execute `body` once inside a VM,
/// repeating the body logic `ITERATIONS` times inside the bytecode
/// itself (so we pay VM-startup cost exactly once, and the reported
/// per-op time reflects the hot loop, not the setup).
///
/// Returns total duration; divide by ITERATIONS for ns/op.
fn run_and_time(label: &str, emit: impl FnOnce(&mut Chunk)) -> f64 {
    let mut chunk = Chunk::new("<baseline>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let start = Instant::now();
    let _ = vm.run(vec![chunk]).expect("baseline chunk failed");
    let elapsed = start.elapsed();
    let ns_per_op = elapsed.as_nanos() as f64 / ITERATIONS as f64;
    println!(
        "  {:40} {:>8.1} ns/op   ({:.2} ms total)",
        label,
        ns_per_op,
        elapsed.as_secs_f64() * 1000.0
    );
    ns_per_op
}

/// Slot the loop uses for its iteration counter. Body code must use
/// a different slot to avoid clobbering.
const LOOP_COUNTER_SLOT: u16 = 7;

/// Emit a structured WASM counter-driven loop that runs `body`
/// `ITERATIONS` times. The body owns slots 0..=6; the loop owns slot 7.
fn emit_structured_counter_loop(chunk: &mut Chunk, mut body: impl FnMut(&mut Chunk)) {
    chunk.local_count = chunk.local_count.max(LOOP_COUNTER_SLOT + 1);

    let iter_const = chunk.add_constant(Value::I32(ITERATIONS as i32));
    chunk.emit_op_u16(Op::CONST, iter_const, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LOOP_COUNTER_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    body(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, LOOP_COUNTER_SLOT, 0);
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, one, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LOOP_COUNTER_SLOT, 0);
    chunk.emit_br_if(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);
}

/// Baseline table — writes a markdown snapshot to stdout with
/// `--nocapture`. Reviewer copies the output into
/// `docs/perf_baseline_pre_dynamic_runtime.md`.
#[test]
#[ignore = "perf baseline — invoke with --ignored --nocapture to capture numbers"]
fn capture_pre_migration_baseline() {
    println!();
    println!("## Pre-migration baseline ({} iters)", ITERATIONS);
    println!();
    println!("| Operation | ns/op |");
    println!("|---|---:|");

    // ── Array ops (current VM opcodes — will go away in Phase E) ──

    let push_get = run_and_time("vybe:js-array.push + length", |chunk| {
        chunk.emit_array_new_fixed(0, 0, 0);
        let arr_slot = 0;
        chunk.local_count = chunk.local_count.max(1);
        chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, 0);
        // One-time import setup (chunks[0]).
        let push_idx = chunk.add_import("vybe:js-array", "push");

        emit_structured_counter_loop(chunk, |c| {
            c.emit_op_u16(Op::LOCAL_GET, arr_slot, 0);
            let v = c.add_constant(Value::I32(42));
            c.emit_op_u16(Op::CONST, v, 0);
            c.emit_op_u16(Op::CALL_IMPORT, push_idx, 0);
            c.emit(2u8, 0);
            c.emit_op(Op::DROP, 0);
        });
    });
    println!("| `vybe:js-array.push` (import) | {:.1} |", push_get);

    let get_read = run_and_time("array.get (pre-populated)", |chunk| {
        // Pre-populate with one element
        let v = chunk.add_constant(Value::I32(7));
        chunk.emit_op_u16(Op::CONST, v, 0);
        chunk.emit_array_new_fixed(0, 1, 0);
        let arr_slot = 0;
        chunk.local_count = chunk.local_count.max(1);
        chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, 0);

        emit_structured_counter_loop(chunk, |c| {
            c.emit_op_u16(Op::LOCAL_GET, arr_slot, 0);
            let zero = c.add_constant(Value::I32(0));
            c.emit_op_u16(Op::CONST, zero, 0);
            c.emit_op(Op::ARRAY_GET, 0);
            c.emit_op(Op::DROP, 0);
        });
    });
    println!("| `ARRAY_GET` (opcode) | {:.1} |", get_read);

    // ── Struct ops (pure spec GC — our polyfill-path baseline) ──

    let struct_get = run_and_time("struct.get (one field)", |chunk| {
        // Build a simple struct, read a field repeatedly.
        chunk.emit_struct_new(0, 0, 0);
        let obj_slot = 0;
        chunk.local_count = chunk.local_count.max(1);
        chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, 0);

        // Stamp a field once
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, 0);
        let v = chunk.add_constant(Value::I32(99));
        chunk.emit_op_u16(Op::CONST, v, 0);
        let field_name = chunk.add_constant(Value::String("x".into()));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, field_name, 0);
        chunk.emit_op(Op::DROP, 0);

        emit_structured_counter_loop(chunk, |c| {
            c.emit_op_u16(Op::LOCAL_GET, obj_slot, 0);
            let fk = c.add_constant(Value::String("x".into()));
            c.emit_struct_field_op(Op::STRUCT_GET, 0, fk, 0);
            c.emit_op(Op::DROP, 0);
        });
    });
    println!("| `STRUCT_GET` (opcode) | {:.1} |", struct_get);

    // ── Import baseline: wasm:js-string.concat ──
    // Already goes through the import path today; establishes the
    // ceiling we're aiming for when collection ops are imports.
    //
    // Note: CALL_IMPORT needs a fully-registered VM environment with
    // the import registered. Skipping for now — too much harness
    // overhead to set up; a real benchmark will live in a later
    // iteration when we have a standard bench fixture.

    println!();
    println!(
        "_Generated via `cargo test -p vybe_runtime --test perf_baseline -- --ignored --nocapture`_"
    );
}
