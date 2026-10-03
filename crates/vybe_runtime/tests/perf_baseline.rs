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

use std::sync::{Arc, Mutex};
use std::time::Instant;
use vybe_runtime::value::{
    ArrayBufferState, Object, ObjectKind, TypedArrayState, TypedElemKind, Value,
};
use vybe_runtime::{Chunk, ImportTarget, Op, VM};

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

    chunk.emit_i32_const(ITERATIONS as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LOOP_COUNTER_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    body(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, LOOP_COUNTER_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op_u16(Op::LOCAL_TEE, LOOP_COUNTER_SLOT, 0);
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

    println!("| `vybe:js-array.push` (retired import) | skipped |");

    let get_read = run_and_time("array.get (pre-populated)", |chunk| {
        // Pre-populate with one element
        chunk.emit_i32_const(7, 0);
        chunk.emit_array_new_fixed(0, 1, 0);
        let arr_slot = 0;
        chunk.local_count = chunk.local_count.max(1);
        chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, 0);

        emit_structured_counter_loop(chunk, |c| {
            c.emit_op_u16(Op::LOCAL_GET, arr_slot, 0);
            c.emit_i32_const(0, 0);
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
        chunk.emit_i32_const(99, 0);
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

/// Runtime-only fixture for the cached fread bridge identified in
/// `documentation/vega_runtime_hotpaths.md`.
///
/// Shape:
/// managed array byte source -> ARRAY_GET -> I32_STORE8 -> linear memory 0.
///
/// This deliberately avoids C compilation and `.wasm` loading so we can time
/// the VM path even while the emitted workload module has a loader-validation
/// blocker.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn managed_array_to_linear_i32_store8_bridge_counters() {
    const LEN: usize = 256;
    const REPEATS: usize = 128;
    const TOTAL: usize = LEN * REPEATS;
    const ARR_GLOBAL: &str = "__perf_managed_array_source";

    const ARR_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const I_SLOT: u16 = 3;
    const BYTE_SLOT: u16 = 4;

    let mut setup = Chunk::new("<managed-array-source-setup>");
    for i in 0..LEN {
        setup.emit_i32_const((i & 0xff) as i32, 0);
    }
    setup.emit_array_new_fixed(0, LEN as u16, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let source = vm.run(vec![setup]).expect("bridge source setup failed");
    vm.set_global(ARR_GLOBAL, source);

    let mut chunk = Chunk::new("<managed-array-to-linear-bridge>");
    chunk.local_count = 5;

    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);

    let arr_global_idx = chunk.add_constant(Value::String(ARR_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, arr_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const((LEN - 1) as i32, 0);
    chunk.emit_op(Op::I32_AND, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    chunk.emit_op(Op::I32_STORE8, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);

    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("bridge fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "managed array -> linear bridge: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
}

/// Runtime-only fixture for the exact cached fread bridge shape from
/// `documentation/vega_runtime_hotpaths.md`:
/// managed array + source offset + i -> ARRAY_GET -> temporary locals ->
/// I32_STORE8 -> linear memory 0.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn managed_array_offset_to_linear_i32_store8_bridge_counters() {
    const SRC_OFFSET: usize = 7;
    const TOTAL: usize = 32 * 1024;
    const SOURCE_LEN: usize = SRC_OFFSET + TOTAL;
    const ARR_GLOBAL: &str = "__perf_managed_array_offset_source";

    const ARR_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const I_SLOT: u16 = 3;
    const BYTE_SLOT: u16 = 4;
    const STORE_VALUE_SLOT: u16 = 5;
    const STORE_ADDR_SLOT: u16 = 6;
    const SRC_OFFSET_SLOT: u16 = 7;

    let mut setup = Chunk::new("<managed-array-offset-source-setup>");
    for i in 0..SOURCE_LEN {
        setup.emit_i32_const((i & 0xff) as i32, 0);
    }
    setup.emit_array_new_fixed(0, SOURCE_LEN as u16, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let source = vm.run(vec![setup]).expect("offset bridge source setup failed");
    vm.set_global(ARR_GLOBAL, source);

    let mut chunk = Chunk::new("<managed-array-offset-to-linear-bridge>");
    chunk.local_count = 8;

    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);

    let arr_global_idx = chunk.add_constant(Value::String(ARR_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, arr_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i32_const(SRC_OFFSET as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_OFFSET_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, SRC_OFFSET_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, STORE_VALUE_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, STORE_ADDR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, STORE_ADDR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, STORE_VALUE_SLOT, 0);
    chunk.emit_op(Op::I32_STORE8, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, STORE_VALUE_SLOT, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);

    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("offset bridge fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((SRC_OFFSET + TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "managed array offset -> linear bridge: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
    assert_eq!(counters.super_managed_to_linear, 1);
    assert_eq!(counters.array_get, TOTAL as u64);
    assert_eq!(counters.i32_store8, TOTAL as u64);
}

/// Same bridge shape as `managed_array_offset_to_linear_i32_store8_bridge_counters`,
/// but with a packed Uint8Array source. This exercises the runtime fast path
/// that can copy source bytes directly instead of converting a `Vec<Value>`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn typed_array_offset_to_linear_i32_store8_bridge_counters() {
    const SRC_OFFSET: usize = 7;
    const TOTAL: usize = 32 * 1024;
    const SOURCE_LEN: usize = SRC_OFFSET + TOTAL;
    const ARR_GLOBAL: &str = "__perf_typed_array_offset_source";

    const ARR_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const I_SLOT: u16 = 3;
    const BYTE_SLOT: u16 = 4;
    const STORE_VALUE_SLOT: u16 = 5;
    const STORE_ADDR_SLOT: u16 = 6;
    const SRC_OFFSET_SLOT: u16 = 7;

    let bytes = Arc::new(Mutex::new(
        (0..SOURCE_LEN).map(|i| (i & 0xff) as u8).collect::<Vec<_>>(),
    ));
    let mut buffer_obj = Object::new();
    buffer_obj.kind = ObjectKind::ArrayBuffer(ArrayBufferState {
        bytes: bytes.clone(),
        max_byte_length: SOURCE_LEN,
        resizable: false,
        detached: false,
        shared: false,
    });
    let buffer_obj = Arc::new(Mutex::new(buffer_obj));
    let mut typed_obj = Object::new();
    typed_obj.kind = ObjectKind::TypedArray(TypedArrayState {
        elem: TypedElemKind::U8,
        buffer: bytes,
        buffer_obj,
        byte_offset: 0,
        length: SOURCE_LEN,
    });

    let mut vm = VM::new();
    vm.set_global(ARR_GLOBAL, Value::Object(Arc::new(Mutex::new(typed_obj))));

    let mut chunk = Chunk::new("<typed-array-offset-to-linear-bridge>");
    chunk.local_count = 8;

    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);

    let arr_global_idx = chunk.add_constant(Value::String(ARR_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, arr_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i32_const(SRC_OFFSET as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_OFFSET_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, SRC_OFFSET_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, STORE_VALUE_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, STORE_ADDR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, STORE_ADDR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, STORE_VALUE_SLOT, 0);
    chunk.emit_op(Op::I32_STORE8, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, STORE_VALUE_SLOT, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);

    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("typed offset bridge fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((SRC_OFFSET + TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "typed array offset -> linear bridge: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
    assert_eq!(counters.super_managed_to_linear, 1);
    assert_eq!(counters.array_get, TOTAL as u64);
    assert_eq!(counters.i32_store8, TOTAL as u64);
}

/// Runtime-only fixture for the reverse bridge:
/// linear memory 0 -> I32_LOAD8_U -> ARRAY_SET -> managed array byte dest.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn linear_to_managed_array_i32_load8_bridge_counters() {
    const LEN: usize = 256;
    const REPEATS: usize = 128;
    const TOTAL: usize = LEN * REPEATS;
    const ARR_GLOBAL: &str = "__perf_linear_to_managed_target";

    const INIT_I_SLOT: u16 = 0;
    const ARR_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const I_SLOT: u16 = 3;
    const BYTE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-source-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(0xff, 0);
    init.emit_op(Op::I32_AND, 0);
    init.emit_op(Op::I32_STORE8, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(1, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut setup = Chunk::new("<linear-to-managed-target-setup>");
    for _ in 0..LEN {
        setup.emit_i32_const(0, 0);
    }
    setup.emit_array_new_fixed(0, LEN as u16, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut bridge = Chunk::new("<linear-to-managed-array-bridge>");
    bridge.local_count = 5;

    let arr_global_idx = bridge.add_constant(Value::String(ARR_GLOBAL.into()));
    bridge.emit_op_u32(Op::GLOBAL_GET, arr_global_idx, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    bridge.emit_i32_const(0, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    bridge.emit_i32_const(TOTAL as i32, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    bridge.emit_i32_const(0, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);

    let outer = bridge.emit_block(0);
    let (lp, _loop_start) = bridge.emit_loop_s(0);
    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    bridge.emit_op(Op::I32_GE_U, 0);
    bridge.emit_br_if(1, 0);

    bridge.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_op(Op::I32_ADD, 0);
    bridge.emit_op(Op::I32_LOAD8_U, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    bridge.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_i32_const((LEN - 1) as i32, 0);
    bridge.emit_op(Op::I32_AND, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    bridge.emit_op(Op::ARRAY_SET, 0);

    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_i32_const(1, 0);
    bridge.emit_op(Op::I32_ADD, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    bridge.emit_br(0, 0);
    bridge.emit_end(0);
    bridge.patch_loop(lp);
    bridge.emit_end(0);
    bridge.patch_block(outer);

    bridge.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    bridge.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear source init failed");
    let target = vm
        .run(vec![setup])
        .expect("linear target setup failed");
    vm.set_global(ARR_GLOBAL, target);
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![bridge]).expect("reverse bridge fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear -> managed array bridge: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
}

/// Runtime-only fixture for the offset reverse bridge:
/// linear memory 0 -> I32_LOAD8_U -> ARRAY_SET at dst_offset + i.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_to_managed_array_offset_i32_load8_bridge_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_OFFSET: usize = 16;
    const ARR_GLOBAL: &str = "__perf_linear_to_managed_offset_target";

    const INIT_I_SLOT: u16 = 0;
    const ARR_SLOT: u16 = 0;
    const DST_OFFSET_SLOT: u16 = 1;
    const SRC_SLOT: u16 = 2;
    const LEN_SLOT: u16 = 3;
    const I_SLOT: u16 = 4;
    const BYTE_SLOT: u16 = 5;

    let mut init = Chunk::new("<linear-offset-source-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(0xff, 0);
    init.emit_op(Op::I32_AND, 0);
    init.emit_op(Op::I32_STORE8, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(1, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut setup = Chunk::new("<linear-to-managed-offset-target-setup>");
    for _ in 0..(DST_OFFSET + TOTAL) {
        setup.emit_i32_const(0, 0);
    }
    setup.emit_array_new_fixed(0, (DST_OFFSET + TOTAL) as u16, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut bridge = Chunk::new("<linear-to-managed-array-offset-bridge>");
    bridge.local_count = 6;

    let arr_global_idx = bridge.add_constant(Value::String(ARR_GLOBAL.into()));
    bridge.emit_op_u32(Op::GLOBAL_GET, arr_global_idx, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);
    bridge.emit_i32_const(DST_OFFSET as i32, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, DST_OFFSET_SLOT, 0);
    bridge.emit_i32_const(0, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    bridge.emit_i32_const(TOTAL as i32, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    bridge.emit_i32_const(0, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);

    let outer = bridge.emit_block(0);
    let (lp, _loop_start) = bridge.emit_loop_s(0);
    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    bridge.emit_op(Op::I32_GE_U, 0);
    bridge.emit_br_if(1, 0);

    bridge.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_op(Op::I32_ADD, 0);
    bridge.emit_op(Op::I32_LOAD8_U, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    bridge.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, DST_OFFSET_SLOT, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_op(Op::I32_ADD, 0);
    bridge.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    bridge.emit_op(Op::ARRAY_SET, 0);

    bridge.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    bridge.emit_i32_const(1, 0);
    bridge.emit_op(Op::I32_ADD, 0);
    bridge.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    bridge.emit_br(0, 0);
    bridge.emit_end(0);
    bridge.patch_loop(lp);
    bridge.emit_end(0);
    bridge.patch_block(outer);

    bridge.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    bridge.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear offset source init failed");
    let target = vm
        .run(vec![setup])
        .expect("linear offset target setup failed");
    vm.set_global(ARR_GLOBAL, target);
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![bridge])
        .expect("offset reverse bridge fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear -> managed array offset bridge: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
    assert_eq!(counters.super_linear_to_managed, 1);
    assert_eq!(counters.array_set_dense, TOTAL as u64);
}

/// Runtime-only fixture for pure linear-memory byte copy:
/// memory 0 -> I32_LOAD8_U -> I32_STORE8 -> memory 0.
///
/// This isolates scalar memory/address/branch overhead from managed array
/// overhead in the cached fread bridge fixtures.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn linear_i32_load8_to_i32_store8_copy_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;

    const I_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const COPY_I_SLOT: u16 = 3;
    const BYTE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-byte-copy-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(0xff, 0);
    init.emit_op(Op::I32_AND, 0);
    init.emit_op(Op::I32_STORE8, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(1, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<linear-byte-copy>");
    copy.local_count = 5;
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);

    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::I32_LOAD8_U, 0);
    copy.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    copy.emit_op(Op::I32_STORE8, 0);

    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_i32_const(1, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);
    copy.emit_br(0, 0);
    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);

    copy.emit_i32_const((DST_BASE + TOTAL - 1) as i32, 0);
    copy.emit_op(Op::I32_LOAD8_U, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear byte source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("linear byte copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear byte copy: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
}

/// Runtime-only fixture for repeated word stores:
/// memory 0 <- I32_STORE(value) with `i += 4`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i32_store_fill_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const PATTERN: i32 = 0x1020_3040;

    const DST_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const VALUE_SLOT: u16 = 3;

    let mut chunk = Chunk::new("<linear-i32-store-fill>");
    chunk.local_count = 4;
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(DST_BASE as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i32_const(PATTERN, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(4, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_i32_const((DST_BASE + TOTAL - 4) as i32, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("linear i32 store fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i32.store fill loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_fill, 1);
    assert_eq!(counters.memory_fill, 1);
    assert_eq!(counters.memory_fill_bytes, TOTAL as u64);
}

/// Runtime-only fixture for repeated doubleword stores:
/// memory 0 <- I64_STORE(value) with `i += 8`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i64_store_fill_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const PATTERN: i64 = 0x1020_3040_5060_7080;

    const DST_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const VALUE_SLOT: u16 = 3;

    let mut chunk = Chunk::new("<linear-i64-store-fill>");
    chunk.local_count = 4;
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(DST_BASE as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i64_const(PATTERN, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::I64_STORE, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(8, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_i32_const((DST_BASE + TOTAL - 8) as i32, 0);
    chunk.emit_op(Op::I64_LOAD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("linear i64 store fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I64(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i64.store fill loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_fill, 1);
    assert_eq!(counters.memory_fill, 1);
    assert_eq!(counters.memory_fill_bytes, TOTAL as u64);
}

/// Runtime-only fixture for repeated f32 stores:
/// memory 0 <- F32_STORE(value) with `i += 4`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_f32_store_fill_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const PATTERN: f32 = 13.25;

    const DST_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const VALUE_SLOT: u16 = 3;

    let mut chunk = Chunk::new("<linear-f32-store-fill>");
    chunk.local_count = 4;
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(DST_BASE as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_f32_const(PATTERN, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::F32_STORE, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(4, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_i32_const((DST_BASE + TOTAL - 4) as i32, 0);
    chunk.emit_op(Op::F32_LOAD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("linear f32 store fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::F32(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear f32.store fill loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_fill, 1);
    assert_eq!(counters.memory_fill, 1);
    assert_eq!(counters.memory_fill_bytes, TOTAL as u64);
}

/// Runtime-only fixture for repeated f64 stores:
/// memory 0 <- F64_STORE(value) with `i += 8`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_f64_store_fill_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const PATTERN: f64 = 13.25;

    const DST_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const VALUE_SLOT: u16 = 3;

    let mut chunk = Chunk::new("<linear-f64-store-fill>");
    chunk.local_count = 4;
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(DST_BASE as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_f64_const(PATTERN, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::F64_STORE, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(8, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_i32_const((DST_BASE + TOTAL - 8) as i32, 0);
    chunk.emit_op(Op::F64_LOAD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("linear f64 store fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::F64(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear f64.store fill loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_fill, 1);
    assert_eq!(counters.memory_fill, 1);
    assert_eq!(counters.memory_fill_bytes, TOTAL as u64);
}

/// Runtime-only fixture for repeated halfword stores:
/// memory 0 <- I32_STORE16(value) with `i += 2`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i32_store16_fill_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const PATTERN: i32 = 0x1122_3344;
    const EXPECTED: i32 = 0x3344;

    const DST_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const VALUE_SLOT: u16 = 3;

    let mut chunk = Chunk::new("<linear-i32-store16-fill>");
    chunk.local_count = 4;
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(DST_BASE as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i32_const(PATTERN, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::I32_STORE16, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_i32_const((DST_BASE + TOTAL - 2) as i32, 0);
    chunk.emit_op(Op::I32_LOAD16_U, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("linear i32 store16 fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(EXPECTED));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i32.store16 fill loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_fill, 1);
    assert_eq!(counters.memory_fill, 1);
    assert_eq!(counters.memory_fill_bytes, TOTAL as u64);
}

/// Runtime-only fixture for repeated low-dword stores from i64:
/// memory 0 <- I64_STORE32(value) with `i += 4`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i64_store32_fill_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const PATTERN: i64 = 0x1020_3040_5060_7080;
    const EXPECTED: i64 = 0x5060_7080;

    const DST_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const VALUE_SLOT: u16 = 3;

    let mut chunk = Chunk::new("<linear-i64-store32-fill>");
    chunk.local_count = 4;
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(DST_BASE as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i64_const(PATTERN, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::I64_STORE32, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(4, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_i32_const((DST_BASE + TOTAL - 4) as i32, 0);
    chunk.emit_op(Op::I64_LOAD32_U, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("linear i64 store32 fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I64(EXPECTED));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i64.store32 fill loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_fill, 1);
    assert_eq!(counters.memory_fill, 1);
    assert_eq!(counters.memory_fill_bytes, TOTAL as u64);
}

/// Runtime-only fixture for word-sized linear-memory copy:
/// memory 0 -> I32_LOAD -> I32_STORE -> memory 0 with `i += 4`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i32_load_to_i32_store_copy_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;

    const INIT_I_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const COPY_I_SLOT: u16 = 3;
    const VALUE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-i32-copy-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op(Op::I32_STORE, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(4, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<linear-i32-copy>");
    copy.local_count = 5;
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);

    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::I32_LOAD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    copy.emit_op(Op::I32_STORE, 0);

    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_i32_const(4, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);
    copy.emit_br(0, 0);
    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);

    copy.emit_i32_const((DST_BASE + TOTAL - 4) as i32, 0);
    copy.emit_op(Op::I32_LOAD, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear i32 source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("linear i32 copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32((TOTAL - 4) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i32 copy: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
    assert_eq!(counters.super_linear_copy, 1);
    assert_eq!(counters.i32_load, (TOTAL / 4 + 1) as u64);
    assert_eq!(counters.i32_store, (TOTAL / 4) as u64);
}

/// Runtime-only fixture for word-sized linear-memory copy:
/// memory 0 -> I64_LOAD -> I64_STORE -> memory 0 with `i += 8`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i64_load_to_i64_store_copy_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;
    const PATTERN: i64 = 0x1020_3040_5060_7080;

    const INIT_I_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const COPY_I_SLOT: u16 = 3;
    const VALUE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-i64-copy-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_i64_const(PATTERN, 0);
    init.emit_op(Op::I64_STORE, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(8, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<linear-i64-copy>");
    copy.local_count = 5;
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);

    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::I64_LOAD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    copy.emit_op(Op::I64_STORE, 0);

    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_i32_const(8, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);
    copy.emit_br(0, 0);
    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);

    copy.emit_i32_const((DST_BASE + TOTAL - 8) as i32, 0);
    copy.emit_op(Op::I64_LOAD, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear i64 source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("linear i64 copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I64(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i64 copy: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
    assert_eq!(counters.super_linear_copy, 1);
    assert_eq!(counters.i64_load, (TOTAL / 8 + 1) as u64);
    assert_eq!(counters.i64_store, (TOTAL / 8) as u64);
}

/// Runtime-only fixture for float word linear-memory copy:
/// memory 0 -> F32_LOAD -> F32_STORE -> memory 0 with `i += 4`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_f32_load_to_f32_store_copy_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;
    const PATTERN: f32 = 123.5;

    const INIT_I_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const COPY_I_SLOT: u16 = 3;
    const VALUE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-f32-copy-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_f32_const(PATTERN, 0);
    init.emit_op(Op::F32_STORE, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(4, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<linear-f32-copy>");
    copy.local_count = 5;
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);
    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::F32_LOAD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    copy.emit_op(Op::F32_STORE, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_i32_const(4, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);
    copy.emit_br(0, 0);
    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);
    copy.emit_i32_const((DST_BASE + TOTAL - 4) as i32, 0);
    copy.emit_op(Op::F32_LOAD, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear f32 source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("linear f32 copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::F32(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear f32 copy: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_copy, 1);
    assert_eq!(counters.f32_load, (TOTAL / 4 + 1) as u64);
    assert_eq!(counters.f32_store, (TOTAL / 4) as u64);
}

/// Runtime-only fixture for double word linear-memory copy:
/// memory 0 -> F64_LOAD -> F64_STORE -> memory 0 with `i += 8`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_f64_load_to_f64_store_copy_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;
    const PATTERN: f64 = 456.25;

    const INIT_I_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const COPY_I_SLOT: u16 = 3;
    const VALUE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-f64-copy-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_f64_const(PATTERN, 0);
    init.emit_op(Op::F64_STORE, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(8, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<linear-f64-copy>");
    copy.local_count = 5;
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);
    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::F64_LOAD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    copy.emit_op(Op::F64_STORE, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_i32_const(8, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);
    copy.emit_br(0, 0);
    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);
    copy.emit_i32_const((DST_BASE + TOTAL - 8) as i32, 0);
    copy.emit_op(Op::F64_LOAD, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear f64 source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("linear f64 copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::F64(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear f64 copy: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_copy, 1);
    assert_eq!(counters.f64_load, (TOTAL / 8 + 1) as u64);
    assert_eq!(counters.f64_store, (TOTAL / 8) as u64);
}

/// Runtime-only fixture for packed signed halfword linear-memory copy:
/// memory 0 -> I32_LOAD16_S -> I32_STORE16 -> memory 0 with `i += 2`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i32_load16_s_to_i32_store16_copy_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;
    const PATTERN: i32 = -2;

    const INIT_I_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const COPY_I_SLOT: u16 = 3;
    const VALUE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-i32-load16-copy-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_i32_const(PATTERN, 0);
    init.emit_op(Op::I32_STORE16, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(2, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<linear-i32-load16-copy>");
    copy.local_count = 5;
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);
    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::I32_LOAD16_S, 0);
    copy.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    copy.emit_op(Op::I32_STORE16, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_i32_const(2, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);
    copy.emit_br(0, 0);
    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);
    copy.emit_i32_const((DST_BASE + TOTAL - 2) as i32, 0);
    copy.emit_op(Op::I32_LOAD16_S, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear i32 load16 source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("linear i32 load16 copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i32 load16/store16 copy: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_copy, 1);
    assert_eq!(counters.memory_copy, 1);
    assert_eq!(counters.memory_copy_bytes, TOTAL as u64);
}

/// Runtime-only fixture for packed unsigned dword linear-memory copy:
/// memory 0 -> I64_LOAD32_U -> I64_STORE32 -> memory 0 with `i += 4`.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_i64_load32_u_to_i64_store32_copy_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;
    const PATTERN: i64 = 0xffff_ff80;

    const INIT_I_SLOT: u16 = 0;
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const COPY_I_SLOT: u16 = 3;
    const VALUE_SLOT: u16 = 4;

    let mut init = Chunk::new("<linear-i64-load32-copy-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_i64_const(PATTERN, 0);
    init.emit_op(Op::I64_STORE32, 0);
    init.emit_op_u16(Op::LOCAL_GET, INIT_I_SLOT, 0);
    init.emit_i32_const(4, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, INIT_I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<linear-i64-load32-copy>");
    copy.local_count = 5;
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);
    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::I64_LOAD32_U, 0);
    copy.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    copy.emit_op(Op::I64_STORE32, 0);
    copy.emit_op_u16(Op::LOCAL_GET, COPY_I_SLOT, 0);
    copy.emit_i32_const(4, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, COPY_I_SLOT, 0);
    copy.emit_br(0, 0);
    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);
    copy.emit_i32_const((DST_BASE + TOTAL - 4) as i32, 0);
    copy.emit_op(Op::I64_LOAD32_U, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("linear i64 load32 source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("linear i64 load32 copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I64(PATTERN));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i64 load32/store32 copy: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_linear_copy, 1);
    assert_eq!(counters.memory_copy, 1);
    assert_eq!(counters.memory_copy_bytes, TOTAL as u64);
}

#[test]
fn linear_i32_offset_copy_loop_uses_superop() {
    const TOTAL: usize = 64;
    const SRC_BASE: usize = 4;
    const SRC_OFFSET: usize = 8;
    const DST_BASE: usize = 256;
    const DST_OFFSET: usize = 16;
    const FILL: i32 = 0x3c;

    const SRC_SLOT: u16 = 0;
    const SRC_OFFSET_SLOT: u16 = 1;
    const DST_SLOT: u16 = 2;
    const DST_OFFSET_SLOT: u16 = 3;
    const LEN_SLOT: u16 = 4;
    const I_SLOT: u16 = 5;
    const BYTE_SLOT: u16 = 6;

    let mut copy = Chunk::new("<linear-byte-offset-copy>");
    copy.local_count = 7;

    copy.emit_f64_const(1.0, 0);
    copy.emit_op(Op::MEMORY_GROW, 0);
    copy.emit_leb_u32(0, 0);
    copy.emit_op(Op::DROP, 0);

    copy.emit_i32_const((SRC_BASE + SRC_OFFSET) as i32, 0);
    copy.emit_i32_const(FILL, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op(Op::MEMORY_FILL, 0);
    copy.emit_leb_u32(0, 0);

    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    copy.emit_i32_const(SRC_OFFSET as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, SRC_OFFSET_SLOT, 0);
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    copy.emit_i32_const(DST_OFFSET as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, DST_OFFSET_SLOT, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    copy.emit_i32_const(0, 0);
    copy.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);

    let outer = copy.emit_block(0);
    let (lp, _loop_start) = copy.emit_loop_s(0);

    copy.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    copy.emit_op(Op::I32_GE_U, 0);
    copy.emit_br_if(1, 0);

    copy.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, SRC_OFFSET_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op(Op::I32_LOAD8_U, 0);
    copy.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    copy.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    copy.emit_op_u16(Op::LOCAL_GET, DST_OFFSET_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    copy.emit_op(Op::I32_STORE8, 0);

    copy.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    copy.emit_i32_const(1, 0);
    copy.emit_op(Op::I32_ADD, 0);
    copy.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    copy.emit_br(0, 0);

    copy.emit_end(0);
    copy.patch_loop(lp);
    copy.emit_end(0);
    copy.patch_block(outer);

    copy.emit_i32_const((DST_BASE + DST_OFFSET + TOTAL - 1) as i32, 0);
    copy.emit_op(Op::I32_LOAD8_U, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let result = vm.run(vec![copy]).expect("offset copy loop fixture failed");
    assert_eq!(result, Value::I32(FILL));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    assert_eq!(counters.super_linear_copy, 1);
}

/// Runtime-only fixture for bulk `memory.copy`.
///
/// This gives a VM-only baseline for the path that should already be much
/// cheaper than scalar `i32.load8_u -> i32.store8` loops.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn linear_memory_copy_bulk_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 32 * 1024;
    const I_SLOT: u16 = 0;

    let mut init = Chunk::new("<memory-copy-bulk-init>");
    init.local_count = 1;
    init.emit_f64_const(1.0, 0);
    init.emit_op(Op::MEMORY_GROW, 0);
    init.emit_leb_u32(0, 0);
    init.emit_op(Op::DROP, 0);
    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(0xff, 0);
    init.emit_op(Op::I32_AND, 0);
    init.emit_op(Op::I32_STORE8, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(1, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<memory-copy-bulk>");
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op(Op::MEMORY_COPY, 0);
    copy.emit_leb_u32(0, 0);
    copy.emit_leb_u32(0, 0);
    copy.emit_i32_const((DST_BASE + TOTAL - 1) as i32, 0);
    copy.emit_op(Op::I32_LOAD8_U, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init]).expect("memory.copy source init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![copy]).expect("memory.copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "memory.copy bulk: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
}

/// Runtime-only fixture for bulk `memory.copy` from memory 0 to memory 1.
///
/// This guards the source-memory0/destination-extra-memory helper branch.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn linear_memory_copy_memory0_to_extra_bulk_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const DST_BASE: usize = 0;
    const I_SLOT: u16 = 0;

    let mut init = Chunk::new("<memory-copy-memory0-to-extra-init>");
    init.memory_min_pages = vec![1, 1];
    init.memory_max_pages = vec![None, None];
    init.local_count = 1;

    init.emit_i32_const(0, 0);
    init.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    let outer = init.emit_block(0);
    let (lp, _loop_start) = init.emit_loop_s(0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(TOTAL as i32, 0);
    init.emit_op(Op::I32_GE_U, 0);
    init.emit_br_if(1, 0);
    init.emit_i32_const(SRC_BASE as i32, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(0xff, 0);
    init.emit_op(Op::I32_AND, 0);
    init.emit_op(Op::I32_STORE8, 0);
    init.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    init.emit_i32_const(1, 0);
    init.emit_op(Op::I32_ADD, 0);
    init.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    init.emit_br(0, 0);
    init.emit_end(0);
    init.patch_loop(lp);
    init.emit_end(0);
    init.patch_block(outer);
    init.emit_i32_const(0, 0);
    init.emit_op(Op::RETURN, 0);

    let mut copy = Chunk::new("<memory-copy-memory0-to-extra-bulk>");
    copy.emit_i32_const(DST_BASE as i32, 0);
    copy.emit_i32_const(SRC_BASE as i32, 0);
    copy.emit_i32_const(TOTAL as i32, 0);
    copy.emit_op(Op::MEMORY_COPY, 0);
    copy.emit_leb_u32(1, 0);
    copy.emit_leb_u32(0, 0);
    copy.emit_i32_const((DST_BASE + TOTAL - 1) as i32, 0);
    copy.emit_op(Op::I32_LOAD8_U, 0);
    copy.emit_leb_u32(0x80 | 0x40, 0);
    copy.emit_leb_u32(0, 0);
    copy.emit_leb_u32(1, 0);
    copy.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.run(vec![init])
        .expect("memory.copy memory0 to extra init failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![copy])
        .expect("memory.copy memory0 to extra fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "memory.copy memory0->extra bulk: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.memory_copy, 1);
    assert_eq!(counters.memory_copy_bytes, TOTAL as u64);
}

/// Runtime-only fixture for bulk `memory.fill`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn linear_memory_fill_bulk_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const FILL: i32 = 0x5a;

    let mut fill = Chunk::new("<memory-fill-bulk>");
    fill.emit_f64_const(1.0, 0);
    fill.emit_op(Op::MEMORY_GROW, 0);
    fill.emit_leb_u32(0, 0);
    fill.emit_op(Op::DROP, 0);
    fill.emit_i32_const(DST_BASE as i32, 0);
    fill.emit_i32_const(FILL, 0);
    fill.emit_i32_const(TOTAL as i32, 0);
    fill.emit_op(Op::MEMORY_FILL, 0);
    fill.emit_leb_u32(0, 0);
    fill.emit_i32_const((DST_BASE + TOTAL - 1) as i32, 0);
    fill.emit_op(Op::I32_LOAD8_U, 0);
    fill.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![fill]).expect("memory.fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(FILL));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "memory.fill bulk: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
}

/// Runtime-only fixture for scalar byte fill loops that should collapse to the
/// guarded linear-fill superinstruction.
///
/// Shape:
/// memory 0 <- repeated I32_STORE8 with loop-carried index.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn linear_i32_store8_fill_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const DST_BASE: usize = 0;
    const FILL: i32 = 0x6b;

    const DST_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const VALUE_SLOT: u16 = 3;

    let mut fill = Chunk::new("<linear-i32-store8-fill-loop>");
    fill.local_count = 4;

    fill.emit_f64_const(1.0, 0);
    fill.emit_op(Op::MEMORY_GROW, 0);
    fill.emit_leb_u32(0, 0);
    fill.emit_op(Op::DROP, 0);

    fill.emit_i32_const(DST_BASE as i32, 0);
    fill.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    fill.emit_i32_const(TOTAL as i32, 0);
    fill.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    fill.emit_i32_const(0, 0);
    fill.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    fill.emit_i32_const(FILL, 0);
    fill.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = fill.emit_block(0);
    let (lp, _loop_start) = fill.emit_loop_s(0);

    fill.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    fill.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    fill.emit_op(Op::I32_GE_U, 0);
    fill.emit_br_if(1, 0);

    fill.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    fill.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    fill.emit_op(Op::I32_ADD, 0);
    fill.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    fill.emit_op(Op::I32_STORE8, 0);

    fill.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    fill.emit_i32_const(1, 0);
    fill.emit_op(Op::I32_ADD, 0);
    fill.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    fill.emit_br(0, 0);

    fill.emit_end(0);
    fill.patch_loop(lp);
    fill.emit_end(0);
    fill.patch_block(outer);

    fill.emit_i32_const((DST_BASE + TOTAL - 1) as i32, 0);
    fill.emit_op(Op::I32_LOAD8_U, 0);
    fill.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![fill]).expect("linear store8 fill fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(FILL));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i32.store8 fill loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
}

#[test]
fn linear_i32_store8_offset_fill_loop_uses_superop() {
    const TOTAL: usize = 64;
    const DST_BASE: usize = 4;
    const DST_OFFSET: usize = 8;
    const FILL: i32 = 0x5a;

    const DST_SLOT: u16 = 0;
    const OFFSET_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const I_SLOT: u16 = 3;
    const VALUE_SLOT: u16 = 4;

    let mut fill = Chunk::new("<linear-i32-store8-offset-fill-loop>");
    fill.local_count = 5;

    fill.emit_f64_const(1.0, 0);
    fill.emit_op(Op::MEMORY_GROW, 0);
    fill.emit_leb_u32(0, 0);
    fill.emit_op(Op::DROP, 0);

    fill.emit_i32_const(DST_BASE as i32, 0);
    fill.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    fill.emit_i32_const(DST_OFFSET as i32, 0);
    fill.emit_op_u16(Op::LOCAL_SET, OFFSET_SLOT, 0);
    fill.emit_i32_const(TOTAL as i32, 0);
    fill.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    fill.emit_i32_const(0, 0);
    fill.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    fill.emit_i32_const(FILL, 0);
    fill.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let outer = fill.emit_block(0);
    let (lp, _loop_start) = fill.emit_loop_s(0);

    fill.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    fill.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    fill.emit_op(Op::I32_GE_U, 0);
    fill.emit_br_if(1, 0);

    fill.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    fill.emit_op_u16(Op::LOCAL_GET, OFFSET_SLOT, 0);
    fill.emit_op(Op::I32_ADD, 0);
    fill.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    fill.emit_op(Op::I32_ADD, 0);
    fill.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    fill.emit_op(Op::I32_STORE8, 0);

    fill.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    fill.emit_i32_const(1, 0);
    fill.emit_op(Op::I32_ADD, 0);
    fill.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    fill.emit_br(0, 0);

    fill.emit_end(0);
    fill.patch_loop(lp);
    fill.emit_end(0);
    fill.patch_block(outer);

    fill.emit_i32_const((DST_BASE + DST_OFFSET + TOTAL - 1) as i32, 0);
    fill.emit_op(Op::I32_LOAD8_U, 0);
    fill.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let result = vm.run(vec![fill]).expect("offset fill loop fixture failed");
    assert_eq!(result, Value::I32(FILL));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    assert_eq!(counters.super_linear_fill, 1);
    assert_eq!(counters.memory_fill_bytes, TOTAL as u64);
}

#[test]
fn linear_i32_load8_offset_scan_loop_uses_superop() {
    const TOTAL: usize = 64;
    const SRC_BASE: usize = 4;
    const SRC_OFFSET: usize = 8;
    const FILL: i32 = 7;

    const SRC_SLOT: u16 = 0;
    const OFFSET_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const I_SLOT: u16 = 3;
    const ACC_SLOT: u16 = 4;

    let mut scan = Chunk::new("<linear-i32-load8-offset-scan-loop>");
    scan.local_count = 5;

    scan.emit_f64_const(1.0, 0);
    scan.emit_op(Op::MEMORY_GROW, 0);
    scan.emit_leb_u32(0, 0);
    scan.emit_op(Op::DROP, 0);

    scan.emit_i32_const((SRC_BASE + SRC_OFFSET) as i32, 0);
    scan.emit_i32_const(FILL, 0);
    scan.emit_i32_const(TOTAL as i32, 0);
    scan.emit_op(Op::MEMORY_FILL, 0);
    scan.emit_leb_u32(0, 0);

    scan.emit_i32_const(SRC_BASE as i32, 0);
    scan.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    scan.emit_i32_const(SRC_OFFSET as i32, 0);
    scan.emit_op_u16(Op::LOCAL_SET, OFFSET_SLOT, 0);
    scan.emit_i32_const(TOTAL as i32, 0);
    scan.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    scan.emit_i32_const(0, 0);
    scan.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    scan.emit_i32_const(0, 0);
    scan.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    let outer = scan.emit_block(0);
    let (lp, _loop_start) = scan.emit_loop_s(0);

    scan.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    scan.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    scan.emit_op(Op::I32_GE_U, 0);
    scan.emit_br_if(1, 0);

    scan.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    scan.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    scan.emit_op_u16(Op::LOCAL_GET, OFFSET_SLOT, 0);
    scan.emit_op(Op::I32_ADD, 0);
    scan.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    scan.emit_op(Op::I32_ADD, 0);
    scan.emit_op(Op::I32_LOAD8_U, 0);
    scan.emit_op(Op::I32_ADD, 0);
    scan.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    scan.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    scan.emit_i32_const(1, 0);
    scan.emit_op(Op::I32_ADD, 0);
    scan.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    scan.emit_br(0, 0);

    scan.emit_end(0);
    scan.patch_loop(lp);
    scan.emit_end(0);
    scan.patch_block(outer);

    scan.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    scan.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let result = vm.run(vec![scan]).expect("offset scan loop fixture failed");
    assert_eq!(result, Value::I32((TOTAL as i32) * FILL));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    assert_eq!(counters.super_linear_scan, 1);
    assert_eq!(counters.i32_load8_u, TOTAL as u64);
}

/// Runtime-only fixture for scalar `i32.load8_u` scan loops.
///
/// Initializes memory with one bulk fill, then scans byte-by-byte into an i32
/// accumulator. This gives a VM-only view of scalar load, integer arithmetic,
/// local traffic, and branch/backedge overhead.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn linear_i32_load8_scan_loop_counters() {
    const TOTAL: usize = 32 * 1024;
    const SRC_BASE: usize = 0;
    const FILL: i32 = 7;

    const SRC_SLOT: u16 = 0;
    const LEN_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const ACC_SLOT: u16 = 3;

    let mut scan = Chunk::new("<linear-i32-load8-scan-loop>");
    scan.local_count = 4;

    scan.emit_f64_const(1.0, 0);
    scan.emit_op(Op::MEMORY_GROW, 0);
    scan.emit_leb_u32(0, 0);
    scan.emit_op(Op::DROP, 0);

    scan.emit_i32_const(SRC_BASE as i32, 0);
    scan.emit_i32_const(FILL, 0);
    scan.emit_i32_const(TOTAL as i32, 0);
    scan.emit_op(Op::MEMORY_FILL, 0);
    scan.emit_leb_u32(0, 0);

    scan.emit_i32_const(SRC_BASE as i32, 0);
    scan.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    scan.emit_i32_const(TOTAL as i32, 0);
    scan.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    scan.emit_i32_const(0, 0);
    scan.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    scan.emit_i32_const(0, 0);
    scan.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    let outer = scan.emit_block(0);
    let (lp, _loop_start) = scan.emit_loop_s(0);

    scan.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    scan.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    scan.emit_op(Op::I32_GE_U, 0);
    scan.emit_br_if(1, 0);

    scan.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    scan.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    scan.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    scan.emit_op(Op::I32_ADD, 0);
    scan.emit_op(Op::I32_LOAD8_U, 0);
    scan.emit_op(Op::I32_ADD, 0);
    scan.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    scan.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    scan.emit_i32_const(1, 0);
    scan.emit_op(Op::I32_ADD, 0);
    scan.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    scan.emit_br(0, 0);

    scan.emit_end(0);
    scan.patch_loop(lp);
    scan.emit_end(0);
    scan.patch_block(outer);

    scan.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    scan.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![scan]).expect("linear load8 scan fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32((TOTAL as i32) * FILL));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "linear i32.load8_u scan loop: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
}

/// Runtime-only fixture for hot bytecode function calls through `ref.func` /
/// `call_ref`.
///
/// This isolates function-reference materialization, call dispatch, frame
/// setup/teardown, and return-result restoration from compiler and loader cost.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn call_ref_direct_function_counters() {
    const FUNC_SLOT: u16 = 0;
    const ACC_SLOT: u16 = 1;

    let mut main = Chunk::new("<call-ref-direct-main>");
    main.local_count = 2;

    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0, 0);
    main.emit_op_u16(Op::LOCAL_SET, FUNC_SLOT, 0);

    main.emit_i32_const(0, 0);
    main.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    emit_structured_counter_loop(&mut main, |c| {
        c.emit_op_u16(Op::LOCAL_GET, FUNC_SLOT, 0);
        c.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
        c.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
        c.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);
    });

    main.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    main.emit_op(Op::RETURN, 0);

    let mut add_one = Chunk::new("<call-ref-add-one>");
    add_one.arity = 1;
    add_one.local_count = 1;
    add_one.emit_op_u16(Op::LOCAL_GET, 0, 0);
    add_one.emit_i32_const(1, 0);
    add_one.emit_op(Op::I32_ADD, 0);
    add_one.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![main, add_one])
        .expect("call_ref direct fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(ITERATIONS as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "call_ref direct function: {} calls in {:.2} ms ({:.1} ns/call)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_call_ref_i32_add_loop, 1);
}

/// Runtime-only fixture for direct bytecode calls emitted as spec `call` and
/// resolved through the normal VM import table to a linked chunk function.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn call_direct_chunk_function_counters() {
    const ACC_SLOT: u16 = 0;

    let mut main = Chunk::new("<call-direct-main>");
    main.local_count = 1;
    main.emit_i32_const(0, 0);
    main.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    emit_structured_counter_loop(&mut main, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
        c.emit_call(0, 1, 0);
        c.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);
    });

    main.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    main.emit_op(Op::RETURN, 0);

    let mut add_one = Chunk::new("<call-direct-add-one>");
    add_one.arity = 1;
    add_one.param_count = 1;
    add_one.result_arity = 1;
    add_one.local_count = 1;
    add_one.emit_op_u16(Op::LOCAL_GET, 0, 0);
    add_one.emit_i32_const(1, 0);
    add_one.emit_op(Op::I32_ADD, 0);
    add_one.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run_linked(
            vec![main, add_one],
            vec![ImportTarget::ChunkFn {
                chunk_index: 1,
                arity: 1,
            }],
        )
        .expect("direct call chunk fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(ITERATIONS as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "direct call chunk function: {} calls in {:.2} ms ({:.1} ns/call)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_call_i32_add_loop, 1);
}

/// Runtime-only fixture for repeated module global reads into a stable local.
///
/// The optimized shape is `global.get idx; local.set dst; countdown`. The
/// chunk carries a normal module global table so `VM::merge_global_table`
/// rewrites the operand to the authoritative VM global slot before execution.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn global_get_local_loop_counters() {
    const GLOBAL_NAME: &str = "__perf_global_get_loop_value";
    const VALUE_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<global-get-local-loop>");
    chunk.local_count = 1;
    chunk.globals = std::sync::Arc::new(vec![GLOBAL_NAME.to_string()]);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u32(Op::GLOBAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.set_global(GLOBAL_NAME, Value::I32(314));
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("global_get local fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(314));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "global.get local loop: {} reads in {:.2} ms ({:.1} ns/read)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_global_get_local_loop, 1);
}

/// Runtime-only fixture for repeated local writes from an i32 constant.
///
/// The optimized shape is `i32.const value; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_set_const_loop_counters() {
    const VALUE_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<local-set-const-loop>");
    chunk.local_count = 1;

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_i32_const(1234, 0);
        c.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("local_set const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(1234));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local.set const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_set_const_loop, 1);
}

/// Runtime-only fixture for repeated local reads whose values are discarded
/// immediately.
///
/// The optimized shape is `local.get src; drop; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_get_drop_loop_counters() {
    const VALUE_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<local-get-drop-loop>");
    chunk.local_count = 1;
    chunk.emit_i32_const(4321, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
        c.emit_op(Op::DROP, 0);
    });

    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("local_get drop fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(1));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local.get drop loop: {} reads in {:.2} ms ({:.1} ns/read)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_get_drop_loop, 1);
}

/// Runtime-only fixture for repeated local-to-local copies.
///
/// The optimized shape is `local.get src; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_copy_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-copy-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(2468, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("local_copy fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(2468));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local copy loop: {} copies in {:.2} ms ({:.1} ns/copy)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_copy_loop, 1);
}

/// Runtime-only fixture for repeated stable local plus constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.add; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_add_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-add-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(2468, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(7, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_add_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(2475));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 add const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_add_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local minus constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.sub; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_sub_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-sub-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(2468, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(7, 0);
        c.emit_op(Op::I32_SUB, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_sub_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(2461));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 sub const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_sub_const_loop, 1);
}

fn run_local_i32_accum_const_loop_fixture(
    label: &str,
    initial: i32,
    rhs: i32,
    op: Op,
    expected: i32,
) {
    const SLOT: u16 = 0;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 1;
    chunk.emit_i32_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SLOT, 0);
        c.emit_i32_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_accum_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "{}: {} accumulations in {:.2} ms ({:.1} ns/iter)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_accum_const_loop, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_accum_add_const_loop_counters() {
    let expected = 5i32.wrapping_add(7i32.wrapping_mul(ITERATIONS as i32));
    run_local_i32_accum_const_loop_fixture(
        "<local-i32-accum-add-const-loop>",
        5,
        7,
        Op::I32_ADD,
        expected,
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_accum_sub_const_loop_counters() {
    let expected = i32::MIN.wrapping_sub(3i32.wrapping_mul(ITERATIONS as i32));
    run_local_i32_accum_const_loop_fixture(
        "<local-i32-accum-sub-const-loop>",
        i32::MIN,
        3,
        Op::I32_SUB,
        expected,
    );
}

/// Runtime-only fixture for repeated stable local bitwise-and constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.and; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_and_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-and-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(0x7f3d, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(0x00ff, 0);
        c.emit_op(Op::I32_AND, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_and_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0x003d));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 and const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_and_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local bitwise-or constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.or; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_or_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-or-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(0x700d, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(0x00f0, 0);
        c.emit_op(Op::I32_OR, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_or_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0x70fd));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 or const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_or_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local multiply constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.mul; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_mul_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-mul-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(2468, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(7, 0);
        c.emit_op(Op::I32_MUL, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_mul_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(17276));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 mul const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_mul_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local shift-left constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.shl; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_shl_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-shl-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(0x1234, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(3, 0);
        c.emit_op(Op::I32_SHL, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_shl_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0x91a0));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 shl const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_shl_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local signed shift-right constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.shr_s; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_shr_s_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-shr-s-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(-1024, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(3, 0);
        c.emit_op(Op::I32_SHR_S, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_shr_s_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(-128));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 shr_s const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_shr_s_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local unsigned shift-right constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.shr_u; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_shr_u_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-shr-u-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(-1024, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(3, 0);
        c.emit_op(Op::I32_SHR_U, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_shr_u_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0x1fffff80));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 shr_u const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_shr_u_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local rotate-left constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.rotl; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_rotl_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-rotl-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(0x12345678, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(8, 0);
        c.emit_op(Op::I32_ROTL, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_rotl_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0x34567812));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 rotl const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_rotl_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local rotate-right constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.rotr; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_rotr_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-rotr-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(0x12345678, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(8, 0);
        c.emit_op(Op::I32_ROTR, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_rotr_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0x78123456));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 rotr const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_rotr_const_loop, 1);
}

/// Runtime-only fixture for repeated stable local bitwise-xor constant writes.
///
/// The optimized shape is `local.get src; i32.const k; i32.xor; local.set dst; countdown`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_xor_const_loop_counters() {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<local-i32-xor-const-loop>");
    chunk.local_count = 2;
    chunk.emit_i32_const(0x7f3d, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(0x00ff, 0);
        c.emit_op(Op::I32_XOR, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_xor_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0x7fc2));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "local i32 xor const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i32_xor_const_loop, 1);
}

fn run_local_i32_compare_const_loop_fixture(
    label: &str,
    initial: i32,
    rhs: i32,
    op: Op,
    expected: i32,
    counter_name: &str,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    chunk.emit_i32_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i32_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i32_compare_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains(&format!("{counter_name}=1")),
        "expected {counter_name}=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_eq_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-eq-const-loop>",
        1234,
        1234,
        Op::I32_EQ,
        1,
        "super_local_i32_eq_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_ne_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-ne-const-loop>",
        1234,
        7,
        Op::I32_NE,
        1,
        "super_local_i32_ne_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_lt_s_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-lt-s-const-loop>",
        -3,
        2,
        Op::I32_LT_S,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_lt_u_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-lt-u-const-loop>",
        3,
        -1,
        Op::I32_LT_U,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_gt_s_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-gt-s-const-loop>",
        7,
        -2,
        Op::I32_GT_S,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_gt_u_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-gt-u-const-loop>",
        -1,
        7,
        Op::I32_GT_U,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_le_s_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-le-s-const-loop>",
        -2,
        -2,
        Op::I32_LE_S,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_le_u_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-le-u-const-loop>",
        7,
        -1,
        Op::I32_LE_U,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_ge_s_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-ge-s-const-loop>",
        4,
        4,
        Op::I32_GE_S,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_ge_u_const_loop_counters() {
    run_local_i32_compare_const_loop_fixture(
        "<local-i32-ge-u-const-loop>",
        -1,
        4,
        Op::I32_GE_U,
        1,
        "super_local_i32_ord_cmp_const_loop",
    );
}

fn run_local_eqz_loop_fixture(
    label: &str,
    initial: Value,
    eqz_op: Op,
    expected: Value,
    counter_name: &str,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    match initial {
        Value::I32(value) => chunk.emit_i32_const(value, 0),
        Value::I64(value) => chunk.emit_i64_const(value, 0),
        other => panic!("unsupported eqz fixture initial value: {other:?}"),
    }
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_op(eqz_op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("local_eqz_loop fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, expected);

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains(&format!("{counter_name}=1")),
        "expected {counter_name}=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i32_eqz_loop_counters() {
    run_local_eqz_loop_fixture(
        "<local-i32-eqz-loop>",
        Value::I32(0),
        Op::I32_EQZ,
        Value::I32(1),
        "super_local_i32_eqz_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_eqz_loop_counters() {
    run_local_eqz_loop_fixture(
        "<local-i64-eqz-loop>",
        Value::I64(7),
        Op::I64_EQZ,
        Value::I32(0),
        "super_local_i64_eqz_loop",
    );
}

fn run_local_i64_compare_const_loop_fixture(
    label: &str,
    initial: i64,
    rhs: i64,
    op: Op,
    expected: i32,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    chunk.emit_i64_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i64_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i64_compare_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains("super_local_i64_cmp_const_loop=1"),
        "expected super_local_i64_cmp_const_loop=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_eq_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-eq-const-loop>", 1234, 1234, Op::I64_EQ, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_ne_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-ne-const-loop>", 1234, 7, Op::I64_NE, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_lt_s_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-lt-s-const-loop>", -3, 2, Op::I64_LT_S, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_lt_u_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-lt-u-const-loop>", 3, -1, Op::I64_LT_U, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_gt_s_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-gt-s-const-loop>", 7, -2, Op::I64_GT_S, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_gt_u_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-gt-u-const-loop>", -1, 7, Op::I64_GT_U, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_le_s_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-le-s-const-loop>", -2, -2, Op::I64_LE_S, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_le_u_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-le-u-const-loop>", 7, -1, Op::I64_LE_U, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_ge_s_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-ge-s-const-loop>", 4, 4, Op::I64_GE_S, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_ge_u_const_loop_counters() {
    run_local_i64_compare_const_loop_fixture("<local-i64-ge-u-const-loop>", -1, 4, Op::I64_GE_U, 1);
}

fn run_local_f32_compare_const_loop_fixture(
    label: &str,
    initial: f32,
    rhs: f32,
    op: Op,
    expected: i32,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    chunk.emit_f32_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_f32_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_f32_compare_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains("super_local_f32_cmp_const_loop=1"),
        "expected super_local_f32_cmp_const_loop=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_eq_const_loop_counters() {
    run_local_f32_compare_const_loop_fixture("<local-f32-eq-const-loop>", 1.5, 1.5, Op::F32_EQ, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_ne_nan_const_loop_counters() {
    run_local_f32_compare_const_loop_fixture("<local-f32-ne-nan-const-loop>", f32::NAN, 1.0, Op::F32_NE, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_lt_const_loop_counters() {
    run_local_f32_compare_const_loop_fixture("<local-f32-lt-const-loop>", -1.0, 2.0, Op::F32_LT, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_gt_const_loop_counters() {
    run_local_f32_compare_const_loop_fixture("<local-f32-gt-const-loop>", 3.5, 2.0, Op::F32_GT, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_le_const_loop_counters() {
    run_local_f32_compare_const_loop_fixture("<local-f32-le-const-loop>", 2.0, 2.0, Op::F32_LE, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_ge_const_loop_counters() {
    run_local_f32_compare_const_loop_fixture("<local-f32-ge-const-loop>", 3.0, 2.0, Op::F32_GE, 1);
}

fn run_local_f64_compare_const_loop_fixture(
    label: &str,
    initial: f64,
    rhs: f64,
    op: Op,
    expected: i32,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    chunk.emit_f64_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_f64_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_f64_compare_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains("super_local_f64_cmp_const_loop=1"),
        "expected super_local_f64_cmp_const_loop=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_eq_const_loop_counters() {
    run_local_f64_compare_const_loop_fixture("<local-f64-eq-const-loop>", 1.5, 1.5, Op::F64_EQ, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_ne_nan_const_loop_counters() {
    run_local_f64_compare_const_loop_fixture("<local-f64-ne-nan-const-loop>", f64::NAN, 1.0, Op::F64_NE, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_lt_const_loop_counters() {
    run_local_f64_compare_const_loop_fixture("<local-f64-lt-const-loop>", -1.0, 2.0, Op::F64_LT, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_gt_const_loop_counters() {
    run_local_f64_compare_const_loop_fixture("<local-f64-gt-const-loop>", 3.5, 2.0, Op::F64_GT, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_le_const_loop_counters() {
    run_local_f64_compare_const_loop_fixture("<local-f64-le-const-loop>", 2.0, 2.0, Op::F64_LE, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_ge_const_loop_counters() {
    run_local_f64_compare_const_loop_fixture("<local-f64-ge-const-loop>", 3.0, 2.0, Op::F64_GE, 1);
}

fn run_local_f32_binary_const_loop_fixture(
    label: &str,
    initial: f32,
    rhs: f32,
    op: Op,
    expected: f32,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    chunk.emit_f32_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_f32_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_f32_binary_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::F32(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains("super_local_f32_binary_const_loop=1"),
        "expected super_local_f32_binary_const_loop=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_add_const_loop_counters() {
    run_local_f32_binary_const_loop_fixture("<local-f32-add-const-loop>", 1.25, 2.5, Op::F32_ADD, 3.75);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_sub_const_loop_counters() {
    run_local_f32_binary_const_loop_fixture("<local-f32-sub-const-loop>", 5.5, 2.25, Op::F32_SUB, 3.25);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_mul_const_loop_counters() {
    run_local_f32_binary_const_loop_fixture("<local-f32-mul-const-loop>", 1.5, 4.0, Op::F32_MUL, 6.0);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f32_div_const_loop_counters() {
    run_local_f32_binary_const_loop_fixture("<local-f32-div-const-loop>", 7.5, 3.0, Op::F32_DIV, 2.5);
}

fn run_local_f64_binary_const_loop_fixture(
    label: &str,
    initial: f64,
    rhs: f64,
    op: Op,
    expected: f64,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    chunk.emit_f64_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_f64_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_f64_binary_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::F64(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains("super_local_f64_binary_const_loop=1"),
        "expected super_local_f64_binary_const_loop=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_add_const_loop_counters() {
    run_local_f64_binary_const_loop_fixture("<local-f64-add-const-loop>", 1.25, 2.5, Op::F64_ADD, 3.75);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_sub_const_loop_counters() {
    run_local_f64_binary_const_loop_fixture("<local-f64-sub-const-loop>", 5.5, 2.25, Op::F64_SUB, 3.25);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_mul_const_loop_counters() {
    run_local_f64_binary_const_loop_fixture("<local-f64-mul-const-loop>", 1.5, 4.0, Op::F64_MUL, 6.0);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_f64_div_const_loop_counters() {
    run_local_f64_binary_const_loop_fixture("<local-f64-div-const-loop>", 7.5, 3.0, Op::F64_DIV, 2.5);
}

fn run_local_i64_binary_const_loop_fixture(
    label: &str,
    initial: i64,
    rhs: i64,
    op: Op,
    expected: i64,
    counter_name: &str,
) {
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 2;
    chunk.emit_i64_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
        c.emit_i64_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("local_i64_binary_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I64(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    let report = counters.format_report();
    println!(
        "{}: {} writes in {:.2} ms ({:.1} ns/write)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", report);
    assert!(
        report.contains(&format!("{counter_name}=1")),
        "expected {counter_name}=1 in perf report"
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_add_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-add-const-loop>",
        4_000_000_000,
        17,
        Op::I64_ADD,
        4_000_000_017,
        "super_local_i64_add_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_sub_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-sub-const-loop>",
        4_000_000_000,
        17,
        Op::I64_SUB,
        3_999_999_983,
        "super_local_i64_sub_const_loop",
    );
}

fn run_local_i64_accum_const_loop_fixture(
    label: &str,
    initial: i64,
    rhs: i64,
    op: Op,
    expected: i64,
) {
    const SLOT: u16 = 0;

    let mut chunk = Chunk::new(label);
    chunk.local_count = 1;
    chunk.emit_i64_const(initial, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, SLOT, 0);
        c.emit_i64_const(rhs, 0);
        c.emit_op(op, 0);
        c.emit_op_u16(Op::LOCAL_SET, SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("local_i64_accum_const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I64(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "{}: {} accumulations in {:.2} ms ({:.1} ns/iter)",
        label,
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_local_i64_accum_const_loop, 1);
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_accum_add_const_loop_counters() {
    let expected = 5i64.wrapping_add(7i64.wrapping_mul(ITERATIONS as i64));
    run_local_i64_accum_const_loop_fixture(
        "<local-i64-accum-add-const-loop>",
        5,
        7,
        Op::I64_ADD,
        expected,
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_accum_sub_const_loop_counters() {
    let expected = i64::MIN.wrapping_sub(3i64.wrapping_mul(ITERATIONS as i64));
    run_local_i64_accum_const_loop_fixture(
        "<local-i64-accum-sub-const-loop>",
        i64::MIN,
        3,
        Op::I64_SUB,
        expected,
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_and_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-and-const-loop>",
        0x7fff_ffff_ffff_1234,
        0xffff,
        Op::I64_AND,
        0x1234,
        "super_local_i64_and_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_mul_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-mul-const-loop>",
        4_000_000_000,
        3,
        Op::I64_MUL,
        12_000_000_000,
        "super_local_i64_mul_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_or_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-or-const-loop>",
        0x7000_0000_0000_000d,
        0xf0,
        Op::I64_OR,
        0x7000_0000_0000_00fd,
        "super_local_i64_or_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_shl_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-shl-const-loop>",
        0x1234,
        4,
        Op::I64_SHL,
        0x12340,
        "super_local_i64_shl_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_shr_s_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-shr-s-const-loop>",
        -1024,
        3,
        Op::I64_SHR_S,
        -128,
        "super_local_i64_shr_s_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_shr_u_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-shr-u-const-loop>",
        -1024,
        3,
        Op::I64_SHR_U,
        0x1fff_ffff_ffff_ff80,
        "super_local_i64_shr_u_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_rotl_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-rotl-const-loop>",
        0x1234_5678_9abc_def0,
        8,
        Op::I64_ROTL,
        0x3456_789a_bcde_f012,
        "super_local_i64_rotl_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_rotr_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-rotr-const-loop>",
        0x1234_5678_9abc_def0,
        8,
        Op::I64_ROTR,
        0xf012_3456_789a_bcde_u64 as i64,
        "super_local_i64_rotr_const_loop",
    );
}

#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn local_i64_xor_const_loop_counters() {
    run_local_i64_binary_const_loop_fixture(
        "<local-i64-xor-const-loop>",
        0x7fff_ffff_ffff_1234,
        0xffff,
        Op::I64_XOR,
        0x7fff_ffff_ffff_edcb,
        "super_local_i64_xor_const_loop",
    );
}

/// Runtime-only fixture for repeated module global reads whose values are
/// discarded immediately.
///
/// The optimized shape is `global.get idx; drop; countdown`. The chunk carries
/// a normal module global table so `VM::merge_global_table` rewrites the
/// operand to the authoritative VM global slot before execution.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn global_get_drop_loop_counters() {
    const GLOBAL_NAME: &str = "__perf_global_get_drop_loop_value";

    let mut chunk = Chunk::new("<global-get-drop-loop>");
    chunk.globals = std::sync::Arc::new(vec![GLOBAL_NAME.to_string()]);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u32(Op::GLOBAL_GET, 0, 0);
        c.emit_op(Op::DROP, 0);
    });

    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.set_global(GLOBAL_NAME, Value::I32(314));
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("global_get drop fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(1));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "global.get drop loop: {} reads in {:.2} ms ({:.1} ns/read)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_global_get_drop_loop, 1);
}

/// Runtime-only fixture for repeated module global writes from a stable local.
///
/// The optimized shape is `local.get src; global.set idx; countdown`. The
/// chunk carries a normal module global table so `VM::merge_global_table`
/// rewrites the operand to the authoritative VM global slot before execution.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn global_set_local_loop_counters() {
    const GLOBAL_NAME: &str = "__perf_global_set_loop_value";
    const VALUE_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<global-set-local-loop>");
    chunk.local_count = 1;
    chunk.globals = std::sync::Arc::new(vec![GLOBAL_NAME.to_string()]);

    chunk.emit_i32_const(271, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
        c.emit_op_u32(Op::GLOBAL_SET, 0, 0);
    });

    chunk.emit_op_u32(Op::GLOBAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("global_set local fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(271));
    assert_eq!(vm.global(GLOBAL_NAME), Some(&Value::I32(271)));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "global.set local loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_global_set_local_loop, 1);
}

/// Runtime-only fixture for repeated module global writes from an i32 constant.
///
/// The optimized shape is `i32.const value; global.set idx; countdown`. The
/// chunk carries a normal module global table so `VM::merge_global_table`
/// rewrites the operand to the authoritative VM global slot before execution.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn global_set_const_loop_counters() {
    const GLOBAL_NAME: &str = "__perf_global_set_const_loop_value";

    let mut chunk = Chunk::new("<global-set-const-loop>");
    chunk.globals = std::sync::Arc::new(vec![GLOBAL_NAME.to_string()]);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_i32_const(919, 0);
        c.emit_op_u32(Op::GLOBAL_SET, 0, 0);
    });

    chunk.emit_op_u32(Op::GLOBAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("global_set const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(919));
    assert_eq!(vm.global(GLOBAL_NAME), Some(&Value::I32(919)));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "global.set const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_global_set_const_loop, 1);
}

/// Runtime-only fixture for hot bytecode function calls through a wasm table
/// and `call_indirect`.
///
/// The optimized measured loop is `local.get acc; i32.const 0; call_indirect
/// 1, table 0, 1; local.set acc; countdown`. Setup populates table 0 through
/// ordinary `ref.func`/`table.set` bytecode before counters are enabled.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn call_indirect_direct_function_counters() {
    const ACC_SLOT: u16 = 0;

    let mut setup = Chunk::new("<call-indirect-setup>");
    setup.emit_i32_const(0, 0);
    setup.emit_op_u16(Op::REF_FUNC, 1, 0);
    setup.emit(0, 0);
    setup.emit_op(Op::TABLE_SET, 0);
    setup.emit_leb_u32(0, 0);
    setup.emit_i32_const(0, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut main = Chunk::new("<call-indirect-direct-main>");
    main.local_count = 1;
    main.emit_i32_const(0, 0);
    main.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    emit_structured_counter_loop(&mut main, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::CALL_INDIRECT, 0);
        c.emit(1, 0);
        c.emit(0, 0);
        c.emit(1, 0);
        c.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);
    });

    main.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    main.emit_op(Op::RETURN, 0);

    let mut add_one = Chunk::new("<call-indirect-add-one>");
    add_one.arity = 1;
    add_one.param_count = 1;
    add_one.result_arity = 1;
    add_one.local_count = 1;
    add_one.emit_op_u16(Op::LOCAL_GET, 0, 0);
    add_one.emit_i32_const(1, 0);
    add_one.emit_op(Op::I32_ADD, 0);
    add_one.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.wasm_tables.push(vec![Value::Null]);
    vm.run(vec![setup, add_one.clone()])
        .expect("call_indirect setup failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![main, add_one])
        .expect("call_indirect direct fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(ITERATIONS as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "call_indirect direct function: {} calls in {:.2} ms ({:.1} ns/call)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_call_indirect_i32_add_loop, 1);
}

/// Same indirect-call hot loop as `call_indirect_direct_function_counters`, but
/// with the table element index held in a stable local. This matches function
/// pointer/table-dispatch lowering where the callee index is loaded once and
/// reused across the loop.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn call_indirect_local_index_direct_function_counters() {
    const ACC_SLOT: u16 = 0;
    const INDEX_SLOT: u16 = 1;

    let mut setup = Chunk::new("<call-indirect-local-index-setup>");
    setup.emit_i32_const(0, 0);
    setup.emit_op_u16(Op::REF_FUNC, 1, 0);
    setup.emit(0, 0);
    setup.emit_op(Op::TABLE_SET, 0);
    setup.emit_leb_u32(0, 0);
    setup.emit_i32_const(0, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut main = Chunk::new("<call-indirect-local-index-main>");
    main.local_count = 2;
    main.emit_i32_const(0, 0);
    main.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);
    main.emit_i32_const(0, 0);
    main.emit_op_u16(Op::LOCAL_SET, INDEX_SLOT, 0);

    emit_structured_counter_loop(&mut main, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
        c.emit_op_u16(Op::LOCAL_GET, INDEX_SLOT, 0);
        c.emit_op(Op::CALL_INDIRECT, 0);
        c.emit(1, 0);
        c.emit(0, 0);
        c.emit(1, 0);
        c.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);
    });

    main.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    main.emit_op(Op::RETURN, 0);

    let mut add_one = Chunk::new("<call-indirect-local-index-add-one>");
    add_one.arity = 1;
    add_one.param_count = 1;
    add_one.result_arity = 1;
    add_one.local_count = 1;
    add_one.emit_op_u16(Op::LOCAL_GET, 0, 0);
    add_one.emit_i32_const(1, 0);
    add_one.emit_op(Op::I32_ADD, 0);
    add_one.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.wasm_tables.push(vec![Value::Null]);
    vm.run(vec![setup, add_one.clone()])
        .expect("call_indirect local-index setup failed");
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![main, add_one])
        .expect("call_indirect local-index fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(ITERATIONS as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "call_indirect local-index direct function: {} calls in {:.2} ms ({:.1} ns/call)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_call_indirect_i32_add_loop, 1);
}

/// Runtime-only fixture for repeated plain own-property reads.
///
/// The optimized shape is `local.get obj; struct.get name; drop` inside the
/// structured counter loop. It must only fire for own data properties; getters,
/// prototype reads, typed-field fallback, nulls, and task `.result` stay on the
/// generic `STRUCT_GET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn struct_get_drop_loop_counters() {
    const OBJ_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<struct-get-drop-loop>");
    chunk.local_count = 1;

    chunk.emit_struct_new(0, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, OBJ_SLOT, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
    chunk.emit_i32_const(99, 0);
    let field_name = chunk.add_constant(Value::String("x".into()));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, field_name, 0);
    chunk.emit_op(Op::DROP, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
        let fk = c.add_constant(Value::String("x".into()));
        c.emit_struct_field_op(Op::STRUCT_GET, 0, fk, 0);
        c.emit_op(Op::DROP, 0);
    });

    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("struct_get drop fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(1));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "struct.get drop loop: {} reads in {:.2} ms ({:.1} ns/read)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_struct_get_drop_loop, 1);
}

/// Runtime-only fixture for repeated dynamic object property reads into a
/// stable local.
///
/// The optimized shape is `local.get obj; struct.get 0 name; local.set dst`
/// inside the structured counter loop. It only fires for own data properties:
/// task `.result`, getters, prototype lookup, typed fallback, and invalid
/// receivers stay on the generic `STRUCT_GET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn struct_get_local_loop_counters() {
    const OBJ_SLOT: u16 = 0;
    const VALUE_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<struct-get-local-loop>");
    chunk.local_count = 2;

    chunk.emit_struct_new(0, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, OBJ_SLOT, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
    chunk.emit_i32_const(77, 0);
    let field_name = chunk.add_constant(Value::String("x".into()));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, field_name, 0);
    chunk.emit_op(Op::DROP, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
        c.emit_struct_field_op(Op::STRUCT_GET, 0, field_name, 0);
        c.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("struct_get local fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(77));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "struct.get local loop: {} reads in {:.2} ms ({:.1} ns/read)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_struct_get_local_loop, 1);
}

/// Runtime-only fixture for repeated dynamic object property writes from a
/// stable local.
///
/// The optimized shape is `local.get obj; local.get value; struct.set 0 name`
/// inside the structured counter loop. It only fires for untyped object
/// property bags with no setter; typed-field synchronization and setter calls
/// stay on the generic `STRUCT_SET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn struct_set_local_loop_counters() {
    const OBJ_SLOT: u16 = 0;
    const VALUE_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<struct-set-local-loop>");
    chunk.local_count = 2;

    chunk.emit_struct_new(0, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, OBJ_SLOT, 0);
    chunk.emit_i32_const(123, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    let field_name = chunk.add_constant(Value::String("x".into()));
    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
        c.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
        c.emit_struct_field_op(Op::STRUCT_SET, 0, field_name, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, field_name, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("struct_set local fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(123));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "struct.set local loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_struct_set_local_loop, 1);
}

/// Runtime-only fixture for repeated dynamic object property writes of the
/// same i32 constant.
///
/// The optimized shape is `local.get obj; i32.const value; struct.set 0 name`
/// inside the structured counter loop. It only fires for untyped object
/// property bags with no setter; typed-field synchronization and setter calls
/// stay on the generic `STRUCT_SET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn struct_set_const_loop_counters() {
    const OBJ_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<struct-set-const-loop>");
    chunk.local_count = 1;

    chunk.emit_struct_new(0, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, OBJ_SLOT, 0);

    let field_name = chunk.add_constant(Value::String("x".into()));
    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
        c.emit_i32_const(321, 0);
        c.emit_struct_field_op(Op::STRUCT_SET, 0, field_name, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, OBJ_SLOT, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, field_name, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("struct_set const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(321));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "struct.set const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_struct_set_const_loop, 1);
}

/// Runtime-only fixture for repeated dense array reads whose value is dropped.
///
/// The optimized shape is `local.get arr; i32.const index; array.get; drop`
/// inside the structured counter loop. It only fires for in-bounds dense array
/// reads; typed arrays, maps, dynamic property fallback, and out-of-bounds GC
/// traps stay on the generic `ARRAY_GET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn array_get_drop_loop_counters() {
    const ARR_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<array-get-drop-loop>");
    chunk.local_count = 1;

    chunk.emit_i32_const(7, 0);
    chunk.emit_array_new_fixed(0, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
        c.emit_op(Op::DROP, 0);
    });

    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("array_get drop fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(1));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array.get drop loop: {} reads in {:.2} ms ({:.1} ns/read)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_array_get_drop_loop, 1);
}

/// Runtime-only fixture for repeated dense array reads into a stable local.
///
/// The optimized shape is `local.get arr; i32.const index; array.get;
/// local.set dst` inside the structured counter loop. It only fires for
/// existing dense slots; typed arrays, maps, dynamic property lookup, negative
/// indexes, and out-of-bounds GC traps stay on the generic `ARRAY_GET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn array_get_local_loop_counters() {
    const ARR_SLOT: u16 = 0;
    const VALUE_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<array-get-local-loop>");
    chunk.local_count = 2;

    chunk.emit_i32_const(7, 0);
    chunk.emit_array_new_fixed(0, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
        c.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("array_get local fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(7));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array.get local loop: {} reads in {:.2} ms ({:.1} ns/read)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_array_get_local_loop, 1);
}

/// Runtime-only fixture for repeated dense array writes of the same i32 value.
///
/// The optimized shape is `local.get arr; i32.const index; i32.const value;
/// array.set` inside the structured counter loop. It only fires for existing
/// dense slots; growth, typed arrays, maps, `__setitem__`, and out-of-bounds GC
/// traps stay on the generic `ARRAY_SET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn array_set_const_loop_counters() {
    const ARR_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<array-set-const-loop>");
    chunk.local_count = 1;

    chunk.emit_i32_const(0, 0);
    chunk.emit_array_new_fixed(0, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
        c.emit_i32_const(0, 0);
        c.emit_i32_const(9, 0);
        c.emit_op(Op::ARRAY_SET, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("array_set const fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(9));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array.set const loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_array_set_const_loop, 1);
}

/// Runtime-only fixture for repeated dense array writes from a stable local.
///
/// The optimized shape is `local.get arr; i32.const index; local.get value;
/// array.set` inside the structured counter loop. It keeps dynamic growth,
/// typed arrays, maps, `__setitem__`, and out-of-bounds GC traps on the generic
/// `ARRAY_SET` path.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn array_set_local_loop_counters() {
    const ARR_SLOT: u16 = 0;
    const VALUE_SLOT: u16 = 1;

    let mut chunk = Chunk::new("<array-set-local-loop>");
    chunk.local_count = 2;

    chunk.emit_i32_const(0, 0);
    chunk.emit_array_new_fixed(0, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);
    chunk.emit_i32_const(11, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, VALUE_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::LOCAL_GET, VALUE_SLOT, 0);
        c.emit_op(Op::ARRAY_SET, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("array_set local fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(11));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array.set local loop: {} writes in {:.2} ms ({:.1} ns/write)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_array_set_local_loop, 1);
}

/// Runtime-only fixture for a managed-array to managed-array copy loop that was
/// emitted as `ARRAY_GET` + `ARRAY_SET` instead of `ARRAY_COPY`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn array_copy_loop_counters() {
    const LEN: usize = 4096;
    const SRC_GLOBAL: &str = "__perf_array_copy_loop_source";
    const DST_GLOBAL: &str = "__perf_array_copy_loop_destination";
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const LEN_SLOT: u16 = 3;

    let mut setup_src = Chunk::new("<array-copy-loop-source-setup>");
    for i in 0..LEN {
        setup_src.emit_i32_const((i & 0xff) as i32, 0);
    }
    setup_src.emit_array_new_fixed(0, LEN as u16, 0);
    setup_src.emit_op(Op::RETURN, 0);

    let mut setup_dst = Chunk::new("<array-copy-loop-destination-setup>");
    for _ in 0..LEN {
        setup_dst.emit_i32_const(0, 0);
    }
    setup_dst.emit_array_new_fixed(0, LEN as u16, 0);
    setup_dst.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let source = vm
        .run(vec![setup_src])
        .expect("array copy source setup failed");
    let destination = vm
        .run(vec![setup_dst])
        .expect("array copy destination setup failed");
    vm.set_global(SRC_GLOBAL, source);
    vm.set_global(DST_GLOBAL, destination);

    let mut chunk = Chunk::new("<array-copy-loop>");
    chunk.local_count = 4;

    let src_global_idx = chunk.add_constant(Value::String(SRC_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, src_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    let dst_global_idx = chunk.add_constant(Value::String(DST_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, dst_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i32_const(LEN as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::ARRAY_SET, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_i32_const((LEN - 1) as i32, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("array copy loop fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((LEN - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array copy loop: {} elements in {:.2} ms ({:.1} ns/element)",
        LEN,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / LEN as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_array_copy_loop, 1);
}

/// Runtime-only fixture for direct `ARRAY_COPY`.
///
/// This covers the generic opcode path separately from the `ARRAY_GET` +
/// `ARRAY_SET` loop superinstruction fixtures above.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn array_copy_direct_opcode_counters() {
    const LEN: usize = 4096;
    const SRC_GLOBAL: &str = "__perf_array_copy_direct_source";
    const DST_GLOBAL: &str = "__perf_array_copy_direct_destination";

    let mut setup_src = Chunk::new("<array-copy-direct-source-setup>");
    for i in 0..LEN {
        setup_src.emit_i32_const((i & 0xff) as i32, 0);
    }
    setup_src.emit_array_new_fixed(0, LEN as u16, 0);
    setup_src.emit_op(Op::RETURN, 0);

    let mut setup_dst = Chunk::new("<array-copy-direct-destination-setup>");
    for _ in 0..LEN {
        setup_dst.emit_i32_const(0, 0);
    }
    setup_dst.emit_array_new_fixed(0, LEN as u16, 0);
    setup_dst.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let source = vm
        .run(vec![setup_src])
        .expect("array copy direct source setup failed");
    let destination = vm
        .run(vec![setup_dst])
        .expect("array copy direct destination setup failed");
    vm.set_global(SRC_GLOBAL, source);
    vm.set_global(DST_GLOBAL, destination);

    let mut chunk = Chunk::new("<array-copy-direct>");
    let dst_global_idx = chunk.add_constant(Value::String(DST_GLOBAL.into()));
    let src_global_idx = chunk.add_constant(Value::String(SRC_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, dst_global_idx, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u32(Op::GLOBAL_GET, src_global_idx, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(LEN as i32, 0);
    chunk.emit_op(Op::ARRAY_COPY, 0);
    chunk.emit_op_u32(Op::GLOBAL_GET, dst_global_idx, 0);
    chunk.emit_i32_const((LEN - 1) as i32, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("array copy direct fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((LEN - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array.copy direct: {} elements in {:.2} ms ({:.1} ns/element)",
        LEN,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / LEN as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.array_copy, 1);
}

/// Runtime-only fixture for direct `ARRAY_FILL`.
///
/// This covers the generic opcode path separately from any loop-lowered
/// element writes. `ARRAY_FILL` has no dedicated perf counter today, so the
/// fixture validates the filled value and prints the normal runtime counters.
#[test]
#[ignore = "runtime perf fixture - invoke with --ignored --nocapture"]
fn array_fill_direct_opcode_counters() {
    const LEN: usize = 4096;
    const ARRAY_GLOBAL: &str = "__perf_array_fill_direct_target";

    let mut setup = Chunk::new("<array-fill-direct-setup>");
    for _ in 0..LEN {
        setup.emit_i32_const(0, 0);
    }
    setup.emit_array_new_fixed(0, LEN as u16, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let target = vm.run(vec![setup]).expect("array fill direct setup failed");
    vm.set_global(ARRAY_GLOBAL, target);

    let mut chunk = Chunk::new("<array-fill-direct>");
    let target_global_idx = chunk.add_constant(Value::String(ARRAY_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, target_global_idx, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(7, 0);
    chunk.emit_i32_const(LEN as i32, 0);
    chunk.emit_op(Op::ARRAY_FILL, 0);
    chunk.emit_op_u32(Op::GLOBAL_GET, target_global_idx, 0);
    chunk.emit_i32_const((LEN - 1) as i32, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("array fill direct fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(7));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array.fill direct: {} elements in {:.2} ms ({:.1} ns/element)",
        LEN,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / LEN as f64
    );
    println!("{}", counters.format_report());
}

/// Runtime-only fixture for offset managed-array copy loops:
/// `dst[dst_offset + i] = src[src_offset + i]`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn array_copy_offset_loop_counters() {
    const SRC_OFFSET: usize = 7;
    const DST_OFFSET: usize = 13;
    const LEN: usize = 4096;
    const SRC_LEN: usize = SRC_OFFSET + LEN;
    const DST_LEN: usize = DST_OFFSET + LEN;
    const SRC_GLOBAL: &str = "__perf_array_copy_offset_loop_source";
    const DST_GLOBAL: &str = "__perf_array_copy_offset_loop_destination";
    const SRC_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const I_SLOT: u16 = 2;
    const LEN_SLOT: u16 = 3;
    const SRC_OFFSET_SLOT: u16 = 4;
    const DST_OFFSET_SLOT: u16 = 5;

    let mut setup_src = Chunk::new("<array-copy-offset-loop-source-setup>");
    for i in 0..SRC_LEN {
        setup_src.emit_i32_const((i & 0xff) as i32, 0);
    }
    setup_src.emit_array_new_fixed(0, SRC_LEN as u16, 0);
    setup_src.emit_op(Op::RETURN, 0);

    let mut setup_dst = Chunk::new("<array-copy-offset-loop-destination-setup>");
    for _ in 0..DST_LEN {
        setup_dst.emit_i32_const(0, 0);
    }
    setup_dst.emit_array_new_fixed(0, DST_LEN as u16, 0);
    setup_dst.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let source = vm
        .run(vec![setup_src])
        .expect("array copy offset source setup failed");
    let destination = vm
        .run(vec![setup_dst])
        .expect("array copy offset destination setup failed");
    vm.set_global(SRC_GLOBAL, source);
    vm.set_global(DST_GLOBAL, destination);

    let mut chunk = Chunk::new("<array-copy-offset-loop>");
    chunk.local_count = 6;

    let src_global_idx = chunk.add_constant(Value::String(SRC_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, src_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_SLOT, 0);
    let dst_global_idx = chunk.add_constant(Value::String(DST_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, dst_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i32_const(LEN as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(SRC_OFFSET as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_OFFSET_SLOT, 0);
    chunk.emit_i32_const(DST_OFFSET as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_OFFSET_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_OFFSET_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, SRC_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, SRC_OFFSET_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::ARRAY_SET, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, DST_OFFSET_SLOT, 0);
    chunk.emit_i32_const((LEN - 1) as i32, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm
        .run(vec![chunk])
        .expect("array copy offset loop fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((SRC_OFFSET + LEN - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "array copy offset loop: {} elements in {:.2} ms ({:.1} ns/element)",
        LEN,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / LEN as f64
    );
    println!("{}", counters.format_report());
    assert_eq!(counters.super_array_copy_loop, 1);
}

/// Runtime-only fixture for core i32 arithmetic dispatch.
///
/// This isolates integer opcode dispatch, typed stack pops, wrapping arithmetic,
/// local get/set traffic, and loop backedge overhead from compiler and loader
/// cost.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn i32_arithmetic_hot_loop_counters() {
    const ACC_SLOT: u16 = 0;

    let mut chunk = Chunk::new("<i32-arithmetic-hot-loop>");
    chunk.local_count = 1;
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);

    emit_structured_counter_loop(&mut chunk, |c| {
        c.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
        c.emit_i32_const(1_664_525, 0);
        c.emit_op(Op::I32_MUL, 0);
        c.emit_i32_const(1_013_904_223, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, ACC_SLOT, 0);
    });

    chunk.emit_op_u16(Op::LOCAL_GET, ACC_SLOT, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut expected = 0i32;
    for _ in 0..ITERATIONS {
        expected = expected
            .wrapping_mul(1_664_525)
            .wrapping_add(1_013_904_223);
    }

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("i32 arithmetic fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(expected));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "i32 arithmetic hot loop: {} iterations in {:.2} ms ({:.1} ns/iter)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
}

/// Runtime-only fixture for plain structured branch/backedge overhead.
///
/// The loop body is intentionally empty; the measured work is the counter
/// local traffic, decrement, `br_if`, and `br 0`/loop label machinery emitted
/// by `emit_structured_counter_loop`.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn branch_backedge_hot_loop_counters() {
    let mut chunk = Chunk::new("<branch-backedge-hot-loop>");
    emit_structured_counter_loop(&mut chunk, |_c| {});
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("branch fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "branch/backedge hot loop: {} iterations in {:.2} ms ({:.1} ns/iter)",
        ITERATIONS,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
}

/// Runtime-only fixture for hot `br_table` dispatch.
///
/// This keeps the branch target semantics simple: each table entry exits one
/// inner block, then the surrounding structured counter loop continues. The
/// useful work is the repeated selection from a 64-entry table, which used to
/// re-decode all LEB depths on every execution.
#[test]
#[ignore = "runtime perf fixture — invoke with --ignored --nocapture"]
fn br_table_hot_loop_counters() {
    const TABLE_LEN: usize = 64;
    let mut chunk = Chunk::new("<br-table-hot-loop>");
    let depths = vec![0u32; TABLE_LEN];
    emit_structured_counter_loop(&mut chunk, |c| {
        let inner = c.emit_block(0);
        c.emit_i32_const(17, 0);
        c.emit_br_table(&depths, 0, 0);
        c.emit_end(0);
        c.patch_block(inner);
    });
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("br_table fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(0));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "br_table hot loop: {} iterations, {} entries in {:.2} ms ({:.1} ns/iter)",
        ITERATIONS,
        TABLE_LEN,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / ITERATIONS as f64
    );
    println!("{}", counters.format_report());
}
